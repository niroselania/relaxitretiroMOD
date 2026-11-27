import hmac
import io
import json
import os
import secrets
import sqlite3
from datetime import datetime, timedelta
from functools import wraps

import barcode
from barcode.writer import SVGWriter
from markupsafe import Markup, escape
from flask import (Flask, abort, flash, g, make_response, redirect, render_template,
                   request, send_file, session, url_for, jsonify)
from openpyxl import Workbook, load_workbook
from werkzeug.security import check_password_hash, generate_password_hash


app = Flask(__name__)
app.config.update(
    SECRET_KEY=os.environ.get("SECRET_KEY", ""),
    DATABASE_PATH=os.environ.get("DATABASE_PATH", "data/deposito.sqlite3"),
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("SESSION_COOKIE_SECURE", "false").lower() == "true",
    MAX_CONTENT_LENGTH=12 * 1024 * 1024,
)

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
  id INTEGER PRIMARY KEY, username TEXT NOT NULL UNIQUE, password_hash TEXT NOT NULL,
  role TEXT NOT NULL DEFAULT 'consulta' CHECK(role IN ('admin','consulta'))
);
CREATE TABLE IF NOT EXISTS containers (
  id INTEGER PRIMARY KEY, code TEXT NOT NULL UNIQUE, name TEXT NOT NULL,
  location TEXT NOT NULL DEFAULT '', site TEXT NOT NULL DEFAULT 'Bariloche', active INTEGER NOT NULL DEFAULT 1,
  parent_id INTEGER REFERENCES containers(id),
  is_virtual INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS products (
  id INTEGER PRIMARY KEY, sku TEXT NOT NULL UNIQUE, name TEXT NOT NULL,
  item TEXT NOT NULL DEFAULT '', color TEXT NOT NULL DEFAULT '', size TEXT NOT NULL DEFAULT '',
  description TEXT NOT NULL DEFAULT '', color_name TEXT NOT NULL DEFAULT '',
  season TEXT NOT NULL DEFAULT '', retail_price REAL, legacy_units INTEGER NOT NULL DEFAULT 0,
  virtual_units INTEGER NOT NULL DEFAULT 0,
  source_location TEXT NOT NULL DEFAULT '', active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS stock (
  container_id INTEGER NOT NULL REFERENCES containers(id),
  product_id INTEGER NOT NULL REFERENCES products(id), quantity INTEGER NOT NULL CHECK(quantity >= 0),
  PRIMARY KEY(container_id, product_id)
);
CREATE TABLE IF NOT EXISTS movements (
  id INTEGER PRIMARY KEY, created_at TEXT NOT NULL, user_id INTEGER NOT NULL REFERENCES users(id),
  from_container_id INTEGER REFERENCES containers(id), to_container_id INTEGER REFERENCES containers(id),
  from_source TEXT NOT NULL DEFAULT 'external', product_id INTEGER NOT NULL REFERENCES products(id), quantity INTEGER NOT NULL CHECK(quantity > 0),
  note TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_movements_created ON movements(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_stock_product ON stock(product_id);
"""


def now_text():
    return datetime.now().astimezone().isoformat(timespec="seconds")


def get_db():
    if "db" not in g:
        os.makedirs(os.path.dirname(app.config["DATABASE_PATH"]) or ".", exist_ok=True)
        g.db = sqlite3.connect(app.config["DATABASE_PATH"])
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
        g.db.execute("PRAGMA journal_mode = WAL")
    return g.db


@app.teardown_appcontext
def close_db(_error=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    db = get_db()
    db.executescript(SCHEMA)
    # Keep existing local databases usable when new product and site fields are added.
    for table, column, declaration in [
        ("users", "role", "TEXT NOT NULL DEFAULT 'consulta'"),
        ("users", "site", "TEXT NOT NULL DEFAULT ''"),
        ("containers", "site", "TEXT NOT NULL DEFAULT 'Bariloche'"),
        ("containers", "parent_id", "INTEGER REFERENCES containers(id)"),
        ("containers", "is_virtual", "INTEGER NOT NULL DEFAULT 0"),
        ("containers", "from_stock_sheet", "INTEGER NOT NULL DEFAULT 0"),
        ("products", "item", "TEXT NOT NULL DEFAULT ''"),
        ("products", "color", "TEXT NOT NULL DEFAULT ''"),
        ("products", "size", "TEXT NOT NULL DEFAULT ''"),
        ("products", "color_name", "TEXT NOT NULL DEFAULT ''"),
        ("products", "season", "TEXT NOT NULL DEFAULT ''"),
        ("products", "retail_price", "REAL"),
        ("products", "legacy_units", "INTEGER NOT NULL DEFAULT 0"),
        ("products", "virtual_units", "INTEGER NOT NULL DEFAULT 0"),
        ("products", "source_location", "TEXT NOT NULL DEFAULT ''"),
        ("products", "active", "INTEGER NOT NULL DEFAULT 1"),
        ("movements", "from_source", "TEXT NOT NULL DEFAULT 'external'"),
    ]:
        columns = {row[1] for row in db.execute(f"PRAGMA table_info({table})")}
        if column not in columns:
            db.execute(f"ALTER TABLE {table} ADD COLUMN {column} {declaration}")
    # Racks creados por la planilla de stock (versiones previas los marcaban solo con la ubicación).
    db.execute("UPDATE containers SET from_stock_sheet=1 WHERE from_stock_sheet=0 AND is_virtual=0 "
               "AND location='Planilla de stock' AND name LIKE 'RACK %'")
    # Normalize the two fields used by the size finder so its composite index can serve searches.
    db.execute("UPDATE products SET item=TRIM(item), color=UPPER(TRIM(color)) "
               "WHERE item != TRIM(item) OR color != UPPER(TRIM(color))")
    db.execute("CREATE INDEX IF NOT EXISTS idx_products_finder ON products(item, color, active)")
    username = os.environ.get("ADMIN_USERNAME", "RETIRO").strip()
    password = os.environ.get("ADMIN_PASSWORD", "")
    if not app.config["SECRET_KEY"] or not password:
        raise RuntimeError("Configurá SECRET_KEY y ADMIN_PASSWORD antes de iniciar el sistema.")
    # Cada usuario trabaja con un depósito: el antiguo "admin" pasa a ser el usuario RETIRO.
    if not db.execute("SELECT 1 FROM users WHERE UPPER(username)='RETIRO'").fetchone():
        db.execute("UPDATE users SET username='RETIRO', site='Retiro' WHERE LOWER(username)='admin'")
    db.execute("UPDATE users SET site='Retiro' WHERE UPPER(username)='RETIRO' AND site=''")
    db.execute("UPDATE users SET site='Bariloche' WHERE UPPER(username)='BARILOCHE' AND site=''")
    db.execute("UPDATE users SET site='Retiro' WHERE site NOT IN ('Retiro','Bariloche','Todos')")
    if username.lower() == "admin":
        username = "RETIRO"
    # El usuario inicial solo se crea en una instalación nueva (o si no queda ningún administrador).
    if not db.execute("SELECT 1 FROM users WHERE role='admin'").fetchone() and \
            not db.execute("SELECT 1 FROM users WHERE username = ?", (username,)).fetchone():
        bootstrap_site = {"RETIRO": "Retiro", "BARILOCHE": "Bariloche"}.get(username.upper(), "Todos")
        db.execute("INSERT INTO users(username, password_hash, role, site) VALUES (?, ?, 'admin', ?)",
                   (username, generate_password_hash(password), bootstrap_site))
    db.execute("UPDATE users SET role='admin' WHERE username=?", (username,))
    # Keep the installation recoverable if ADMIN_USERNAME changed after setup.
    # The oldest account is the original owner account in existing installations.
    db.execute("UPDATE users SET role='admin' WHERE id=(SELECT MIN(id) FROM users) "
               "AND NOT EXISTS (SELECT 1 FROM users WHERE role='admin')")
    if not db.execute("SELECT 1 FROM containers WHERE is_virtual=1 AND active=1").fetchone():
        last = db.execute("SELECT MAX(CAST(code AS INTEGER)) FROM containers WHERE code GLOB '[0-9]*' AND length(code)=12").fetchone()[0]
        candidate = (int(last) + 1) if last else 1
        while db.execute("SELECT 1 FROM containers WHERE code=?", (f"{candidate:012d}",)).fetchone():
            candidate += 1
        code = f"{candidate:012d}"
        db.execute("INSERT INTO containers(code,name,location,site,active,parent_id,is_virtual,created_at) VALUES(?,?,?,?,1,NULL,1,?)",
                   (code, "DEPOSITO VIRTUAL", "Stock lógico sin inventario físico", "Retiro", now_text()))
    db.commit()


with app.app_context():
    init_db()


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if session.get("role") != "admin":
            abort(403, "Esta acción requiere permisos de administrador.")
        return view(*args, **kwargs)
    return wrapped


def admin_for_writes(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if request.method != "GET" and session.get("role") != "admin":
            abort(403, "El perfil de consulta no puede modificar información.")
        return view(*args, **kwargs)
    return wrapped


SITES = ("Bariloche", "Retiro")


def user_site():
    """Depósito del usuario conectado; None si ve todos (perfil "Todos")."""
    site = session.get("site", "")
    return site if site in SITES else None


def visible_sql(alias="c"):
    """Contenedores que ve el usuario: los de su depósito más los compartidos
    (Depósito Virtual y racks de la planilla de stock)."""
    site = user_site()
    if site is None:
        return "1=1", ()
    return f"({alias}.is_virtual=1 OR {alias}.from_stock_sheet=1 OR {alias}.site=?)", (site,)


def own_sql(alias="c", strict=True):
    """Racks propios del depósito del usuario (sin compartidos). Para un usuario "Todos",
    strict=False no filtra nada (historial, tablero y exportación completos)."""
    site = user_site()
    if site is None:
        return (f"({alias}.is_virtual=0 AND {alias}.from_stock_sheet=0)", ()) if strict else ("1=1", ())
    return f"({alias}.is_virtual=0 AND {alias}.from_stock_sheet=0 AND {alias}.site=?)", (site,)


def can_see(container):
    site = user_site()
    return bool(container) and (site is None or container["is_virtual"] or container["from_stock_sheet"]
                                or container["site"] == site)


def visible_container(db, container_id):
    container = db.execute("SELECT * FROM containers WHERE id=? AND active=1", (container_id,)).fetchone()
    if not can_see(container):
        abort(404)
    return container


@app.before_request
def protect_forms():
    if request.method == "POST":
        sent = request.form.get("csrf_token", "")
        saved = session.get("csrf_token", "")
        if not sent or not saved or not hmac.compare_digest(sent, saved):
            abort(400, "La sesión del formulario venció. Volvé a cargar la pantalla.")


@app.before_request
def refresh_session_role():
    # Roles live in SQLite; refresh the cookie on every request so a role
    # migration or admin recovery takes effect without requiring a new login.
    user_id = session.get("user_id")
    if user_id:
        user = get_db().execute("SELECT username, role, site FROM users WHERE id=?", (user_id,)).fetchone()
        if user:
            session["username"] = user["username"]
            session["role"] = user["role"]
            session["site"] = user["site"]
        else:
            session.clear()


@app.context_processor
def inject_template_values():
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_urlsafe(24)
    return {"csrf_token": session["csrf_token"], "active_path": request.path,
            "current_role": session.get("role"), "user_site": user_site()}


@app.template_filter("qty")
def quantity_format(value):
    return f"{value:,}".replace(",", ".")


@app.template_filter("date_short")
def date_short(value):
    try:
        return datetime.fromisoformat(value).strftime("%d/%m/%Y %H:%M")
    except (TypeError, ValueError):
        return value or "—"


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = get_db().execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        if user and check_password_hash(user["password_hash"], password):
            session.clear()
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]
            session["site"] = user["site"]
            session["csrf_token"] = secrets.token_urlsafe(24)
            return redirect(url_for("dashboard"))
        flash("Usuario o contraseña incorrectos.", "error")
    return render_template("login.html")


@app.post("/logout")
@login_required
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/cuenta", methods=["GET", "POST"])
@login_required
def account():
    if request.method == "POST":
        current = request.form.get("current_password", "")
        new = request.form.get("new_password", "")
        confirm = request.form.get("confirm_password", "")
        user = get_db().execute("SELECT * FROM users WHERE id=?", (session["user_id"],)).fetchone()
        if not check_password_hash(user["password_hash"], current):
            flash("La contraseña actual no es correcta.", "error")
        elif len(new) < 12:
            flash("La nueva contraseña debe tener al menos 12 caracteres.", "error")
        elif new != confirm:
            flash("La confirmación no coincide.", "error")
        else:
            db = get_db()
            db.execute("UPDATE users SET password_hash=? WHERE id=?", (generate_password_hash(new), session["user_id"]))
            db.commit()
            flash("Contraseña actualizada.", "success")
            return redirect(url_for("account"))
    return render_template("account.html")


@app.route("/usuarios", methods=["GET", "POST"])
@login_required
@admin_required
def users_admin():
    db = get_db()
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        role = request.form.get("role", "consulta")
        site = request.form.get("site", "")
        if request.form.get("action") == "site":
            user_id = request.form.get("user_id", type=int)
            if site not in (*SITES, "Todos") or not user_id:
                flash("Elegí un depósito válido.", "error")
            else:
                db.execute("UPDATE users SET site=? WHERE id=?", (site, user_id))
                db.commit()
                flash("Depósito del usuario actualizado.", "success")
            return redirect(url_for("users_admin"))
        if site not in (*SITES, "Todos"):
            flash("Elegí el depósito del usuario.", "error")
        elif len(username) < 3 or len(username) > 40:
            flash("El usuario debe tener entre 3 y 40 caracteres.", "error")
        elif len(password) < 12:
            flash("La contraseña debe tener al menos 12 caracteres.", "error")
        elif role not in ("admin", "consulta"):
            flash("Seleccioná un perfil válido.", "error")
        else:
            try:
                db.execute("INSERT INTO users(username,password_hash,role,site) VALUES(?,?,?,?)",
                           (username, generate_password_hash(password), role, site))
                db.commit()
                flash(f"Usuario {username} creado con perfil {role} para {site}.", "success")
                return redirect(url_for("users_admin"))
            except sqlite3.IntegrityError:
                flash("Ese nombre de usuario ya existe.", "error")
    users = db.execute("SELECT id,username,role,site FROM users ORDER BY username COLLATE NOCASE").fetchall()
    return render_template("users.html", users=users)


@app.route("/")
@login_required
def dashboard():
    db = get_db()
    own, own_args = own_sql("c")
    own_f, own_f_args = own_sql("f", strict=False)
    own_t, own_t_args = own_sql("t", strict=False)
    move_filter = f"({own_f} OR {own_t})"
    move_args = (*own_f_args, *own_t_args)
    stats = {
        "containers": db.execute(f"SELECT COUNT(*) FROM containers c WHERE c.active=1 AND {own}", own_args).fetchone()[0],
        "products": db.execute("SELECT COUNT(*) FROM products").fetchone()[0],
        "units": db.execute(f"SELECT COALESCE(SUM(s.quantity),0) FROM stock s JOIN containers c ON c.id=s.container_id WHERE c.active=1 AND {own}", own_args).fetchone()[0],
        "movements": db.execute(f"""SELECT COUNT(*) FROM movements m LEFT JOIN containers f ON f.id=m.from_container_id
          LEFT JOIN containers t ON t.id=m.to_container_id WHERE date(m.created_at)=date('now','localtime') AND {move_filter}""", move_args).fetchone()[0],
    }
    recent = db.execute("""SELECT m.*, p.sku, p.name product,
      CASE WHEN f.name IS NOT NULL THEN f.name WHEN m.from_source='legacy' THEN 'Planilla sin ubicación' WHEN m.from_source='container' THEN 'Retiro manual' ELSE 'Ingreso externo' END from_name,
      t.name to_name
      FROM movements m JOIN products p ON p.id=m.product_id
      LEFT JOIN containers f ON f.id=m.from_container_id
      LEFT JOIN containers t ON t.id=m.to_container_id
      WHERE {move_filter}
      ORDER BY m.id DESC LIMIT 8""".replace("{move_filter}", move_filter), move_args).fetchall()
    week_rows = {
        row[0]: row[1]
        for row in db.execute(
            "SELECT date(m.created_at), COUNT(*) FROM movements m LEFT JOIN containers f ON f.id=m.from_container_id "
            "LEFT JOIN containers t ON t.id=m.to_container_id "
            f"WHERE date(m.created_at) >= date('now','localtime','-6 days') AND {move_filter} GROUP BY 1", move_args
        )
    }
    today = datetime.now().astimezone().date()
    labels = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
    week = []
    peak = 1
    for offset in range(6, -1, -1):
        day = today - timedelta(days=offset)
        count = int(week_rows.get(day.isoformat(), 0) or 0)
        peak = max(peak, count)
        week.append({"label": labels[day.weekday()], "count": count})
    for item in week:
        item["pct"] = round(100 * item["count"] / peak) if peak else 0
    return render_template("dashboard.html", stats=stats, recent=recent, week=week)


def next_container_code(db):
    # Numeric Code 128 identifiers remain stable and are never recycled.
    last = db.execute("SELECT MAX(CAST(code AS INTEGER)) FROM containers WHERE code GLOB '[0-9]*' AND length(code)=12").fetchone()[0]
    candidate = (int(last) + 1) if last else 1
    while db.execute("SELECT 1 FROM containers WHERE code=?", (f"{candidate:012d}",)).fetchone():
        candidate += 1
    return f"{candidate:012d}"


@app.route("/contenedores")
@login_required
def containers():
    q = request.args.get("q", "").strip()
    visible, visible_args = visible_sql("c")
    own, own_args = own_sql("c")
    rows = get_db().execute(f"""SELECT c.*, COALESCE(SUM(s.quantity),0) units,
      COUNT(DISTINCT s.product_id) products_count
      FROM containers c LEFT JOIN stock s ON s.container_id=c.id
      WHERE c.active=1 AND c.is_virtual=0 AND {visible} AND (?='' OR c.name LIKE ? OR c.code LIKE ? OR c.location LIKE ?)
      GROUP BY c.id ORDER BY c.name COLLATE NOCASE""", (*visible_args, q, f"%{q}%", f"%{q}%", f"%{q}%")).fetchall()
    parents = get_db().execute(f"SELECT id,name,code,site FROM containers c WHERE active=1 AND {own} ORDER BY name COLLATE NOCASE", own_args).fetchall()
    virtual = get_db().execute("SELECT * FROM containers WHERE active=1 AND is_virtual=1 LIMIT 1").fetchone()
    virtual_product_count = 0
    if virtual:
        virtual_product_count = get_db().execute("SELECT COUNT(DISTINCT product_id) FROM stock WHERE container_id=?", (virtual["id"],)).fetchone()[0]
    return render_template("containers.html", containers=rows, q=q, parent_options=parents,
                           virtual=virtual, virtual_product_count=virtual_product_count)


@app.post("/contenedores/nuevo")
@login_required
@admin_required
def create_container():
    name = request.form.get("name", "").strip()
    location = request.form.get("location", "").strip()
    site = user_site() or request.form.get("site", "").strip()
    parent_id = request.form.get("parent_id", type=int)
    if not name:
        flash("Ingresá un nombre para el contenedor.", "error")
        return redirect(url_for("containers"))
    if parent_id:
        parent = get_db().execute("SELECT * FROM containers WHERE id=? AND active=1", (parent_id,)).fetchone()
        if not can_see(parent):
            flash("El contenedor principal ya no está disponible.", "error")
            return redirect(url_for("containers"))
        if parent["is_virtual"]:
            flash("El Depósito Virtual es aislado y no puede tener subcontenedores.", "error")
            return redirect(url_for("containers"))
        site = parent["site"]
    if site not in ("Bariloche", "Retiro"):
        flash("Elegí Bariloche o Retiro como depósito.", "error")
        return redirect(url_for("containers"))
    db = get_db()
    code = next_container_code(db)
    db.execute("INSERT INTO containers(code,name,location,site,parent_id,created_at) VALUES(?,?,?,?,?,?)", (code, name, location, site, parent_id, now_text()))
    db.commit()
    flash(f"Contenedor creado. Código {code}.", "success")
    return redirect(url_for("container_detail", container_id=db.execute("SELECT id FROM containers WHERE code=?", (code,)).fetchone()[0]))


@app.post("/contenedores/crear-lote")
@login_required
@admin_required
def create_container_batch():
    prefix = request.form.get("prefix", "").strip()
    location = request.form.get("location", "").strip()
    site = user_site() or request.form.get("site", "").strip()
    parent_value = request.form.get("parent_id", "")
    parent_id = int(parent_value) if parent_value.isdigit() else None
    try:
        start_text = request.form.get("start", "").strip()
        end_text = request.form.get("end", "").strip()
        start, end = int(start_text), int(end_text)
    except ValueError:
        flash("Ingresá números válidos para el inicio y el fin del rango.", "error")
        return redirect(url_for("containers"))
    if not prefix or start < 0 or end < start:
        flash("Revisá el prefijo y el rango: el inicio debe ser menor o igual al final.", "error")
        return redirect(url_for("containers"))
    if end - start + 1 > 500:
        flash("Podés crear hasta 500 contenedores por tanda.", "error")
        return redirect(url_for("containers"))
    db = get_db()
    if parent_id:
        parent = db.execute("SELECT * FROM containers WHERE id=? AND active=1", (parent_id,)).fetchone()
        if not can_see(parent):
            flash("El contenedor principal ya no está disponible.", "error")
            return redirect(url_for("containers"))
        if parent["is_virtual"]:
            flash("El Depósito Virtual es aislado y no puede tener subcontenedores.", "error")
            return redirect(url_for("containers"))
        site = parent["site"]
    if site not in ("Bariloche", "Retiro"):
        flash("Elegí Bariloche o Retiro como depósito.", "error")
        return redirect(url_for("containers"))
    width = max(len(start_text), len(end_text))
    names = [f"{prefix} {number:0{width}d}" for number in range(start, end + 1)]
    placeholders = ",".join("?" for _ in names)
    existing = db.execute(f"SELECT name FROM containers WHERE active=1 AND site=? AND name IN ({placeholders})", (site, *names)).fetchall()
    if existing:
        duplicates = ", ".join(row["name"] for row in existing[:5])
        flash(f"No se creó el lote porque ya existen nombres como: {duplicates}.", "error")
        return redirect(url_for("containers"))
    try:
        with db:
            for name in names:
                code = next_container_code(db)
                db.execute("INSERT INTO containers(code,name,location,site,parent_id,created_at) VALUES(?,?,?,?,?,?)",
                           (code, name, location, site, parent_id, now_text()))
    except sqlite3.IntegrityError:
        flash("No se pudo crear el lote completo. Revisá si algún código o nombre ya está registrado.", "error")
        return redirect(url_for("containers"))
    flash(f"Se crearon {len(names)} contenedores, desde {names[0]} hasta {names[-1]}, con códigos de barras únicos.", "success")
    return redirect(url_for("containers", q=prefix))


@app.route("/contenedores/<int:container_id>")
@login_required
def container_detail(container_id):
    db = get_db()
    container = visible_container(db, container_id)
    products = db.execute("""SELECT p.*,s.quantity FROM stock s JOIN products p ON p.id=s.product_id
      WHERE s.container_id=? ORDER BY p.name COLLATE NOCASE""", (container_id,)).fetchall()
    children = db.execute("""SELECT c.*,COALESCE(SUM(s.quantity),0) units,COUNT(DISTINCT s.product_id) products_count
      FROM containers c LEFT JOIN stock s ON s.container_id=c.id WHERE c.parent_id=? AND c.active=1
      GROUP BY c.id ORDER BY c.name COLLATE NOCASE""", (container_id,)).fetchall()
    parent = db.execute("SELECT id,name FROM containers WHERE id=?", (container["parent_id"],)).fetchone() if container["parent_id"] else None
    subtree = [container_id]
    for cid in subtree:
        subtree.extend(r[0] for r in db.execute("SELECT id FROM containers WHERE parent_id=? AND active=1", (cid,)).fetchall())
    excluded = ",".join("?" for _ in subtree)
    own, own_args = own_sql("containers")
    move_targets = db.execute(f"SELECT id,name,code,site FROM containers WHERE active=1 AND {own} AND id NOT IN ({excluded}) ORDER BY name COLLATE NOCASE", (*own_args, *subtree)).fetchall()
    return render_template("container_detail.html", container=container, products=products, children=children,
                           parent=parent, move_targets=move_targets)


@app.post("/contenedores/<int:container_id>/editar")
@login_required
@admin_required
def edit_container(container_id):
    db = get_db()
    container = visible_container(db, container_id)
    if container["is_virtual"]:
        flash("El Depósito Virtual no se puede editar ni mover.", "error")
        return redirect(url_for("container_detail", container_id=container_id))
    name = request.form.get("name", "").strip()
    location = request.form.get("location", "").strip()
    parent_value = request.form.get("parent_id", "")
    parent_id = int(parent_value) if parent_value.isdigit() else None
    subtree = [container_id]
    for cid in subtree:
        subtree.extend(r[0] for r in db.execute("SELECT id FROM containers WHERE parent_id=? AND active=1", (cid,)).fetchall())
    if not name:
        flash("El nombre del contenedor no puede quedar vacío.", "error")
    elif parent_id in subtree:
        flash("No podés mover un contenedor dentro de sí mismo o de una de sus cajas.", "error")
    else:
        parent = db.execute("SELECT * FROM containers WHERE id=? AND active=1", (parent_id,)).fetchone() if parent_id else None
        if parent_id and not can_see(parent):
            flash("El contenedor de destino no está disponible.", "error")
        elif parent and parent["is_virtual"]:
            flash("El Depósito Virtual es aislado y no puede contener otros contenedores.", "error")
        else:
            site = parent["site"] if parent else (user_site() or request.form.get("site", ""))
            if site not in ("Bariloche", "Retiro"):
                flash("Elegí Bariloche o Retiro como depósito.", "error")
            else:
                placeholders = ",".join("?" for _ in subtree)
                with db:
                    db.execute("UPDATE containers SET name=?,location=?,parent_id=?,site=? WHERE id=?", (name,location,parent_id,site,container_id))
                    db.execute(f"UPDATE containers SET site=? WHERE id IN ({placeholders})", [site,*subtree])
                flash("Contenedor actualizado.", "success")
                return redirect(url_for("container_detail", container_id=container_id))
    return redirect(url_for("container_detail", container_id=container_id) + "#editar")


@app.post("/contenedores/<int:container_id>/productos/<int:product_id>/quitar")
@login_required
@admin_required
def remove_container_product(container_id, product_id):
    db = get_db()
    visible_container(db, container_id)
    row = db.execute("SELECT quantity FROM stock WHERE container_id=? AND product_id=?", (container_id, product_id)).fetchone()
    if not row:
        flash("Ese producto ya no está en el contenedor.", "error")
    else:
        db.execute("DELETE FROM stock WHERE container_id=? AND product_id=?", (container_id, product_id))
        db.execute("INSERT INTO movements(created_at,user_id,from_container_id,to_container_id,from_source,product_id,quantity,note) VALUES(?,?,?,NULL,'container',?,?,?)",
                   (now_text(), session["user_id"], container_id, product_id, row["quantity"], "Retiro manual del contenedor"))
        db.commit()
        flash("Producto quitado del contenedor; la salida quedó registrada en movimientos.", "success")
    return redirect(url_for("container_detail", container_id=container_id))


@app.post("/contenedores/<int:container_id>/eliminar")
@login_required
@admin_required
def delete_container(container_id):
    db = get_db()
    container = visible_container(db, container_id)
    if container["is_virtual"]:
        flash("El Depósito Virtual protegido no se puede eliminar.", "error")
        return redirect(url_for("containers"))
    subtree = [container_id]
    for cid in subtree:
        subtree.extend(r[0] for r in db.execute("SELECT id FROM containers WHERE parent_id=? AND active=1", (cid,)).fetchall())
    placeholders = ",".join("?" for _ in subtree)
    if request.form.get("confirm") != "yes":
        abort(400)
    # Deleting a hierarchy also removes its remaining stock, recorded as outgoing adjustments.
    for cid in subtree:
        stock_rows = db.execute("SELECT product_id,quantity FROM stock WHERE container_id=?", (cid,)).fetchall()
        for row in stock_rows:
            db.execute("INSERT INTO movements(created_at,user_id,from_container_id,to_container_id,from_source,product_id,quantity,note) VALUES(?,?,?,NULL,'container',?,?,?)",
                       (now_text(), session["user_id"], cid, row["product_id"], row["quantity"], "Eliminación de contenedor"))
    db.execute(f"DELETE FROM stock WHERE container_id IN ({placeholders})", subtree)
    db.execute(f"UPDATE containers SET active=0 WHERE id IN ({placeholders})", subtree)
    db.commit()
    flash("Contenedor y sus subcontenedores eliminados; las cantidades se registraron como salidas.", "success")
    return redirect(url_for("containers"))


@app.route("/contenedores/<int:container_id>/etiqueta.svg")
@login_required
def container_label(container_id):
    container = visible_container(get_db(), container_id)
    code128 = barcode.get("code128", container["code"], writer=SVGWriter())
    svg = code128.render({"module_width": 0.28, "module_height": 14, "quiet_zone": 3,
                          "font_size": 10, "text_distance": 2, "write_text": True})
    response = make_response(svg)
    response.headers["Content-Type"] = "image/svg+xml; charset=utf-8"
    response.headers["Content-Disposition"] = f'inline; filename="contenedor-{container["code"]}.svg"'
    return response


@app.route("/contenedores/<int:container_id>/imprimir")
@login_required
def print_container_label(container_id):
    container = visible_container(get_db(), container_id)
    return render_template("label.html", container=container)


@app.route("/productos")
@login_required
def products():
    code = request.args.get("codigo", "").strip()
    q = request.args.get("q", "").strip()
    requested_sites = request.args.getlist("deposito")
    sites = [site for site in ("Bariloche", "Retiro") if site in requested_sites]
    if not request.args:
        sites = ["Bariloche", "Retiro"]
    # Un usuario de depósito ve el catálogo completo (compartido) con las ubicaciones de su depósito
    # y de los contenedores compartidos (planilla de stock).
    own_site = user_site()
    if own_site:
        sites = [own_site]
    db = get_db()
    archived = request.args.get("archivados") == "1"
    page_size = 100
    page = max(1, request.args.get("page", 1, type=int))
    offset = (page - 1) * page_size
    prefetched_has_next = None
    if archived:
        rows = db.execute("""SELECT p.*,p.legacy_units+COALESCE(SUM(s.quantity),0) units,
          COALESCE(GROUP_CONCAT(c.name || ' (' || s.quantity || ')', ' · '),p.source_location,'') locations
          FROM products p LEFT JOIN stock s ON s.product_id=p.id AND s.quantity>0
          LEFT JOIN containers c ON c.id=s.container_id AND (? IS NULL OR c.site=? OR c.is_virtual=1 OR c.from_stock_sheet=1)
          WHERE p.active=0 AND (?='' OR p.sku LIKE ? OR p.name LIKE ? OR p.item LIKE ? OR p.color LIKE ? OR p.size LIKE ? OR p.description LIKE ?)
            AND (?='' OR p.sku LIKE ?)
          GROUP BY p.id ORDER BY p.sku COLLATE NOCASE LIMIT ? OFFSET ?""",
          (own_site,own_site,q,f"%{q}%",f"%{q}%",f"%{q}%",f"%{q}%",f"%{q}%",f"%{q}%",code,f"%{code}%",page_size+1,offset)).fetchall()
    elif (len(sites) == 2 or own_site) and not q and not code:
        page_ids = db.execute("SELECT id FROM products WHERE active=1 ORDER BY sku COLLATE NOCASE LIMIT ? OFFSET ?",
                              (page_size + 1, offset)).fetchall()
        prefetched_has_next = len(page_ids) > page_size
        ids = [row["id"] for row in page_ids[:page_size]]
        if ids:
            placeholders = ",".join("?" for _ in ids)
            site_placeholders = ",".join("?" for _ in sites)
            rows = db.execute(f"""SELECT p.*,p.legacy_units+COALESCE(SUM(CASE WHEN c.id IS NOT NULL THEN s.quantity ELSE 0 END),0) units,
              CASE WHEN p.legacy_units>0 THEN 'Pendiente de ubicación (' || p.legacy_units || ')'
                || CASE WHEN GROUP_CONCAT(CASE WHEN c.id IS NOT NULL THEN c.name || ' (' || s.quantity || ')' END, ' · ') IS NOT NULL
                  THEN ' · ' || GROUP_CONCAT(CASE WHEN c.id IS NOT NULL THEN c.name || ' (' || s.quantity || ')' END, ' · ') ELSE '' END
                ELSE COALESCE(GROUP_CONCAT(CASE WHEN c.id IS NOT NULL THEN c.name || ' (' || s.quantity || ')' END, ' · '),NULLIF(p.source_location,''),'') END locations
              FROM products p LEFT JOIN stock s ON s.product_id=p.id AND s.quantity>0
              LEFT JOIN containers c ON c.id=s.container_id AND c.active=1 AND (c.site IN ({site_placeholders}) OR c.from_stock_sheet=1)
              WHERE p.active=1 AND p.id IN ({placeholders})
              GROUP BY p.id ORDER BY p.sku COLLATE NOCASE""", (*sites, *ids)).fetchall()
        else:
            rows = []
    elif len(sites) == 2 or own_site:
        placeholders = ",".join("?" for _ in sites)
        rows = db.execute(f"""SELECT p.*,p.legacy_units+COALESCE(SUM(CASE WHEN c.id IS NOT NULL THEN s.quantity ELSE 0 END),0) units,
          CASE WHEN p.legacy_units>0 THEN 'Pendiente de ubicación (' || p.legacy_units || ')'
            || CASE WHEN GROUP_CONCAT(CASE WHEN c.id IS NOT NULL THEN c.name || ' (' || s.quantity || ')' END, ' · ') IS NOT NULL
              THEN ' · ' || GROUP_CONCAT(CASE WHEN c.id IS NOT NULL THEN c.name || ' (' || s.quantity || ')' END, ' · ') ELSE '' END
            ELSE COALESCE(GROUP_CONCAT(CASE WHEN c.id IS NOT NULL THEN c.name || ' (' || s.quantity || ')' END, ' · '),NULLIF(p.source_location,''),'') END locations
          FROM products p LEFT JOIN stock s ON s.product_id=p.id AND s.quantity>0
          LEFT JOIN containers c ON c.id=s.container_id AND c.active=1 AND (c.site IN ({placeholders}) OR c.from_stock_sheet=1)
          WHERE p.active=1
            AND (?='' OR p.sku LIKE ? OR p.name LIKE ? OR p.item LIKE ? OR p.color LIKE ? OR p.size LIKE ? OR p.description LIKE ? OR c.name LIKE ? OR c.location LIKE ?)
            AND (?='' OR p.sku LIKE ?)
          GROUP BY p.id ORDER BY p.sku COLLATE NOCASE LIMIT ? OFFSET ?""",
          (*sites, q, f"%{q}%", f"%{q}%", f"%{q}%", f"%{q}%", f"%{q}%", f"%{q}%", f"%{q}%", f"%{q}%", code, f"%{code}%", page_size+1, offset)).fetchall()
    elif len(sites) == 1:
        rows = db.execute("""SELECT p.*,COALESCE(SUM(s.quantity),0) units,
          GROUP_CONCAT(c.name || ' (' || s.quantity || ')', ' · ') locations
          FROM products p JOIN stock s ON s.product_id=p.id AND s.quantity>0
          JOIN containers c ON c.id=s.container_id AND c.active=1 AND c.site=?
          WHERE p.active=1
            AND (?='' OR p.sku LIKE ? OR p.name LIKE ? OR p.item LIKE ? OR p.color LIKE ? OR p.size LIKE ? OR p.description LIKE ? OR c.name LIKE ? OR c.location LIKE ?)
            AND (?='' OR p.sku LIKE ?)
          GROUP BY p.id ORDER BY p.sku COLLATE NOCASE LIMIT ? OFFSET ?""",
          (sites[0],q,f"%{q}%",f"%{q}%",f"%{q}%",f"%{q}%",f"%{q}%",f"%{q}%",f"%{q}%",f"%{q}%",code,f"%{code}%",page_size+1,offset)).fetchall()
    else:
        rows = []
    has_next = prefetched_has_next if prefetched_has_next is not None else len(rows) > page_size
    rows = rows[:page_size]
    page_args = {"codigo": code, "q": q, "deposito": sites}
    if own_site:
        page_args.pop("deposito")
    if archived:
        page_args["archivados"] = "1"
    previous_url = url_for("products", page=page-1, **page_args) if page > 1 else ""
    next_url = url_for("products", page=page+1, **page_args) if has_next else ""
    start_item = offset + 1 if rows else 0
    end_item = offset + len(rows)
    return render_template("products.html", products=rows, q=q, code=code, sites=sites, archived=archived,
                           page=page, page_size=page_size, has_next=has_next, previous_url=previous_url,
                           next_url=next_url, start_item=start_item, end_item=end_item)


@app.get("/buscar-talles")
@login_required
def size_finder():
    code = request.args.get("codigo", "").strip()
    color = request.args.get("color", "").strip().upper()
    rows = []
    searched = bool(request.args)
    visible, visible_args = visible_sql("c")
    if code:
        if color:
            finder_condition = "((p.item=? AND p.color=?) OR p.sku=?)"
            finder_params = (code, color, code)
        else:
            finder_condition = "p.sku=?"
            finder_params = (code,)
        rows = get_db().execute(f"""SELECT p.sku,p.name,p.item,p.color,p.size,
          COALESCE(SUM(CASE WHEN c.is_virtual=0 AND c.active=1 THEN s.quantity ELSE 0 END),0) physical_units,
          GROUP_CONCAT(CASE
            WHEN c.is_virtual=1 AND c.active=1 THEN 'DEPOSITO VIRTUAL · listo para PICK'
            WHEN c.is_virtual=0 AND c.active=1 AND s.quantity>0 THEN c.name || ' (' || s.quantity || ')'
          END, ' · ') locations
          FROM products p
          LEFT JOIN stock s ON s.product_id=p.id
          LEFT JOIN containers c ON c.id=s.container_id AND {visible}
          WHERE p.active=1 AND {finder_condition}
          GROUP BY p.id
          ORDER BY CASE UPPER(TRIM(p.size))
            WHEN 'XXS' THEN 1 WHEN 'XS' THEN 2 WHEN 'S' THEN 3 WHEN 'M' THEN 4
            WHEN 'L' THEN 5 WHEN 'XL' THEN 6 WHEN 'XXL' THEN 7 WHEN 'XXXL' THEN 8 ELSE 20 END,
            p.size COLLATE NOCASE,p.sku""", (*visible_args, *finder_params)).fetchall()
    if request.args.get("formato") == "json":
        return jsonify(results=[dict(row) for row in rows], searched=searched,
                       code=code, color=color)
    return render_template("size_finder.html", results=rows, code=code, color=color, searched=searched)


@app.post("/productos/<int:product_id>/archivar")
@login_required
@admin_required
def archive_product(product_id):
    db = get_db()
    product = db.execute("SELECT id FROM products WHERE id=? AND active=1", (product_id,)).fetchone()
    if not product:
        abort(404)
    held = db.execute("SELECT COALESCE(SUM(quantity),0) FROM stock WHERE product_id=?", (product_id,)).fetchone()[0]
    if held:
        flash("No se puede archivar: todavía hay unidades ubicadas en contenedores.", "error")
    else:
        db.execute("UPDATE products SET active=0,legacy_units=0 WHERE id=?", (product_id,))
        db.commit()
        flash("Producto archivado. Podés recuperarlo desde Productos archivados.", "success")
    return redirect(url_for("products", deposito=["Retiro", "Bariloche"]))


@app.post("/productos/<int:product_id>/restaurar")
@login_required
@admin_required
def restore_product(product_id):
    db = get_db()
    if not db.execute("SELECT 1 FROM products WHERE id=? AND active=0", (product_id,)).fetchone():
        abort(404)
    db.execute("UPDATE products SET active=1 WHERE id=?", (product_id,))
    db.commit()
    flash("Producto recuperado del archivo.", "success")
    return redirect(url_for("products", deposito=["Retiro", "Bariloche"]))


@app.post("/productos/nuevo")
@login_required
@admin_required
def create_product():
    sku = request.form.get("sku", "").strip()
    name = request.form.get("name", "").strip()
    item = request.form.get("item", "").strip()
    color = request.form.get("color", "").strip().upper()
    size = request.form.get("size", "").strip()
    description = request.form.get("description", "").strip()
    color_name = request.form.get("color_name", "").strip().upper()
    season = request.form.get("season", "").strip().upper()
    if not sku or not name:
        flash("El código/SKU y el nombre son obligatorios.", "error")
        return redirect(url_for("products"))
    try:
        db = get_db()
        db.execute("INSERT INTO products(sku,name,item,color,size,description,color_name,season,created_at) VALUES(?,?,?,?,?,?,?,?,?)",
                   (sku, name, item, color, size, description, color_name, season, now_text()))
        db.commit()
    except sqlite3.IntegrityError:
        flash("Ese código de producto ya existe.", "error")
    else:
        flash("Producto agregado.", "success")
    return redirect(url_for("products"))


@app.route("/stock")
@login_required
def stock_view():
    # Mismas columnas que Productos. Solo cuenta los racks del local cargados en el sistema:
    # excluye el Depósito Virtual y los racks generados por la planilla de stock.
    q = request.args.get("q", "").strip()
    page_size = 100
    page = max(1, request.args.get("page", 1, type=int))
    offset = (page - 1) * page_size
    like = f"%{q}%"
    own, own_args = own_sql("c")
    rows = get_db().execute("""SELECT p.id,p.sku,p.name,p.item,p.color,p.size,p.description,
      SUM(s.quantity) units,
      GROUP_CONCAT(c.id || char(31) || c.name || char(31) || s.quantity, char(30)) locations
      FROM stock s JOIN containers c ON c.id=s.container_id JOIN products p ON p.id=s.product_id
      WHERE c.active=1 AND {own} AND s.quantity>0
        AND (?='' OR p.sku LIKE ? OR p.name LIKE ? OR p.item LIKE ? OR p.color LIKE ? OR p.size LIKE ?
             OR p.description LIKE ? OR c.name LIKE ? OR c.code LIKE ? OR c.location LIKE ?)
      GROUP BY p.id ORDER BY p.sku COLLATE NOCASE LIMIT ? OFFSET ?""".replace("{own}", own),
      (*own_args, q, like, like, like, like, like, like, like, like, like, page_size + 1, offset)).fetchall()
    has_next = len(rows) > page_size
    products = []
    for row in rows[:page_size]:
        product = dict(row)
        product["locations"] = [
            {"id": int(cid), "name": name, "quantity": int(qty)}
            for cid, name, qty in (part.split("\x1f") for part in (row["locations"] or "").split("\x1e") if part)]
        products.append(product)
    diagnostics = None
    if not products and not q:
        db = get_db()
        diagnostics = {
            "racks": db.execute(f"SELECT COUNT(*) FROM containers c WHERE c.active=1 AND {own}", own_args).fetchone()[0],
            "virtual_units": db.execute("SELECT COALESCE(SUM(virtual_units),0) FROM products WHERE active=1").fetchone()[0],
        }
    previous_url = url_for("stock_view", q=q, page=page - 1) if page > 1 else ""
    next_url = url_for("stock_view", q=q, page=page + 1) if has_next else ""
    return render_template("stock.html", products=products, q=q, page_size=page_size,
                           start_item=offset + 1 if products else 0, end_item=offset + len(products),
                           previous_url=previous_url, next_url=next_url, diagnostics=diagnostics)


@app.route("/lector")
@login_required
def scan():
    code = request.args.get("codigo", "").strip()
    db = get_db()
    container = db.execute("SELECT * FROM containers WHERE code=? AND active=1", (code,)).fetchone() if code else None
    if not can_see(container):
        container = None
    product = db.execute("SELECT * FROM products WHERE sku=? AND active=1", (code,)).fetchone() if code else None
    visible, visible_args = visible_sql("c")
    container_stock = []
    product_stock = []
    if container:
        container_stock = db.execute("""SELECT p.sku,p.name,s.quantity FROM stock s JOIN products p ON p.id=s.product_id
          WHERE s.container_id=? AND s.quantity>0 ORDER BY p.name COLLATE NOCASE""", (container["id"],)).fetchall()
    if product:
        product_stock = db.execute(f"""SELECT c.id,c.name,c.code,c.location,s.quantity FROM stock s JOIN containers c ON c.id=s.container_id
          WHERE s.product_id=? AND s.quantity>0 AND c.active=1 AND {visible} ORDER BY c.name COLLATE NOCASE""", (product["id"], *visible_args)).fetchall()
    return render_template("scan.html", code=code, container=container, product=product,
                           container_stock=container_stock, product_stock=product_stock)


@app.route("/movimientos", methods=["GET", "POST"])
@login_required
@admin_for_writes
def movements():
    db = get_db()
    if request.method == "POST":
        if request.form.get("mode") == "pick_batch":
            from_value = request.form.get("from_id", "")
            legacy_source = from_value == "legacy"
            from_id = None if legacy_source else (int(from_value) if from_value.isdigit() else None)
            to_id = request.form.get("to_id", type=int) or None
            note = request.form.get("note", "").strip()
            try:
                scanned_items = json.loads(request.form.get("items", "[]"))
                quantities = {}
                for item in scanned_items:
                    pid, qty = int(item["product_id"]), int(item["quantity"])
                    if qty < 1:
                        raise ValueError("La cantidad escaneada no es válida.")
                    quantities[pid] = quantities.get(pid, 0) + qty
                if not quantities:
                    raise ValueError("Escaneá al menos una prenda antes de mover.")
                if not from_id and not legacy_source:
                    raise ValueError("Escaneá o seleccioná un contenedor de origen.")
                if not to_id:
                    raise ValueError("Escaneá o seleccioná un contenedor de destino.")
                if from_id == to_id:
                    raise ValueError("El origen y el destino deben ser distintos.")
                source_container = db.execute("SELECT * FROM containers WHERE id=? AND active=1", (from_id,)).fetchone() if from_id else None
                if from_id and not can_see(source_container):
                    raise ValueError("El contenedor de origen ya no está disponible.")
                destination_container = db.execute("SELECT * FROM containers WHERE id=? AND active=1", (to_id,)).fetchone()
                if not can_see(destination_container):
                    raise ValueError("El contenedor de destino ya no está disponible.")
                if destination_container["is_virtual"]:
                    raise ValueError("El Depósito Virtual solo se puede usar como origen PICK.")
                with db:
                    for pid, quantity in quantities.items():
                        product = db.execute("SELECT id,legacy_units,virtual_units,name FROM products WHERE id=? AND active=1", (pid,)).fetchone()
                        if not product:
                            raise ValueError("Uno de los productos escaneados no existe o está archivado.")
                        if legacy_source:
                            if product["legacy_units"] < quantity:
                                raise ValueError(f"No hay {quantity} unidades pendientes de {product['name']} en la planilla.")
                            db.execute("UPDATE products SET legacy_units=legacy_units-? WHERE id=?", (quantity, pid))
                            source_label = "legacy"
                        elif source_container and source_container["is_virtual"]:
                            if product["virtual_units"] < quantity:
                                raise ValueError(f"No hay {quantity} unidades disponibles de {product['name']} en el Depósito Virtual.")
                            db.execute("UPDATE products SET virtual_units=virtual_units-? WHERE id=?", (quantity, pid))
                            source_label = "virtual"
                        else:
                            available = db.execute("SELECT quantity FROM stock WHERE container_id=? AND product_id=?", (from_id, pid)).fetchone()
                            if not available or available["quantity"] < quantity:
                                raise ValueError(f"No hay {quantity} unidades de {product['name']} en el contenedor de origen.")
                            remaining = available["quantity"] - quantity
                            if remaining:
                                db.execute("UPDATE stock SET quantity=? WHERE container_id=? AND product_id=?", (remaining, from_id, pid))
                            else:
                                db.execute("DELETE FROM stock WHERE container_id=? AND product_id=?", (from_id, pid))
                            source_label = "external"
                        db.execute("INSERT INTO stock(container_id,product_id,quantity) VALUES(?,?,?) ON CONFLICT(container_id,product_id) DO UPDATE SET quantity=quantity+excluded.quantity", (to_id, pid, quantity))
                        db.execute("INSERT INTO movements(created_at,user_id,from_container_id,to_container_id,from_source,product_id,quantity,note) VALUES(?,?,?,?,?,?,?,?)",
                                   (now_text(), session["user_id"], from_id, to_id, source_label, pid, quantity, note))
                flash(f"Traslado completado: {sum(quantities.values())} unidades de {len(quantities)} productos.", "success")
                return redirect(url_for("movements"))
            except (ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
                flash(str(exc) or "No se pudo completar el movimiento.", "error")
        else:
            flash("Usá el flujo de PICK para registrar movimientos por lote.", "error")
    own_f, own_f_args = own_sql("f", strict=False)
    own_t, own_t_args = own_sql("t", strict=False)
    move_filter = f"({own_f} OR {own_t} OR (m.from_container_id IS NULL AND m.to_container_id IS NULL))"
    move_args = (*own_f_args, *own_t_args)
    rows = db.execute("""SELECT m.*,p.sku,p.name product,
      CASE WHEN f.name IS NOT NULL THEN f.name WHEN m.from_source='legacy' THEN 'Planilla sin ubicación' WHEN m.from_source='container' THEN 'Retiro manual' ELSE 'Ingreso externo' END from_name,
      t.name to_name
      FROM movements m JOIN products p ON p.id=m.product_id
      LEFT JOIN containers f ON f.id=m.from_container_id LEFT JOIN containers t ON t.id=m.to_container_id
      WHERE {move_filter}
      ORDER BY m.id DESC LIMIT 100""".replace("{move_filter}", move_filter), move_args).fetchall()
    products = db.execute("SELECT * FROM products WHERE active=1 ORDER BY name COLLATE NOCASE").fetchall()
    visible, visible_args = visible_sql("c")
    containers = db.execute(f"SELECT * FROM containers c WHERE active=1 AND is_virtual=0 AND {visible} ORDER BY name COLLATE NOCASE", visible_args).fetchall()
    virtual_container = db.execute("SELECT * FROM containers WHERE active=1 AND is_virtual=1 LIMIT 1").fetchone()
    product_index = {p["sku"]: {"id": p["id"], "name": p["name"], "sku": p["sku"]} for p in products}
    source_index = {c["code"]: {"id": c["id"], "name": c["name"], "code": c["code"]} for c in containers}
    destination_index = dict(source_index)
    if virtual_container:
        source_index[virtual_container["code"]] = {"id": virtual_container["id"], "name": virtual_container["name"], "code": virtual_container["code"]}
    return render_template("movements.html", rows=rows, products=products, containers=containers,
                           product_index=product_index, source_index=source_index,
                           destination_index=destination_index, virtual_container=virtual_container)


def optional_label_fields(row):
    """Columnas opcionales para la etiqueta de 3 × 5 cm: nombre del color y temporada."""
    color_name = row.get("color_nombre") or row.get("nombre_color") or row.get("color_name") or ""
    season = row.get("temporada") or row.get("season") or ""
    return str(color_name).strip().upper(), str(season).strip().upper()


def import_elementos_sheet(db, sheet):
    headers = [str(cell or "").strip().lower() for cell in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]
    required = ["codigo", "nombre", "item", "color", "talle", "descripcion", "unidades"]
    if not set(required).issubset(headers):
        raise ValueError("La hoja elementos debe incluir codigo, nombre, item, color, talle, descripcion y unidades.")
    rows = [dict(zip(headers, row)) for row in sheet.iter_rows(min_row=2, values_only=True) if any(v is not None for v in row)]
    seen = set()
    total_units = 0
    cleaned = []
    for row_number, row in enumerate(rows, start=2):
        code = str(row.get("codigo") or "").strip()
        name = str(row.get("nombre") or "").strip()
        if not code or not name:
            raise ValueError(f"Falta código o nombre en la fila {row_number}.")
        if code in seen:
            raise ValueError(f"El código {code} está repetido en la planilla.")
        seen.add(code)
        # The original export's "unidades" is not treated as a physical inventory count.
        raw_units = row.get("unidades")
        try:
            units = int(raw_units or 0)
            if units < 0 or (raw_units is not None and float(raw_units) != units):
                raise ValueError
        except (TypeError, ValueError):
            raise ValueError(f"La cantidad de unidades no es válida en la fila {row_number}.")
        total_units += units
        cleaned.append((code,name,str(row.get("item") or "").strip(),str(row.get("color") or "").strip().upper(),
          str(row.get("talle") or "").strip(),str(row.get("descripcion") or "").strip(),units,
          str(row.get("ubicaciones") or "").strip(), *optional_label_fields(row)))
    with db:
        virtual = db.execute("SELECT id FROM containers WHERE is_virtual=1 AND active=1 LIMIT 1").fetchone()
        if not virtual:
            raise ValueError("No se encontró el Depósito Virtual. Reiniciá el servicio para actualizarlo.")
        for item in cleaned:
            db.execute("""INSERT INTO products(sku,name,item,color,size,description,legacy_units,virtual_units,source_location,color_name,season,created_at)
              VALUES(?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(sku) DO UPDATE SET
              name=excluded.name,item=excluded.item,color=excluded.color,size=excluded.size,
              description=excluded.description,source_location=excluded.source_location,
              color_name=CASE WHEN excluded.color_name!='' THEN excluded.color_name ELSE products.color_name END,
              season=CASE WHEN excluded.season!='' THEN excluded.season ELSE products.season END,
              virtual_units=excluded.virtual_units,legacy_units=0,active=1""",
              (item[0],item[1],item[2],item[3],item[4],item[5],0,item[6],item[7],item[8],item[9],now_text()))
            product_id = db.execute("SELECT id FROM products WHERE sku=?", (item[0],)).fetchone()[0]
            # Zero quantity records membership in the virtual source without claiming real stock.
            db.execute("INSERT INTO stock(container_id,product_id,quantity) VALUES(?,?,0) ON CONFLICT(container_id,product_id) DO NOTHING",
                       (virtual["id"], product_id))
    location_count = sum(bool(item[7]) for item in cleaned)
    return len(cleaned), total_units, location_count


# ---------------------------------------------------------------------------
# Planilla de stock por racks (hoja STOCK: columnas M a U = racks, K = RETAIL)
# ---------------------------------------------------------------------------
IVA_RATE = 0.21
STOCK_RACK_COLUMNS = range(12, 21)  # columnas M..U (0 = A)
STOCK_PRICE_COLUMN = 10             # columna K (RETAIL, precio final con IVA)


def price_without_iva(price):
    """El precio RETAIL ya incluye 21 % de IVA: el neto es precio / 1,21."""
    return None if price is None else round(price / (1 + IVA_RATE), 2)


def format_ars(value, decimals=None):
    """Formato argentino: $ 74.900 o $ 61.900,83."""
    if value is None:
        return ""
    if decimals is None:
        decimals = 0 if float(value).is_integer() else 2
    text = f"{value:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"$ {text}"


app.jinja_env.filters["ars"] = format_ars


def _cell_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return " ".join(str(value).split())


def import_stock_sheet(db, sheet, site):
    """Crea/actualiza un rack por cada columna M..U y deja su stock igual al de la planilla."""
    rows = sheet.iter_rows(values_only=True)
    headers = [_cell_text(h) for h in next(rows)]
    upper = [h.upper() for h in headers]
    if len(headers) < 21:
        raise ValueError("La planilla de stock debe tener al menos hasta la columna U (racks en M a U).")
    if upper[STOCK_PRICE_COLUMN] != "RETAIL":
        raise ValueError("La columna K de la planilla de stock debe llamarse RETAIL.")

    def col(name, default):
        return upper.index(name) if name in upper else default
    c_item, c_color, c_size, c_name = col("SKU", 1), col("COLOR", 2), col("TALLE", 3), col("NOMBRE", 4)
    c_year, c_season = col("AÑO", 6), col("TEMP", 7)
    racks = [(index, headers[index]) for index in STOCK_RACK_COLUMNS]
    if any(not name for _, name in racks):
        raise ValueError("Cada columna de M a U debe tener en la fila 1 el nombre del rack.")

    merged, order = {}, []
    negatives = duplicates = 0
    for row_number, row in enumerate(rows, start=2):
        row = list(row) + [None] * (21 - len(row))
        item = _cell_text(row[c_item])
        if not item or _cell_text(row[c_name]).upper() == "TOTAL":
            continue  # filas vacías y fila de totales
        color = _cell_text(row[c_color]).upper()
        size = _cell_text(row[c_size])
        key = (item, color, size.upper())
        quantities = []
        for index, rack in racks:
            raw = row[index]
            if raw in (None, "", "-"):
                qty = 0
            elif isinstance(raw, (int, float)) and float(raw).is_integer():
                qty = int(raw)
            else:
                raise ValueError(f"Cantidad no válida en la fila {row_number}, rack {rack}: {raw!r}.")
            if qty < 0:
                negatives += 1
                qty = 0
            quantities.append(qty)
        price = row[STOCK_PRICE_COLUMN]
        price = float(price) if isinstance(price, (int, float)) and price > 0 else None
        year, temp = row[c_year], _cell_text(row[c_season]).upper()
        season = f"{temp}{int(year) % 100:02d}" if temp in ("SP", "SU", "FA", "WI", "HO") and isinstance(year, (int, float)) else ""
        if key in merged:
            duplicates += 1
            entry = merged[key]
            entry["qty"] = [a + b for a, b in zip(entry["qty"], quantities)]
            entry["price"] = entry["price"] or price
        else:
            merged[key] = {"item": item, "color": color, "size": size, "name": _cell_text(row[c_name]) or item,
                           "season": season, "price": price, "qty": quantities}
            order.append(key)
    if not merged:
        raise ValueError("La planilla de stock no tiene artículos.")

    with db:
        rack_ids, created_racks = [], 0
        for _, rack in racks:
            name = f"RACK {rack.upper()}"
            # Solo se reutilizan racks creados por esta planilla; los racks del local no se tocan.
            found = db.execute("SELECT id FROM containers WHERE active=1 AND is_virtual=0 AND from_stock_sheet=1 "
                               "AND site=? AND UPPER(name)=?", (site, name)).fetchone()
            if not found:
                code = next_container_code(db)
                db.execute("INSERT INTO containers(code,name,location,site,parent_id,from_stock_sheet,created_at) "
                           "VALUES(?,?,?,?,NULL,1,?)", (code, name, "Planilla de stock", site, now_text()))
                found = db.execute("SELECT id FROM containers WHERE code=?", (code,)).fetchone()
                created_racks += 1
            rack_ids.append(found["id"])
        # La planilla es una foto completa: se reemplaza el stock de estos racks.
        placeholders = ",".join("?" for _ in rack_ids)
        db.execute(f"DELETE FROM stock WHERE container_id IN ({placeholders})", rack_ids)

        index = {}
        for p in db.execute("SELECT id,item,color,size FROM products ORDER BY active DESC,id"):
            index.setdefault((p["item"].strip(), p["color"].strip().upper(), p["size"].strip().upper()), []).append(p["id"])
        new_products = priced = units = 0
        for key in order:
            entry = merged[key]
            ids = index.get(key)
            if not ids:
                sku = " ".join(part for part in (entry["item"], entry["color"], entry["size"]) if part)
                existing = db.execute("SELECT id FROM products WHERE sku=?", (sku,)).fetchone()
                if existing:
                    ids = [existing["id"]]
                    db.execute("UPDATE products SET item=?,color=?,size=? WHERE id=?",
                               (entry["item"], entry["color"], entry["size"], existing["id"]))
                else:
                    db.execute("""INSERT INTO products(sku,name,item,color,size,description,season,created_at)
                      VALUES(?,?,?,?,?,'',?,?)""", (sku, entry["name"], entry["item"], entry["color"],
                                                     entry["size"], entry["season"], now_text()))
                    ids = [db.execute("SELECT id FROM products WHERE sku=?", (sku,)).fetchone()["id"]]
                    new_products += 1
                index[key] = ids
            if entry["price"] is not None:
                priced += 1
            placeholders = ",".join("?" for _ in ids)
            # El precio se aplica a todas las variantes (códigos de barras) de ese artículo/color/talle.
            db.execute(f"""UPDATE products SET retail_price=COALESCE(?,retail_price),
              season=CASE WHEN season='' THEN ? ELSE season END,active=1 WHERE id IN ({placeholders})""",
              (entry["price"], entry["season"], *ids))
            for rack_id, qty in zip(rack_ids, entry["qty"]):
                if qty > 0:
                    db.execute("INSERT INTO stock(container_id,product_id,quantity) VALUES(?,?,?)", (rack_id, ids[0], qty))
                    units += qty
    return {"articles": len(order), "racks": [f"RACK {r.upper()}" for _, r in racks], "created_racks": created_racks,
            "new_products": new_products, "priced": priced, "units": units,
            "negatives": negatives, "duplicates": duplicates}


@app.route("/importar", methods=["GET", "POST"])
@login_required
@admin_for_writes
def import_data():
    if request.method == "POST":
        upload = request.files.get("file")
        if not upload or not upload.filename.lower().endswith((".xlsx", ".xlsm")):
            flash("Subí un archivo Excel .xlsx con las hojas Contenedores y Stock.", "error")
            return redirect(url_for("import_data"))
        try:
            workbook = load_workbook(upload, read_only=True, data_only=True)
            if request.form.get("tipo") == "stock":
                site = request.form.get("site", "Retiro")
                if site not in ("Bariloche", "Retiro"):
                    raise ValueError("Elegí Bariloche o Retiro como depósito de los racks.")
                sheet = workbook["STOCK"] if "STOCK" in workbook.sheetnames else workbook.worksheets[0]
                r = import_stock_sheet(get_db(), sheet, site)
                n = lambda v: f"{v:,}".replace(",", ".")
                message = (f"Planilla de stock importada en {site}: {n(r['articles'])} artículos y {n(r['units'])} unidades "
                           f"en {len(r['racks'])} racks ({', '.join(r['racks'])}); {r['created_racks']} racks nuevos. "
                           f"Precios RETAIL cargados para {n(r['priced'])} artículos.")
                if r["new_products"]:
                    message += f" Se dieron de alta {n(r['new_products'])} artículos que no estaban en el catálogo."
                if r["negatives"]:
                    message += f" {r['negatives']} cantidades negativas se tomaron como 0."
                if r["duplicates"]:
                    message += f" {r['duplicates']} filas repetidas se sumaron."
                flash(message, "success")
                return redirect(url_for("import_data"))
            if "elementos" in workbook.sheetnames:
                db = get_db()
                count, total_units, location_count = import_elementos_sheet(db, workbook["elementos"])
                flash(f"Importación lista: {count:,} productos cargados en DEPOSITO VIRTUAL con {total_units:,} unidades lógicas para PICK. No se sumaron al stock físico.".replace(",", "."), "success")
                return redirect(url_for("import_data"))
            if not {"Contenedores", "Productos", "Stock"}.issubset(workbook.sheetnames):
                raise ValueError("No encontré la hoja elementos ni la plantilla con hojas Contenedores, Productos y Stock.")
            db = get_db()
            parsed = {}
            for sheet_name, required in {
                "Contenedores": ["codigo", "nombre", "ubicacion", "sucursal"],
                "Productos": ["sku", "nombre", "item", "color", "talle", "descripcion"],
                "Stock": ["codigo_contenedor", "sku", "cantidad"],
            }.items():
                sheet = workbook[sheet_name]
                headers = [str(c or "").strip().lower() for c in next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))]
                if not set(required).issubset(headers):
                    raise ValueError(f"La hoja {sheet_name} debe incluir columnas: {', '.join(required)}.")
                parsed[sheet_name] = [dict(zip(headers, row)) for row in sheet.iter_rows(min_row=2, values_only=True) if any(v is not None for v in row)]
            container_codes = {str(r["codigo"]).strip() for r in parsed["Contenedores"] if r.get("codigo") and r.get("nombre")}
            product_skus = {str(r["sku"]).strip() for r in parsed["Productos"] if r.get("sku") and r.get("nombre")}
            for r in parsed["Stock"]:
                if str(r.get("codigo_contenedor", "")).strip() not in container_codes or str(r.get("sku", "")).strip() not in product_skus:
                    raise ValueError("Cada fila de Stock debe apuntar a un código de contenedor y SKU incluidos en el mismo Excel.")
                if not isinstance(r.get("cantidad"), (int, float)) or int(r["cantidad"]) != r["cantidad"] or r["cantidad"] < 0:
                    raise ValueError("Las cantidades de Stock deben ser números enteros iguales o mayores que cero.")
            if len(container_codes) != len(parsed["Contenedores"]) or len(product_skus) != len(parsed["Productos"]):
                raise ValueError("Hay códigos o SKU vacíos o repetidos en el archivo.")
            with db:
                for r in parsed["Contenedores"]:
                    site = str(r.get("sucursal") or r.get("deposito") or "").strip().title()
                    if site not in ("Bariloche", "Retiro"):
                        raise ValueError("La sucursal de cada contenedor debe ser Bariloche o Retiro.")
                    if user_site() and site != user_site():
                        raise ValueError(f"Tu usuario solo puede importar contenedores de {user_site()}.")
                    db.execute("""INSERT INTO containers(code,name,location,site,created_at) VALUES(?,?,?,?,?)
                      ON CONFLICT(code) DO UPDATE SET name=excluded.name,location=excluded.location,site=excluded.site,active=1""",
                      (str(r["codigo"]).strip(),str(r["nombre"]).strip(),str(r.get("ubicacion") or "").strip(),site,now_text()))
                for r in parsed["Productos"]:
                    pending = r.get("unidades_sin_ubicar") or 0
                    if not isinstance(pending, (int,float)) or int(pending) != pending or pending < 0:
                        raise ValueError("Las unidades sin ubicar deben ser números enteros iguales o mayores que cero.")
                    db.execute("""INSERT INTO products(sku,name,item,color,size,description,legacy_units,source_location,color_name,season,retail_price,created_at)
                      VALUES(?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(sku) DO UPDATE SET
                      retail_price=COALESCE(excluded.retail_price,products.retail_price),
                      name=excluded.name,item=excluded.item,color=excluded.color,size=excluded.size,
                      description=excluded.description,source_location=excluded.source_location,
                      color_name=CASE WHEN excluded.color_name!='' THEN excluded.color_name ELSE products.color_name END,
                      season=CASE WHEN excluded.season!='' THEN excluded.season ELSE products.season END,active=1""",
                      (str(r["sku"]).strip(),str(r["nombre"]).strip(),str(r.get("item") or "").strip(),str(r.get("color") or "").strip().upper(),str(r.get("talle") or r.get("size") or "").strip(),str(r.get("descripcion") or "").strip(),int(pending),str(r.get("ubicaciones") or "").strip(),*optional_label_fields(r),
                       float(r["precio_retail"]) if isinstance(r.get("precio_retail"), (int, float)) and r["precio_retail"] > 0 else None,now_text()))
                for r in parsed["Stock"]:
                    cid = db.execute("SELECT id FROM containers WHERE code=?", (str(r["codigo_contenedor"]).strip(),)).fetchone()[0]
                    pid = db.execute("SELECT id FROM products WHERE sku=?", (str(r["sku"]).strip(),)).fetchone()[0]
                    db.execute("""INSERT INTO stock(container_id,product_id,quantity) VALUES(?,?,?)
                      ON CONFLICT(container_id,product_id) DO UPDATE SET quantity=excluded.quantity""", (cid,pid,int(r["cantidad"])))
            flash(f"Importación lista: {len(parsed['Contenedores'])} contenedores, {len(parsed['Productos'])} productos, {len(parsed['Stock'])} existencias. Los códigos importados se conservaron.", "success")
            return redirect(url_for("import_data"))
        except (ValueError, KeyError, TypeError, sqlite3.IntegrityError) as exc:
            flash(str(exc), "error")
        except Exception:
            app.logger.exception("Error al procesar archivo de importación")
            flash("No se pudo leer el archivo. Descargá la plantilla y comprobá que sea un .xlsx válido.", "error")
    return render_template("import.html")


@app.get("/importar/plantilla.xlsx")
@login_required
def import_template():
    wb = Workbook()
    sheets = {
        "Contenedores": ["codigo", "nombre", "ubicacion", "sucursal"],
        "Productos": ["sku", "nombre", "item", "color", "talle", "descripcion", "unidades_sin_ubicar", "ubicaciones", "color_nombre", "temporada", "precio_retail"],
        "Stock": ["codigo_contenedor", "sku", "cantidad"],
    }
    for index, (name, headers) in enumerate(sheets.items()):
        ws = wb.active if index == 0 else wb.create_sheet()
        ws.title = name
        ws.append(headers)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="plantilla_deposito.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/exportar.xlsx")
@login_required
def export_data():
    db = get_db()
    wb = Workbook()
    for index, (name, query) in enumerate([
        ("Contenedores", "SELECT code codigo,name nombre,location ubicacion,site sucursal FROM containers c WHERE active=1 AND {own} ORDER BY code"),
        ("Productos", "SELECT sku,name nombre,item,color,size talle,description descripcion,legacy_units unidades_sin_ubicar,source_location ubicaciones,color_name color_nombre,season temporada,retail_price precio_retail FROM products ORDER BY sku"),
        ("Stock", "SELECT c.code codigo_contenedor,p.sku,s.quantity cantidad FROM stock s JOIN containers c ON c.id=s.container_id JOIN products p ON p.id=s.product_id WHERE s.quantity>0 AND {own} ORDER BY c.code,p.sku"),
    ]):
        ws = wb.active if index == 0 else wb.create_sheet()
        ws.title = name
        own, own_args = own_sql("c", strict=False)
        rows = db.execute(query.replace("{own}", own), own_args if "{own}" in query else ()).fetchall()
        ws.append(list(rows[0].keys()) if rows else {"Contenedores": ["codigo","nombre","ubicacion","sucursal"],"Productos":["sku","nombre","item","color","talle","descripcion","unidades_sin_ubicar","ubicaciones","color_nombre","temporada","precio_retail"],"Stock":["codigo_contenedor","sku","cantidad"]}[name])
        for row in rows:
            ws.append(list(row))
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="deposito_stock.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------------------
# Etiquetas de prendas (código de artículo, color, talle y código de barras)
# ---------------------------------------------------------------------------
# Patrones Code 128 (anchos de barra/espacio) para los valores 0-106; el 106 es STOP.
CODE128_PATTERNS = (
    "212222 222122 222221 121223 121322 131222 122213 122312 132212 221213 221312 231212 112232 122132 "
    "122231 113222 123122 123221 223211 221132 221231 213212 223112 312131 311222 321122 321221 312212 "
    "322112 322211 212123 212321 232121 111323 131123 131321 112313 132113 132311 211313 231113 231311 "
    "112133 112331 132131 113123 113321 133121 313121 211331 231131 213113 213311 213131 311123 311321 "
    "331121 312113 312311 332111 314111 221411 431111 111224 111422 121124 121421 141122 141221 112214 "
    "112412 122114 122411 142112 142211 241211 221114 413111 241112 134111 111242 121142 121241 114212 "
    "124112 124211 411212 421112 421211 212141 214121 412121 111143 111341 131141 114113 114311 411113 "
    "411311 113141 114131 311141 411131 211412 211214 211232 2331112"
).split()

GARMENT_LABEL_FORMATS = {
    "40x20": {"name": "4 × 2 cm", "width": 40, "height": 20, "scale": 1},
    # Etiqueta comercial: nombre, color, talle, artículo, código de color, temporada y código de barras.
    "50x30": {"name": "3 × 5 cm (completa)", "width": 50, "height": 30, "scale": 1, "layout": "full"},
    "102x51": {"name": "4 × 2 pulgadas", "width": 101.6, "height": 50.8, "scale": 2.5},
}
GARMENT_LABEL_MAX = 1000


def code128_values(text):
    """Devuelve los valores Code 128 (inicio, datos, control) para un texto ASCII imprimible."""
    if not text or any(ord(ch) < 32 or ord(ch) > 126 for ch in text):
        raise ValueError("El código solo puede tener caracteres ASCII imprimibles.")
    if text.isdigit() and len(text) >= 2:
        # Set C: dos dígitos por símbolo; si la cantidad es impar, el último va en set B.
        even = len(text) - (len(text) % 2)
        values = [105] + [int(text[i:i + 2]) for i in range(0, even, 2)]
        if even < len(text):
            values += [100, ord(text[-1]) - 32]
    else:
        values = [104] + [ord(ch) - 32 for ch in text]
    checksum = (values[0] + sum(i * v for i, v in enumerate(values[1:], start=1))) % 103
    return values + [checksum, 106]


def barcode_digits(code):
    """Texto legible bajo el código de barras; un UPC de 12 dígitos se agrupa como 1 95699 46251 9."""
    code = (code or "").strip()
    if len(code) == 12 and code.isdigit():
        return f"{code[0]} {code[1:6]} {code[6:11]} {code[11]}"
    return code


def code128_svg(text, quiet=6):
    """SVG Code 128 sin texto, escalable al tamaño de la etiqueta (barras nítidas)."""
    widths = "".join(CODE128_PATTERNS[v] for v in code128_values(text))
    x, bars = quiet, []
    for index, width in enumerate(widths):
        width = int(width)
        if index % 2 == 0:
            bars.append(f'<rect x="{x}" y="0" width="{width}" height="1"/>')
        x += width
    total = x + quiet
    return Markup(f'<svg class="barcode" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {total} 1" '
                  f'preserveAspectRatio="none" shape-rendering="crispEdges" role="img" '
                  f'aria-label="Código de barras {escape(text)}">{"".join(bars)}</svg>')


@app.get("/etiquetas-prendas")
@login_required
def garment_labels():
    code = request.args.get("codigo", "").strip()
    color = request.args.get("color", "").strip()
    label_format = request.args.get("formato", "40x20")
    if label_format not in GARMENT_LABEL_FORMATS:
        label_format = "40x20"
    rows = []
    if code:
        rows = get_db().execute("""SELECT p.id,p.sku,p.name,p.item,p.color,p.size,p.color_name,p.season,p.retail_price FROM products p
          WHERE p.active=1 AND (p.sku=? OR (TRIM(p.item)=? AND (?='' OR UPPER(TRIM(p.color))=UPPER(?))))
          ORDER BY UPPER(TRIM(p.color)),
            CASE UPPER(TRIM(p.size))
              WHEN 'XXS' THEN 1 WHEN 'XS' THEN 2 WHEN 'S' THEN 3 WHEN 'M' THEN 4
              WHEN 'L' THEN 5 WHEN 'XL' THEN 6 WHEN 'XXL' THEN 7 WHEN 'XXXL' THEN 8 ELSE 20 END,
            p.size COLLATE NOCASE,p.sku LIMIT 300""", (code, code, color, color)).fetchall()
    return render_template("garment_labels.html", results=rows, code=code, color=color,
                           searched="codigo" in request.args, label_format=label_format,
                           formats=GARMENT_LABEL_FORMATS)


@app.get("/etiquetas-prendas/imprimir")
@login_required
def print_garment_labels():
    label_format = request.args.get("formato", "40x20")
    fmt = GARMENT_LABEL_FORMATS.get(label_format, GARMENT_LABEL_FORMATS["40x20"])
    copies = {}
    for raw_id in request.args.getlist("p"):
        if not raw_id.isdigit():
            continue
        try:
            amount = int(request.args.get(f"q{raw_id}", "1"))
        except ValueError:
            amount = 1
        if amount > 0:
            copies[int(raw_id)] = min(amount, GARMENT_LABEL_MAX)
    if not copies:
        flash("Seleccioná al menos una prenda e indicá cuántas etiquetas imprimir.", "error")
        return redirect(request.referrer or url_for("garment_labels"))
    if sum(copies.values()) > GARMENT_LABEL_MAX:
        flash(f"Podés imprimir hasta {GARMENT_LABEL_MAX} etiquetas por vez.", "error")
        return redirect(request.referrer or url_for("garment_labels"))
    placeholders = ",".join("?" for _ in copies)
    products = get_db().execute(f"SELECT id,sku,name,item,color,size,color_name,season,retail_price FROM products WHERE id IN ({placeholders})",
                                tuple(copies)).fetchall()
    labels, skipped = [], []
    for product in sorted(products, key=lambda p: list(copies).index(p["id"])):
        try:
            svg = code128_svg(product["sku"].strip())
        except ValueError:
            skipped.append(product["sku"])
            continue
        price = product["retail_price"]
        labels += [{"product": product, "svg": svg, "digits": barcode_digits(product["sku"]),
                    "price": price, "net": price_without_iva(price)}] * copies[product["id"]]
    if not labels:
        flash("Los códigos seleccionados no se pueden representar en Code 128.", "error")
        return redirect(request.referrer or url_for("garment_labels"))
    return render_template("garment_label_print.html", labels=labels, fmt=fmt, skipped=skipped)


@app.errorhandler(413)
def too_large(_error):
    return "El Excel supera el límite de 12 MB.", 413


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
