"""Telegram inventory search bot using long polling and the shared SQLite database."""

import json
import os
import sqlite3
import time
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
DATABASE_PATH = os.environ.get("DATABASE_PATH", "/data/deposito.sqlite3")
ALLOWED_USER_IDS = {
    value.strip() for value in os.environ.get("TELEGRAM_ALLOWED_USER_IDS", "").split(",") if value.strip()
}
API = f"https://api.telegram.org/bot{TOKEN}/"


def api(method, data):
    request = Request(
        API + method,
        data=json.dumps(data).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urlopen(request, timeout=45) as response:
            result = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        try:
            description = json.loads(exc.read().decode("utf-8")).get("description", "")
        except Exception:
            description = ""
        detail = f": {description[:180]}" if description else ""
        raise RuntimeError(f"Telegram API HTTP {exc.code}{detail}") from None
    except (URLError, TimeoutError) as exc:
        raise RuntimeError(f"Telegram connection error: {type(exc).__name__}") from None
    if not result.get("ok"):
        description = str(result.get("description", ""))[:180]
        raise RuntimeError(f"Telegram API rejected the request: {description}")
    return result.get("result")


def send(chat_id, text):
    api("sendMessage", {"chat_id": chat_id, "text": text})


def search(code, color=None):
    with sqlite3.connect(DATABASE_PATH, timeout=20) as db:
        db.row_factory = sqlite3.Row
        sql = """SELECT p.sku,p.name,p.item,p.color,p.size,
          COALESCE(SUM(CASE WHEN c.is_virtual=0 AND c.active=1 THEN s.quantity ELSE 0 END),0) physical_units,
          GROUP_CONCAT(CASE
            WHEN c.is_virtual=1 AND c.active=1 THEN 'DEPOSITO VIRTUAL · listo para PICK'
            WHEN c.is_virtual=0 AND c.active=1 AND s.quantity>0 THEN c.name || ' (' || s.quantity || ')'
          END, ' · ') locations
          FROM products p
          LEFT JOIN stock s ON s.product_id=p.id
          LEFT JOIN containers c ON c.id=s.container_id
          WHERE p.active=1 AND {condition}
          GROUP BY p.id
          ORDER BY CASE UPPER(TRIM(p.size))
            WHEN 'XXS' THEN 1 WHEN 'XS' THEN 2 WHEN 'S' THEN 3 WHEN 'M' THEN 4
            WHEN 'L' THEN 5 WHEN 'XL' THEN 6 WHEN 'XXL' THEN 7 WHEN 'XXXL' THEN 8 ELSE 20 END,
            p.size COLLATE NOCASE,p.sku"""
        if color:
            condition = "((TRIM(p.item)=? AND UPPER(TRIM(p.color))=UPPER(?)) OR p.sku=?)"
            params = (code, color, code)
        else:
            condition = "p.sku=?"
            params = (code,)
        return db.execute(sql.format(condition=condition), params).fetchall()


def answer(text):
    clean = text.strip()
    if clean.startswith("/"):
        parts = clean.split(maxsplit=1)
        command = parts[0].split("@", 1)[0].lower()
        if command == "/start" or command == "/help":
            return "Buscador del depósito.\nPor artículo y color: enviá CÓDIGO COLOR (ej. 84628 BLK) para listar talles.\nPor prenda: enviá el SKU/código de barras solo.\nTambién podés usar /buscar."
        if command == "/buscar":
            clean = parts[1] if len(parts) > 1 else ""
        else:
            return "No reconozco ese comando. Enviá CÓDIGO COLOR, por ejemplo: 84628 BLK."
    pieces = clean.replace(",", " ").split()
    if len(pieces) not in (1, 2):
        return "Enviá SKU/código de barras solo, o código de artículo y color. Ejemplos: 0198077721550 · 84628 BLK"
    code = pieces[0]
    color = pieces[1] if len(pieces) == 2 else None
    results = search(code, color)
    if not results:
        description = f"código {code}" + (f" y color {color}" if color else "")
        return f"No encontré prendas para {description}. Si buscás un artículo, usá item + color; para una prenda puntual, enviá su SKU completo."
    first = results[0]
    lines = [f"{first['name']}", f"Artículo: {first['item']} · Color: {first['color']}", "Talles y ubicaciones:"]
    for row in results[:35]:
        location = row["locations"] or "Sin ubicación física"
        lines.append(f"• {row['size'] or 'Sin talle'} · {row['physical_units']} u. físicas · {location}")
    if len(results) > 35:
        lines.append(f"… y {len(results) - 35} talles más.")
    output = "\n".join(lines)
    return output if len(output) <= 4000 else output[:3950] + "\n… resultado recortado."


def handle(update):
    message = update.get("message") or update.get("edited_message")
    if not message or not message.get("text"):
        return
    chat = message.get("chat") or {}
    user = message.get("from") or {}
    chat_id = chat.get("id")
    user_id = str(user.get("id", ""))
    text = message["text"].strip()
    command = text.split(maxsplit=1)[0].split("@", 1)[0].lower() if text else ""
    if chat.get("type") != "private":
        send(chat_id, "Por privacidad, usá este buscador en un chat privado con el bot.")
        return
    if command == "/id":
        send(chat_id, f"Tu user ID es: {user_id}\nEste es el número que el administrador debe autorizar en Portainer.")
        return
    if user_id not in ALLOWED_USER_IDS:
        send(chat_id, "Este usuario no está habilitado para consultar el inventario. Pedile al administrador que autorice tu user ID.")
        return
    try:
        send(chat_id, answer(text))
    except sqlite3.Error:
        send(chat_id, "No pude consultar el inventario ahora. Probá de nuevo en unos segundos.")


def main():
    if not TOKEN:
        raise SystemExit("Configurá TELEGRAM_BOT_TOKEN en Portainer.")
    print("Telegram inventory bot polling started.", flush=True)
    offset = None
    while True:
        try:
            payload = {"timeout": 30, "allowed_updates": ["message"]}
            if offset is not None:
                payload["offset"] = offset
            updates = api("getUpdates", payload) or []
            for update in updates:
                try:
                    handle(update)
                except Exception as exc:
                    print(f"Could not process Telegram update: {exc}.", flush=True)
                offset = update["update_id"] + 1
        except Exception as exc:
            print(f"Telegram polling error: {exc}; retrying in 5 seconds.", flush=True)
            time.sleep(5)


if __name__ == "__main__":
    main()
