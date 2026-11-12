# Depósito Patagonia

Primera versión de un sistema web de depósito para ejecutar como contenedor Docker y administrar desde Portainer. La información se guarda en SQLite dentro de un volumen persistente.

## Funciones incluidas

- Inicio de sesión con perfiles **Administrador** y **Solo consulta**. Los administradores crean usuarios desde **Usuarios**; la contraseña inicial debe tener al menos 12 caracteres.
- El perfil de consulta puede ver productos, ubicaciones, stock, etiquetas e historial. No puede crear o modificar productos/contendedores, importar planillas ni registrar movimientos; los permisos se validan en el servidor.
- Alta de contenedores y generación automática de un código numérico de 12 dígitos.
- Creación por lote de hasta 500 contenedores con prefijo y rango numérico, cada uno con código de barras propio.
- Desde la ficha del contenedor principal se pueden crear los racks por lote como sus subcontenedores.
- Jerarquía de contenedores: desde una ficha se pueden crear cajas/subcontenedores, cada uno con etiqueta y código propios.
- Edición de nombre y ubicación, cambio de depósito para un contenedor principal y movimiento de contenedores dentro de otros; el cambio de depósito se propaga a las cajas hijas.
- Retiro de artículos de un contenedor y eliminación de contenedores o de toda una rama, con ajuste de stock registrado en el historial.
- Etiquetas imprimibles con código de barras Code 128. Los códigos importados se conservan y no se reciclan.
- **Etiquetas de prendas** (menú debajo de Usuarios): buscá por código de artículo (`item`), opcionalmente con color, o por SKU; elegí las prendas y la cantidad de copias, e imprimí etiquetas de 4 × 2 cm (o 4 × 2 pulgadas) con código, color, talle y el SKU como código de barras Code 128. Cada etiqueta sale en una página del tamaño elegido, lista para impresoras de rollo.
- Catálogo de productos con columnas de código, nombre, item, color, talle, descripción, unidades y ubicación. El listado permite ver Bariloche, Retiro o ambos; no incluye Centro de Distribución.
- Buscador de talles por código de artículo (`item`) o código de barras, combinado con color; muestra todos los talles, ubicaciones y cantidades físicas.
- Bot de Telegram para consultar talles y ubicaciones enviando `CÓDIGO COLOR` o `/buscar CÓDIGO COLOR`.
- El catálogo permite archivar referencias que no se tienen y restaurarlas si hace falta.
- Stock por contenedor, consulta y búsqueda por ubicación.
- Lector de código por teclado USB (los lectores configurados como teclado suelen enviar el código y Enter).
- Ingresos, salidas y traslados, con control de cantidad disponible y registro de movimientos.
- Flujo PICK para traslados por lote: escaneá el código del contenedor de origen y destino, luego las etiquetas de varias prendas; cada lectura suma una unidad y el lote se aplica al confirmar.
- Depósito Virtual aislado, creado automáticamente: recibe el catálogo y las cantidades lógicas importadas; cada PICK descuenta ese saldo y suma stock únicamente al rack físico de destino.
- Importación y exportación de contenedores, productos y stock en Excel.

El importador admite la planilla original de productos con una hoja `elementos` (con o sin la columna opcional `ubicaciones`) y la plantilla de tres hojas. En la hoja `elementos`, `codigo` se conserva como SKU de la prenda y los productos se asocian al Depósito Virtual. La columna `unidades` queda como saldo lógico disponible para PICK, separado del stock físico; cada movimiento descuenta ese saldo y suma las unidades al rack de destino. La ficha del Depósito Virtual muestra SKU, código de artículo, color, talle y saldo disponible. La importación no carga el historial anterior de movimientos. La importación hace altas/actualizaciones de los códigos incluidos y conserva los códigos existentes; las filas que no aparezcan en un nuevo archivo no se eliminan automáticamente.

## Despliegue desde Portainer

### 1. Preparar la imagen

Copiá la carpeta `deposito-stock` a un equipo con acceso al Docker Engine del servidor. Desde esa carpeta, construí la imagen:

```sh
docker build -t deposito-stock:1.0 .
```

Si el comando se ejecuta desde otro equipo, ese mismo Docker Engine debe ser el del servidor o la imagen tiene que publicarse en un registry accesible por Portainer.

### 2. Crear el Stack

En Portainer, creá un Stack y pegá el contenido de `compose-portainer.yaml`. En **Environment variables**, definí:

- `SECRET_KEY`: valor aleatorio largo usado para firmar sesiones.
- `ADMIN_USERNAME`: usuario inicial (por defecto `admin`).
- `ADMIN_PASSWORD`: contraseña inicial larga.
- `PORT`: puerto HTTP publicado (por defecto `8001`).
- `SESSION_COOKIE_SECURE`: dejalo `false` si entrás por HTTP directo. Ponelo en `true` cuando el sitio quede detrás de HTTPS.
- `TELEGRAM_BOT_TOKEN`: token privado creado con `@BotFather`; nunca lo publiques en GitHub.
- `TELEGRAM_ALLOWED_USER_IDS`: IDs de Telegram autorizados, separados por coma. Dejalo vacío al primer deploy para obtener tu ID con `/id`; luego completalo y redeployá.

Portainer crea el volumen `deposito_data`; ahí persiste la base SQLite aunque se reemplace el contenedor. No borres ese volumen al actualizar.

### 2.1. Vincular el bot de Telegram

1. En Telegram, abrí `@BotFather`, enviá `/newbot` y seguí los pasos para recibir el token.
2. En las variables del Stack de Portainer, cargá `TELEGRAM_BOT_TOKEN` y dejá por ahora vacío `TELEGRAM_ALLOWED_USER_IDS`. Redesplegá el Stack.
3. Abrí un chat privado con el bot, tocá **Start** y enviá `/id`. El bot devuelve tu user ID, pero todavía no consulta el inventario.
4. En Portainer, agregá ese número a `TELEGRAM_ALLOWED_USER_IDS` y redesplegá.
5. En el chat privado, enviá el código de artículo (`item`) y color, por ejemplo `84628 BLK`, o `/buscar 84628 BLK` para listar todos los talles. Para una prenda puntual, enviá su SKU/código de barras solo, sin color.

El bot usa long polling, comparte la base SQLite del Stack y requiere salida a Internet desde el contenedor. No publiques el token ni lo agregues a GitHub. El bot usa `getUpdates`; si ese bot ya tiene un webhook configurado, eliminá el webhook antes de iniciar este servicio, porque Telegram no permite usar ambos métodos a la vez.

### 3. Primer ingreso

Abrí `http://IP-DEL-SERVIDOR:8001` (o el puerto definido), iniciá sesión con el usuario y contraseña configurados en el Stack y cambiá la contraseña en **Cuenta**. Para exponerlo fuera de la red local, configurá HTTPS en el proxy inverso antes de usar `SESSION_COOKIE_SECURE=true`.

## Alternativa con Compose

Copiá `.env.example` a `.env`, reemplazá los secretos y ejecutá:

```sh
docker compose up -d --build
```

## Importación de datos

Desde **Importar Excel**, descargá `plantilla_deposito.xlsx`. El archivo contiene tres hojas con estas columnas:

- `Contenedores`: `codigo`, `nombre`, `ubicacion`, `sucursal` (Bariloche o Retiro)
- `Productos`: `sku`, `nombre`, `item`, `color`, `talle`, `descripcion`, `unidades_sin_ubicar`, `ubicaciones`
- `Stock`: `codigo_contenedor`, `sku`, `cantidad`

La hoja `elementos` del Excel original se transforma directamente: `codigo` pasa a ser el código del producto, y se conservan nombre, item, color, talle, descripción, unidades y ubicaciones.

Las filas de stock relacionan cada SKU con un contenedor mediante sus códigos. Usá códigos de barras existentes en la columna `codigo` para mantenerlos. La sucursal debe ser Bariloche o Retiro. La aplicación guarda los códigos como texto y los representa en Code 128.

## Archivos

- `app.py`: servidor y reglas del inventario.
- `templates/` y `static/`: interfaz web.
- `compose.yaml`: Stack que construye la imagen desde la carpeta.
- `docker-compose.yml`: nombre predeterminado para la ruta Compose de Portainer; construye la imagen desde el repositorio.
- `compose.yaml`: misma configuración, con nombre Compose moderno.
- `compose-portainer.yaml`: Stack para Portainer usando una imagen ya construida.

## Usuarios y permisos

El usuario inicial configurado con `ADMIN_USERNAME` se crea como administrador. Si se cambió esa variable después de la instalación y ya no hay administradores, el sistema recupera el acceso de administrador para la cuenta más antigua. Después del ingreso, abrí **Usuarios** para crear otros accesos y elegir si serán administradores o de solo consulta. Cada cuenta usa su propio nombre de usuario y contraseña; se recomienda cambiar la contraseña inicial desde **Cuenta**. La base existente se migra automáticamente al iniciar, sin borrar sus datos.

- **Administrador:** puede operar el depósito y administrar usuarios.
- **Solo consulta:** puede consultar el inventario, las ubicaciones, los productos y el historial.

La consulta por código busca el código del contenedor o SKU exacto. El lector funciona como teclado USB.
- La impresión de etiquetas usa una medida inicial de 88 × 50 mm; puede ajustarse a la impresora y etiquetas disponibles.
- Los códigos nuevos son números secuenciales de 12 cifras impresos en Code 128. Si los códigos del sistema anterior usan otra simbología o longitud, el importador mantiene su texto, pero conviene comparar una etiqueta antes de imprimir lotes.
- SQLite funciona bien para un depósito chico. La base se mantiene en un único volumen para facilitar copias de seguridad y restauración.
