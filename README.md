# Boletero SUNAT — RPA para emitir boletas electrónicas

Bot (RPA) en Python que **emite boletas de venta electrónicas en el portal SUNAT SOL (SEE-SOL)** a partir de
un Excel de ventas, para **varias empresas (varios RUC)**, y lleva el registro de lo boleteado por empresa
y por mes, con un **tope mensual** por empresa.

- Lee las ventas pendientes de `ventas.xlsx` (escritas a mano o importadas desde **Odoo**).
- Entra a SOL con el RUC de cada empresa, llena la boleta, **verifica el importe** y la emite.
- Anota en el Excel el número de boleta (EB01-…), el estado y la fecha de emisión.
- Guarda el PDF/XML de cada boleta y (opcional) copia el PDF a **Google Drive**.
- No emite una boleta que haría pasar el **tope mensual** de la empresa (S/ 5,000 por defecto).

---

## ¿Es solo para 4 empresas?

**No.** Viene configurado con 4 empresas (`EMPRESA_A` … `EMPRESA_D`), pero el código no tiene un número
fijo de empresas: **todo sale de `config.yaml` y `.env`**. Se pueden **agregar, renombrar o desactivar
empresas sin modificar el código** (probado agregando una 5.ª empresa: tuvo su hoja, su columna en
"Por mes", su tope y emitió normalmente).

### Agregar una empresa (ej. la 5.ª)
1. En `config.yaml`, dentro de `cuentas:`, agrega un bloque:
   ```yaml
     EMPRESA_E:
       clave_env: "SOL_CLAVE_EMPRESA_E"
       serie: "B001"
       odoo: false          # quítalo si sus ventas vienen de Odoo
   ```
2. En `.env`, agrega sus datos (ver la sección [.env](#2-crear-el-archivo-env)):
   ```
   EMPRESA_E_NOMBRE=Nombre comercial
   EMPRESA_E_RUC=20XXXXXXXXX
   EMPRESA_E_USUARIO_SOL=USUARIO
   SOL_CLAVE_EMPRESA_E=clave_sol
   ```
3. Ejecuta el bot: el Excel crea solo la hoja de la nueva empresa y su columna en "Por mes".

- **Renombrar** una empresa: cambia `EMPRESA_X_NOMBRE` en `.env`. La hoja cambia de nombre y no se pierde nada.
- **Desactivar** una empresa: deja vacíos su RUC y usuario en `.env`. Sus ventas se quedan en PENDIENTE.
- **Tope distinto** para una empresa: agrega `limite_mensual: 3000` dentro de su bloque en `config.yaml`.

---

## Archivos

| Archivo | Qué es | ¿Se sube a GitHub? |
|---|---|---|
| `main.py` | Programa principal: lee pendientes, emite y anota resultados | ✅ |
| `emisores.py` | Automatización del portal SUNAT SOL (Playwright) | ✅ |
| `excel_io.py` | Lectura/escritura del Excel: hojas por empresa, "Por mes", tope | ✅ |
| `importar_odoo.py` | Importa a `ventas.xlsx` las órdenes exportadas de Odoo | ✅ |
| `config.yaml` | Empresas, tope mensual, modo por defecto (sin datos privados) | ✅ |
| `.env.example` | Plantilla del `.env` con datos de ejemplo | ✅ |
| `.env` | **RUC, usuarios y claves SOL reales** | ❌ nunca |
| `ventas.xlsx` | Ventas y boletas (tiene datos de clientes) | ❌ |
| `comprobantes/`, `capturas/` | PDF/XML de boletas y capturas del portal | ❌ |

---

## Instalación

Requiere **Windows** y **Python 3.11**.

### 1. Instalar dependencias (en la carpeta del proyecto)
```bash
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\playwright install chromium
```

### 2. Crear el archivo `.env`
Copia `.env.example` como `.env` y completa los datos reales. **El `.env` nunca se sube a GitHub**
(está en `.gitignore`) y las claves SOL no se envían por WhatsApp ni correo.

**Generales**

| Variable | Qué es | Ejemplo |
|---|---|---|
| `MODO` | `simulacion` (no abre SUNAT) · `ensayo` (llena en SUNAT pero **no emite**) · `sunat` (**emite de verdad**) | `ensayo` |
| `EXCEL` | Archivo de ventas | `ventas.xlsx` |
| `HEADLESS` | `false` = ves el navegador trabajar · `true` = trabaja oculto | `false` |
| `PAUSA_ENTRE_BOLETAS_SEG` | Segundos de espera entre boletas | `3` |
| `DRIVE_CARPETA` | Carpeta de Google Drive para copiar los PDF (vacío = no copiar) | `G:\Mi unidad\Boletas` |

**Por cada empresa** (cambia `A` por la letra de la empresa: `A`, `B`, `C`, `D`…)

| Variable | Qué es |
|---|---|
| `EMPRESA_A_NOMBRE` | Nombre comercial: es el nombre de su hoja y de su columna en "Por mes" |
| `EMPRESA_A_RUC` | RUC de 11 dígitos |
| `EMPRESA_A_USUARIO_SOL` | Usuario SOL de ese RUC |
| `SOL_CLAVE_EMPRESA_A` | Clave SOL |
| `EMPRESA_A_SERIE` | Serie (solo se usa en simulación; en SUNAT la serie la asigna el portal, p. ej. `EB01`) |

Lo del `.env` tiene prioridad sobre `config.yaml`.

### 3. El Excel `ventas.xlsx`
No está en GitHub. Para continuar un registro existente, copia el `ventas.xlsx` de la otra PC a la carpeta
del proyecto (con ese nombre exacto). Para empezar uno vacío:
```bash
.venv\Scripts\python -c "import main; from excel_io import crear_plantilla, limites_de; cfg=main.cargar_config(); crear_plantilla(cfg['excel'], list(cfg['cuentas']), [], limites_de(cfg))"
```
> ⚠️ No uses `python main.py --crear-demo` con `MODO=sunat`: crea ventas de ejemplo PENDIENTE que se emitirían de verdad.

### 4. Probar sin emitir
Pon `MODO=ensayo` y ejecuta el bot: entra a SOL con cada empresa y llena la boleta hasta la vista previa
**sin emitirla**, y guarda una captura en `capturas/`. Si todo sale bien, cambia a `MODO=sunat`.

---

## Uso diario

### Registrar ventas
Abre `ventas.xlsx` y escribe la venta en la **hoja de la empresa** (una pestaña por empresa):

| Columna | Qué poner |
|---|---|
| `Fecha` | Día de la venta (si la dejas vacía se pone la del día) |
| `Nro_Doc` | DNI (8 dígitos) o RUC (11) del cliente; vacío si no tiene |
| `Cliente` | Nombre (con DNI/RUC, SUNAT pone el nombre registrado) |
| `Descripcion` | Producto |
| `Cantidad` / `Precio_Unit` | Unidades y precio unitario **con IGV**, con punto decimal (`1.50`) |
| `Total` | Se calcula solo |
| `Estado` | Vacío o `PENDIENTE` |
| `Ref_Odoo` | Solo si la boleta lleva **varios productos**: el mismo código en cada fila (ej. `VENTA-001`) |

- **Un producto por fila.** Las filas con el mismo `Ref_Odoo` salen en **una sola boleta**.
- Las columnas `Serie_Numero`, `Fecha_Emision` y `Error` las llena el bot.
- **La columna A (ID) está oculta:** no la borres, une cada fila con su registro.
- Boleta de **S/ 700 o más** requiere DNI.

### Emitir
Con el Excel **cerrado**:
```bash
.venv\Scripts\python main.py
```

### Importar ventas de Odoo
En Odoo: *Punto de venta → Órdenes → Exportar* (plantilla "Boletas"), y luego:
```bash
.venv\Scripts\python importar_odoo.py "archivo exportado.xlsx" --prueba   # solo muestra
.venv\Scripts\python importar_odoo.py "archivo exportado.xlsx"            # importa
```
En `config.yaml` → `odoo: puntos_de_venta` se indica a qué empresa va cada punto de venta.

---

## El Excel

| Hoja | Para qué |
|---|---|
| **Por mes** | Se abre primero. Boleteado por mes y empresa, tope mensual, disponible y alerta (verde / naranja desde el 80 % / rojo al llegar al tope) |
| **Una hoja por empresa** | Donde se escriben y consultan las ventas de esa empresa |
| `Ventas` *(oculta)* | Registro completo de todas las empresas; es lo que usa el bot |
| `Resumen` *(oculta)* | Totales por empresa |

Las hojas ocultas se ven con clic derecho en una pestaña → *Mostrar*.

### Estados de una venta

| Estado | Significado | Qué hacer |
|---|---|---|
| vacío / `PENDIENTE` | Por emitir | Nada: se emite en la siguiente ejecución |
| `EMITIDA` | Boleta emitida (número en `Serie_Numero`) | — |
| `ERROR` | **No se emitió** (DNI inválido, tope del mes, falta un dato…). Motivo en `Error` | Corrige el dato: se reintenta sola |
| `REVISAR` | Falla después de presionar Emitir: **pudo haberse emitido** | Revisa en SUNAT antes de reintentar (el bot nunca la reintenta solo) |
| `EN_PROCESO` | El bot se cortó a mitad de esa boleta | Igual que REVISAR |

---

## Precauciones

- **Un solo PC debe emitir** con `MODO=sunat`. Cada PC tiene su propio `ventas.xlsx`: si dos emiten la misma
  venta, sale dos veces en SUNAT y el tope no lo detecta.
- **Cierra el Excel** antes de ejecutar: si está abierto el bot no puede guardar (espera y reintenta).
- La fecha de la boleta es **la del día en que se emite**. Conviene ejecutar el bot el mismo día de la venta.
- El portal SOL puede cambiar o pedir captcha: si algo falla antes de emitir, la venta queda en `ERROR`
  con una captura en `capturas/`; lo que depende del portal está concentrado en `emisores.py` (`SELECTORES`).
- Usa de preferencia un usuario SOL secundario con permiso solo para emitir.
