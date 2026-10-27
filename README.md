# Boletero SUNAT (2 RUC) — paso a paso

## Cómo funciona
1. Lee la hoja **Ventas** de `ventas.xlsx` (1 fila = 1 boleta, `Estado` = PENDIENTE).
2. Por cada RUC (columna `Cuenta`) inicia sesión en SOL, emite la boleta y anota el número.
3. Escribe en el Excel: `Estado` (EMITIDA / ERROR / REVISAR), `Serie_Numero`, `Fecha_Emision`, `Error`.
4. La hoja **Resumen** calcula sola, por cuenta: boletas emitidas, total boleteado, pendientes y errores.

Estados: `EMITIDA` ok · `ERROR` dato inválido (corrige y pon PENDIENTE) · `REVISAR` falla técnica,
**verifica en SUNAT si salió antes de reintentar** (el bot nunca reintenta solo para no duplicar) · `EN_PROCESO` el bot se cayó a mitad.

## Pasos
1. **Instalar**: `pip install -r requirements.txt && playwright install chromium`
2. **Configurar** `config.yaml`: tus 2 RUC, usuarios SOL y nombres de cuenta.
3. **Claves SOL** como variables de entorno (nunca en archivos):
   - Windows: `setx SOL_CLAVE_EMPRESA_A "tuclave"` y `setx SOL_CLAVE_EMPRESA_B "tuclave"` (reabrir terminal)
4. **Probar en simulación** (`modo: simulacion`): `python main.py --crear-demo` y luego `python main.py`. Abre `ventas.xlsx` y revisa Resumen.
5. **Calibrar el portal SOL** (una sola vez, es la parte delicada): en `emisores.py` ajusta `SELECTORES`
   (login, ruta del menú, campos del formulario de boleta). Usa `python -m playwright codegen <url de login SOL>`
   para grabar tus clics y copiar los selectores reales.
6. **Prueba real con 1 boleta** de monto bajo: `modo: sunat`, `headless: false`, una sola fila PENDIENTE. Confirma el número en SUNAT.
7. **Subir a producción**: déjalo con varias filas; luego programa `python main.py` con el Programador de tareas de Windows (ej. cada noche).
8. **Odoo**: exporta ventas a Excel/CSV con esas columnas, o luego se conecta por la API de Odoo (XML-RPC) para llenar la hoja automáticamente.

## Riesgos a tener presentes
- El portal SOL cambia y puede pedir captcha/verificación: por eso conviene el piloto del paso 6.
- Compartir/guardar claves SOL tiene riesgo; usa un usuario secundario SOL con solo permiso de emisión.
- Alternativa más estable a futuro: API de un PSE/OSE (mismo Excel, solo cambia `SunatSolEmisor`).
- Boleta ≥ S/700 exige DNI del cliente (el bot lo valida).
- Cierra el Excel antes de ejecutar (si está abierto no se puede guardar).
