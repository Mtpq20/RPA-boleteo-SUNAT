"""Boletero: lee ventas.xlsx, emite boletas pendientes y registra el resultado por cuenta.
Uso:  python main.py                corrida normal (según MODO del .env)
      python main.py --demo         demostración: crea demo_presentacion.xlsx con ventas de las
                                    4 empresas y la procesa en simulación (no toca SUNAT ni ventas.xlsx)
      python main.py --crear-demo   solo crea la plantilla de ejemplo en el Excel configurado
"""
import os
import re
import shutil
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path
import yaml
from dotenv import load_dotenv
from excel_io import agrupar, crear_plantilla, guardar, leer_pendientes, marcar
from emisores import crear_emisor, ErrorValidacion

# Claves SOL desde .env (junto a este archivo); no pisa variables ya definidas en Windows
load_dotenv(Path(__file__).with_name(".env"))


def cargar_config():
    """config.yaml da la estructura; lo privado viene del .env y tiene prioridad.
    Por cuenta: <CUENTA>_RUC, <CUENTA>_USUARIO_SOL, <CUENTA>_SERIE.
    Generales: EXCEL, MODO, HEADLESS, PAUSA_ENTRE_BOLETAS_SEG.
    Una cuenta sin RUC/usuario queda desactivada (cfg["activas"]); un RUC mal escrito detiene todo."""
    cfg = yaml.safe_load(Path(__file__).with_name("config.yaml").read_text(encoding="utf-8"))
    for clave in ("excel", "modo", "headless", "pausa_entre_boletas_seg", "drive_carpeta"):
        valor = os.environ.get(clave.upper())
        if valor:
            # "false" -> False, "3" -> 3; rutas y textos tal cual (G:\Mi unidad\...)
            cfg[clave] = yaml.safe_load(valor) if clave in ("headless", "pausa_entre_boletas_seg") else valor.strip()
    cfg["activas"] = set()
    for nombre, c in cfg["cuentas"].items():
        for campo in ("ruc", "usuario_sol", "serie"):
            valor = os.environ.get(f"{nombre}_{campo.upper()}")
            if valor:
                c[campo] = valor.strip()
        ruc, usuario = str(c.get("ruc") or ""), str(c.get("usuario_sol") or "")
        if not ruc and not usuario:
            if cfg["modo"] != "simulacion":
                print(f"Aviso: {nombre} sin RUC/usuario en .env: desactivada (sus ventas quedan PENDIENTE)")
                continue
        elif not re.fullmatch(r"\d{11}", ruc):
            sys.exit(f"{nombre}: el RUC debe tener 11 dígitos (revisa {nombre}_RUC en .env)")
        elif not usuario or usuario.upper().startswith("USUARIO"):
            sys.exit(f"{nombre}: falta el usuario SOL (revisa {nombre}_USUARIO_SOL en .env)")
        cfg["activas"].add(nombre)
    return cfg


MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def copiar_a_drive(cfg, pdf, b, numero, fecha=None):
    """Copia el PDF a la carpeta de Google Drive (sincronizada por Drive para escritorio),
    en una subcarpeta por mes: <drive_carpeta>/2026-10 Octubre/EB01-168 EMPRESA_A Yesenia Peralta.pdf.
    La boleta ya está emitida: un fallo aquí solo avisa, no cambia su estado."""
    carpeta = cfg.get("drive_carpeta")
    if not carpeta:
        return
    if not pdf or not Path(pdf).exists():
        print(f"       ! {numero}: no hay PDF para subir a Drive")
        return
    try:
        fecha = fecha or datetime.now()
        destino = Path(carpeta) / f"{fecha:%Y-%m} {MESES[fecha.month - 1]}"
        destino.mkdir(parents=True, exist_ok=True)
        cliente = re.sub(r'[\\/:*?"<>|]', "", str(b.get("Cliente") or "Sin nombre")).strip()
        extra = "_captura" if Path(pdf).stem.endswith("_captura") else ""
        final = destino / f"{numero} {b['Cuenta']} {cliente}{extra}.pdf"
        shutil.copy2(pdf, final)
        print(f"       Drive: {final}")
    except Exception as e:
        print(f"       ! {numero}: no se pudo copiar a Drive ({type(e).__name__}: {e})")


# Ventas de ejemplo para la demostración: (ID, Fecha, Cuenta, Tipo_Doc, Nro_Doc, Cliente,
# Descripcion, Cantidad, Precio_Unit[, Ref_Odoo]). Misma Ref_Odoo = una boleta con varios productos.
DEMO_PRESENTACION = [
    (1, "2026-11-10", "EMPRESA_A", 1, "45123456", "Juan Pérez", "Polo algodón", 2, 59.90, "DEMO-A-001"),
    (2, "2026-11-10", "EMPRESA_A", None, None, None, "Gorra bordada", 1, 35.00, "DEMO-A-001"),
    (3, "2026-11-10", "EMPRESA_A", None, None, None, "Medias pack x3", 1, 19.90, "DEMO-A-001"),
    (4, "2026-11-10", "EMPRESA_A", 0, "", "Clientes varios", "Llavero", 1, 8.00),
    (5, "2026-11-10", "EMPRESA_B", 1, "40998877", "Luis Soto", "Libro de inversión", 1, 45.00, "DEMO-B-001"),
    (6, "2026-11-10", "EMPRESA_B", None, None, None, "Separador", 2, 3.50, "DEMO-B-001"),
    (7, "2026-11-10", "EMPRESA_B", 0, "", "Sin DNI", "Laptop (>700, debe fallar)", 1, 2500.00),
    (8, "2026-11-10", "EMPRESA_C", 1, "70123456", "Ana Ruiz", "Pantalón jean", 1, 89.00, "DEMO-C-001"),
    (9, "2026-11-10", "EMPRESA_C", None, None, None, "Correa de cuero", 1, 39.00, "DEMO-C-001"),
    (10, "2026-11-10", "EMPRESA_C", 1, "7012345", "Rosa Díaz", "Casaca (DNI sin el 0 inicial, se corrige solo)", 1, 189.00),
    (11, "2026-11-10", "EMPRESA_D", 0, "", "Clientes varios", "Antifaz", 3, 9.90, "DEMO-D-001"),
    (12, "2026-11-10", "EMPRESA_D", None, None, None, "Disfraz charleston", 1, 79.00, "DEMO-D-001"),
    (13, "2026-11-10", "EMPRESA_D", 1, "4512345A", "Pedro Ríos", "Bisutería (DNI con letra, debe fallar)", 1, 25.00),
]


def main():
    demo_presentacion = "--demo" in sys.argv
    if demo_presentacion:
        # Las variables de entorno tienen prioridad sobre el .env: simulación y Excel aparte
        os.environ["MODO"] = "simulacion"
        os.environ["EXCEL"] = str(Path(__file__).with_name("demo_presentacion.xlsx"))
        os.environ["DRIVE_CARPETA"] = ""
    cfg = cargar_config()
    ruta = cfg["excel"]
    if demo_presentacion:
        cfg["drive_carpeta"] = None
        crear_plantilla(ruta, list(cfg["cuentas"]), DEMO_PRESENTACION)
        print(f"DEMOSTRACIÓN (simulación, no se conecta a SUNAT). Excel: {ruta}\n")

    if "--crear-demo" in sys.argv:
        demo = [
            (1, "2026-10-26", "EMPRESA_A", 1, "45123456", "Juan Pérez", "Polo algodón", 2, 59.90),
            (2, "2026-10-26", "EMPRESA_B", 0, "", "Clientes varios", "Gorra", 1, 35.00),
            (3, "2026-10-26", "EMPRESA_A", 1, "7012345", "Ana Ruiz", "Casaca (DNI mal)", 1, 189.00),
            (4, "2026-10-26", "EMPRESA_B", 1, "40998877", "Luis Soto", "Zapatillas", 1, 420.00),
            (5, "2026-10-26", "EMPRESA_A", 0, "", "Sin DNI", "Laptop (>700)", 1, 2500.00),
            (6, "2026-10-26", "EMPRESA_C", 1, "45123456", "Juan Pérez", "Pantalón", 1, 89.00),
            (7, "2026-10-26", "EMPRESA_D", 0, "", "Clientes varios", "Medias", 3, 9.90),
        ]
        crear_plantilla(ruta, list(cfg["cuentas"]), demo)
        print(f"Plantilla creada: {ruta}")
        return

    wb, ws, idx, pendientes, filas_por_ref = leer_pendientes(ruta, list(cfg["cuentas"]))
    guardar(wb, ruta)   # deja al día filtros/lista/Resumen y confirma que el Excel no está abierto
    boletas = agrupar(pendientes, filas_por_ref)
    inactivas = [b for b in boletas if b["Cuenta"] in cfg["cuentas"] and b["Cuenta"] not in cfg["activas"]]
    boletas = [b for b in boletas if b not in inactivas]
    print(f"{len(boletas)} boleta(s) pendiente(s) ({len(pendientes)} producto(s)). Modo: {cfg['modo']}")
    for cuenta in sorted({b["Cuenta"] for b in inactivas}):
        print(f"  --   {cuenta} desactivada: {sum(b['Cuenta'] == cuenta for b in inactivas)} boleta(s) se quedan PENDIENTE")
    if not boletas:
        return
    emisor = crear_emisor(cfg["modo"], cfg["cuentas"], cfg["headless"])
    ok = err = 0

    def marcar_boleta(b, estado, numero=None, error=None):
        for fila in b["filas"]:   # todas las filas (productos) de la boleta quedan igual
            marcar(ws, idx, fila, estado, numero, error)
        guardar(wb, ruta)   # una vez por boleta: si el bot se cae, no se pierde nada

    try:
        # Agrupar por cuenta para hacer login una sola vez por RUC
        for b in sorted(boletas, key=lambda x: x["Cuenta"]):
            nombre = f"{b['clave']} ({len(b['items'])} prod.)"
            if b["Cuenta"] not in cfg["cuentas"]:
                marcar_boleta(b, "ERROR", error=b["error"] or f"Cuenta desconocida: {b['Cuenta']!r}")
                err += 1
                print(f"  ERR  {nombre}: cuenta {b['Cuenta']!r}")
                continue
            # Marcar antes de emitir: si el bot se cae a mitad, no se duplica la boleta
            marcar_boleta(b, "EN_PROCESO")
            try:
                numero = emisor.emitir(b)
                if cfg["modo"] == "ensayo":
                    # No se emitió nada: las filas siguen pendientes para la corrida real
                    marcar_boleta(b, "PENDIENTE", error=numero)
                    print(f"  ENS  {nombre} {b['Cuenta']}: {numero}")
                    continue
                marcar_boleta(b, "EMITIDA", numero)
                ok += 1
                print(f"  OK   {nombre} {b['Cuenta']} {numero} S/ {b['Total']:.2f}")
                if cfg["modo"] != "simulacion":
                    copiar_a_drive(cfg, getattr(emisor, "ultimo_pdf", None), b, numero)
            except ErrorValidacion as e:
                marcar_boleta(b, "ERROR", error=str(e))
                err += 1
                print(f"  ERR  {nombre}: {e}")
            except Exception as e:
                # Error técnico: no sabemos si SUNAT la emitió. Se deja REVISAR, nunca reintenta solo.
                marcar_boleta(b, "REVISAR", error=f"{type(e).__name__}: {e}")
                err += 1
                print(f"  REV  {nombre}: {e}")
                traceback.print_exc()
            time.sleep(cfg["pausa_entre_boletas_seg"] if cfg["modo"] != "simulacion" else 0)
    finally:
        emisor.cerrar()
    print(f"Listo. Emitidas: {ok} | Con problema: {err}")
    if demo_presentacion:
        resumen_por_empresa(boletas, ws, idx)
        try:
            os.startfile(ruta)   # abre el Excel para mostrar Estado, filtros y la hoja Resumen
        except (AttributeError, OSError):
            pass


def resumen_por_empresa(boletas, ws, idx):
    """Tabla final en pantalla: qué boleta salió para cada empresa y con cuántos productos."""
    print("\nRESUMEN POR EMPRESA")
    print(f"  {'Empresa':<10} {'Boleta':<16} {'Prod.':>5} {'Total S/':>10}  Resultado")
    for b in sorted(boletas, key=lambda x: (x["Cuenta"], x["fila"])):
        celda = lambda col: ws.cell(b["fila"], idx[col] + 1).value
        estado, numero = celda("Estado"), celda("Serie_Numero") or "-"
        detalle = estado if estado == "EMITIDA" else f"{estado}: {celda('Error')}"
        print(f"  {b['Cuenta']:<10} {numero:<16} {len(b['items']):>5} {b['Total']:>10.2f}  {detalle}")


if __name__ == "__main__":
    main()
