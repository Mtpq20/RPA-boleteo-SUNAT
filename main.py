"""Boletero: lee ventas.xlsx, emite boletas pendientes y registra el resultado por cuenta.
Uso:  python main.py [--crear-demo]
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
from excel_io import agrupar, crear_plantilla, leer_pendientes, marcar
from emisores import crear_emisor, ErrorValidacion

# Claves SOL desde .env (junto a este archivo); no pisa variables ya definidas en Windows
load_dotenv(Path(__file__).with_name(".env"))


def cargar_config():
    """config.yaml da la estructura; lo privado viene del .env y tiene prioridad.
    Por cuenta: <CUENTA>_RUC, <CUENTA>_USUARIO_SOL, <CUENTA>_SERIE.
    Generales: EXCEL, MODO, HEADLESS, PAUSA_ENTRE_BOLETAS_SEG."""
    cfg = yaml.safe_load(open(Path(__file__).with_name("config.yaml"), encoding="utf-8"))
    for clave in ("excel", "modo", "headless", "pausa_entre_boletas_seg", "drive_carpeta"):
        valor = os.environ.get(clave.upper())
        if valor:
            # "false" -> False, "3" -> 3; rutas y textos tal cual (G:\Mi unidad\...)
            cfg[clave] = yaml.safe_load(valor) if clave in ("headless", "pausa_entre_boletas_seg") else valor.strip()
    for nombre, c in cfg["cuentas"].items():
        for campo in ("ruc", "usuario_sol", "serie"):
            valor = os.environ.get(f"{nombre}_{campo.upper()}")
            if valor:
                c[campo] = valor.strip()
        if not re.fullmatch(r"\d{11}", str(c.get("ruc", ""))):
            sys.exit(f"{nombre}: el RUC debe tener 11 dígitos (revisa {nombre}_RUC en .env)")
        if not c.get("usuario_sol") or c["usuario_sol"].upper().startswith("USUARIO"):
            sys.exit(f"{nombre}: falta el usuario SOL (revisa {nombre}_USUARIO_SOL en .env)")
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


def main():
    cfg = cargar_config()
    ruta = cfg["excel"]

    if "--crear-demo" in sys.argv:
        demo = [
            (1, "2026-10-26", "EMPRESA_A", 1, "45123456", "Juan Pérez", "Polo algodón", 2, 59.90),
            (2, "2026-10-26", "EMPRESA_B", 0, "", "Clientes varios", "Gorra", 1, 35.00),
            (3, "2026-10-26", "EMPRESA_A", 1, "7012345", "Ana Ruiz", "Casaca (DNI mal)", 1, 189.00),
            (4, "2026-10-26", "EMPRESA_B", 1, "40998877", "Luis Soto", "Zapatillas", 1, 420.00),
            (5, "2026-10-26", "EMPRESA_A", 0, "", "Sin DNI", "Laptop (>700)", 1, 2500.00),
        ]
        crear_plantilla(ruta, list(cfg["cuentas"]), demo)
        print(f"Plantilla creada: {ruta}")
        return

    wb, ws, idx, pendientes = leer_pendientes(ruta)
    boletas = agrupar(ws, idx, pendientes)
    print(f"{len(boletas)} boleta(s) pendiente(s) ({len(pendientes)} producto(s)). Modo: {cfg['modo']}")
    emisor = crear_emisor(cfg["modo"], cfg["cuentas"], cfg["headless"])
    ok = err = 0

    def marcar_boleta(b, estado, numero=None, error=None):
        for fila in b["filas"]:   # todas las filas (productos) de la boleta quedan igual
            marcar(wb, ws, idx, ruta, fila, estado, numero, error)

    try:
        # Agrupar por cuenta para hacer login una sola vez por RUC
        for b in sorted(boletas, key=lambda x: str(x["Cuenta"])):
            nombre = f"{b['clave']} ({len(b['items'])} prod.)"
            if b["Cuenta"] not in cfg["cuentas"]:
                marcar_boleta(b, "ERROR", error=b["error"] or f"Cuenta desconocida: {b['Cuenta']}")
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


if __name__ == "__main__":
    main()
