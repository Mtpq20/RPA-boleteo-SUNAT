"""Lectura/escritura del Excel. Una fila de 'Ventas' = un producto vendido.
Las filas con la misma Ref_Odoo forman UNA boleta (una orden con varios productos);
una fila sin Ref_Odoo es una boleta de un solo producto."""
import hashlib
import time
from copy import copy
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

COLUMNAS = ["ID", "Fecha", "Cuenta", "Tipo_Doc", "Nro_Doc", "Cliente", "Descripcion",
            "Cantidad", "Precio_Unit", "Total", "Estado", "Serie_Numero", "Fecha_Emision", "Error", "Ref_Odoo"]
# Tipo_Doc: 1=DNI, 0=Sin documento (solo si total < S/700), 6=RUC
#   Si Tipo_Doc está vacío se deduce de Nro_Doc: vacío -> 0 (solo nombre), 8 dígitos -> DNI, 11 -> RUC.
# Precio_Unit incluye IGV.
CAMPOS_VENTA = ["Cuenta", "Tipo_Doc", "Nro_Doc", "Cliente", "Descripcion", "Cantidad", "Precio_Unit"]
FILAS_MAX = 5000   # alcance de la lista desplegable y de las fórmulas del Resumen
ENCABEZADO = {"font": Font(bold=True, color="FFFFFF"), "fill": PatternFill("solid", fgColor="1F4E78")}


def _texto(v):
    return str(v or "").strip()


def normalizar_doc(d):
    nd = d["Nro_Doc"]
    # Excel guarda el DNI como número y se come los ceros: 01234567 -> 1234567
    nd = "" if nd is None else str(int(nd)) if isinstance(nd, float) else str(nd).strip()
    if set(nd) == {"0"}:   # Odoo pone 00000000 a "Consumidor Final": es sin documento
        nd = ""
    td = "" if d["Tipo_Doc"] is None else str(d["Tipo_Doc"]).strip().removesuffix(".0")
    if td == "":
        td = "0" if nd == "" else "6" if len(nd) == 11 else "1"
    if td == "1" and nd.isdigit() and len(nd) < 8:
        nd = nd.zfill(8)
    d["Tipo_Doc"], d["Nro_Doc"] = td, nd


def _estilo_encabezado(celdas):
    for c in celdas:
        c.font, c.fill = copy(ENCABEZADO["font"]), copy(ENCABEZADO["fill"])


def asegurar_columna(ws, enc, nombre):
    """Agrega al final una columna que falte en un Excel creado con una versión anterior."""
    if nombre not in enc:
        col = len(enc) + 1
        ws.cell(1, col, nombre)
        ws.cell(1, col).font = copy(ws.cell(1, col - 1).font)
        ws.cell(1, col).fill = copy(ws.cell(1, col - 1).fill)
        enc.append(nombre)


def preparar_libro(wb, cuentas):
    """Deja el libro al día con config.yaml cada vez que se abre:
    columnas técnicas, lista desplegable de empresas, filtros en Ventas y hoja Resumen.
    Devuelve idx {columna: índice 0-based}."""
    ws = wb["Ventas"]
    enc = [c.value for c in ws[1]]
    asegurar_columna(ws, enc, "Ref_Odoo")
    asegurar_columna(ws, enc, "Huella")   # columna técnica para detectar filas corregidas
    idx = {n: i for i, n in enumerate(enc)}
    col_cuenta = get_column_letter(idx["Cuenta"] + 1)

    # Lista desplegable de Cuenta con TODAS las empresas de config.yaml (reemplaza la anterior)
    ws.data_validations.dataValidation = [
        dv for dv in ws.data_validations.dataValidation if not str(dv.sqref).startswith(col_cuenta)]
    dv = DataValidation(type="list", formula1='"' + ",".join(cuentas) + '"', allow_blank=True,
                        showErrorMessage=True, errorTitle="Cuenta no válida",
                        error="Elige una cuenta de la lista: " + ", ".join(cuentas))
    ws.add_data_validation(dv)
    dv.add(f"{col_cuenta}2:{col_cuenta}{FILAS_MAX}")

    # Filtros (flechitas) en el encabezado: filtrar por Cuenta/empresa, Estado, Fecha...
    ws.auto_filter.ref = f"A1:{get_column_letter(len(enc))}{max(ws.max_row, 2)}"
    ws.freeze_panes = "A2"

    escribir_resumen(wb["Resumen"] if "Resumen" in wb.sheetnames else wb.create_sheet("Resumen"), cuentas, idx)
    return idx


def crear_plantilla(ruta, cuentas, filas_demo=()):
    wb = Workbook()
    ws = wb.active
    ws.title = "Ventas"
    ws.append(COLUMNAS)
    _estilo_encabezado(ws[1])
    for i, f in enumerate(filas_demo, start=2):
        ref = f[9] if len(f) > 9 else None   # 10º valor opcional: Ref_Odoo (agrupa productos)
        ws.append(list(f[:9]) + [f"=ROUND(H{i}*I{i},2)", "PENDIENTE", None, None, None, ref])
    for i in range(1, len(COLUMNAS) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 16
    ws.column_dimensions["G"].width = 32
    preparar_libro(wb, cuentas)
    wb.save(ruta)


def escribir_resumen(rs, cuentas, idx=None):
    """Hoja Resumen por cuenta. Una boleta puede ocupar varias filas (una por producto):
    se cuentan números de boleta distintos, no filas."""
    idx = idx or {n: i for i, n in enumerate(COLUMNAS)}
    L = {n: get_column_letter(idx[n] + 1) for n in ("Cuenta", "Total", "Estado", "Serie_Numero")}
    rango = lambda n: f"Ventas!${L[n]}$2:${L[n]}${FILAS_MAX}"
    rs.delete_rows(1, rs.max_row)
    rs.append(["Cuenta", "N° boletas emitidas", "Total boleteado (S/)", "Productos pendientes",
               "Productos con error", "Total pendiente (S/)"])
    _estilo_encabezado(rs[1])
    for n, cuenta in enumerate(cuentas, start=2):
        si = lambda estado: f'{rango("Cuenta")},A{n},{rango("Estado")},"{estado}"'
        rs.append([
            cuenta,
            # Boletas distintas: cada número emitido suma 1/(veces que aparece)
            f'=SUMPRODUCT(({rango("Cuenta")}=A{n})*({rango("Estado")}="EMITIDA")'
            f'/COUNTIFS({rango("Serie_Numero")},{rango("Serie_Numero")}&"",{rango("Cuenta")},{rango("Cuenta")}&""))',
            f'=SUMIFS({rango("Total")},{si("EMITIDA")})',
            f'=COUNTIFS({si("PENDIENTE")})',
            f'=COUNTIFS({si("ERROR")})',
            f'=SUMIFS({rango("Total")},{si("PENDIENTE")})',
        ])
    t = len(cuentas) + 2
    rs.append(["TOTAL"] + [f"=SUM({get_column_letter(c)}2:{get_column_letter(c)}{t-1})" for c in range(2, 7)])
    for c in rs[t]:
        c.font = Font(bold=True)
    for fila in rs.iter_rows(min_row=2):
        fila[2].number_format = fila[5].number_format = '#,##0.00'
    for i in range(1, 7):
        rs.column_dimensions[get_column_letter(i)].width = 24


def huella(d):
    """Resumen de los datos de la venta: si cambia, la fila en ERROR se vuelve a intentar."""
    datos = "|".join(str(d[c]) for c in CAMPOS_VENTA)
    return hashlib.md5(datos.encode("utf-8")).hexdigest()[:12]


def _total(d):
    # El Total es fórmula; sin valor cacheado lo calculamos
    try:
        return round(float(d["Cantidad"]) * float(d["Precio_Unit"]), 2)
    except (TypeError, ValueError):
        return 0.0   # validar() lo rechaza como "total debe ser mayor a 0"


def leer_pendientes(ruta, cuentas):
    """Devuelve (wb, ws, idx, pendientes, filas_por_ref).
    Pendientes = filas PENDIENTE/vacías, y filas ERROR cuyos datos cambiaron desde el error
    (ERROR = nunca se llegó a emitir, reintentar es seguro). REVISAR nunca se reintenta solo.
    Se recorre la hoja una sola vez."""
    wb = load_workbook(ruta)
    ws = wb["Ventas"]
    idx = preparar_libro(wb, cuentas)
    nombres = list(idx)
    filas, filas_por_ref = [], {}
    for r, valores in enumerate(ws.iter_rows(min_row=2, max_col=len(nombres), values_only=True), start=2):
        if valores[0] is None:   # sin ID = fila vacía
            continue
        d = dict(zip(nombres, valores))
        d["fila"] = r
        d["Ref_Odoo"] = _texto(d.get("Ref_Odoo"))
        filas.append(d)
        if d["Ref_Odoo"]:
            filas_por_ref.setdefault(d["Ref_Odoo"], []).append(r)

    pend = [d for d in filas if d["Estado"] in (None, "", "PENDIENTE")
            or (d["Estado"] == "ERROR" and d["Huella"] != huella(d))]
    # Si se corrigió una fila de una orden en ERROR, se reintenta la orden completa
    refs = {d["Ref_Odoo"] for d in pend} - {""}
    ya = {d["fila"] for d in pend}
    pend += [d for d in filas if d["fila"] not in ya and d["Estado"] == "ERROR" and d["Ref_Odoo"] in refs]
    pend.sort(key=lambda d: d["fila"])
    for d in pend:
        d["Total"] = _total(d)
    return wb, ws, idx, pend, filas_por_ref


def agrupar(pend, filas_por_ref):
    """Junta las filas pendientes en boletas: misma Ref_Odoo = una boleta con varios productos.
    Cuenta/Cliente/DNI se toman de la fila de la orden que los tenga (basta con la primera)."""
    grupos = {}
    for d in pend:
        grupos.setdefault(d["Ref_Odoo"] or f"fila {d['fila']}", []).append(d)
    boletas = []
    for clave, filas in grupos.items():
        b = {"clave": clave, "fila": filas[0]["fila"], "filas": [d["fila"] for d in filas], "error": None}
        for campo in ("Cuenta", "Tipo_Doc", "Nro_Doc", "Cliente"):
            # texto -> valor original (el DNI puede venir como número de Excel)
            valores = {_texto(d[campo]): d[campo] for d in filas if _texto(d[campo])}
            b[campo] = next(iter(valores.values())) if len(valores) == 1 else None
            if len(valores) > 1:
                b["error"] = f"Las filas de {clave} tienen distinto {campo}: {sorted(valores)}"
        b["Cuenta"] = _texto(b["Cuenta"])
        b["items"] = [{k: d[k] for k in ("Descripcion", "Cantidad", "Precio_Unit", "Total")} for d in filas]
        b["Total"] = round(sum(i["Total"] for i in b["items"]), 2)
        # Toda la orden debe estar pendiente: si una parte ya se emitió, no emitir el resto solo
        en_hoja = filas_por_ref.get(clave)
        if en_hoja and en_hoja != b["filas"]:
            b["error"] = f"La orden {clave} tiene filas en otro estado (filas {en_hoja}); revísala"
        normalizar_doc(b)
        boletas.append(b)
    return boletas


def marcar(ws, idx, fila, estado, serie_numero=None, error=None):
    """Escribe el resultado en la fila (sin guardar: usar guardar() al terminar la boleta)."""
    celda = lambda col: ws.cell(fila, idx[col] + 1)
    celda("Estado").value = estado
    celda("Serie_Numero").value = serie_numero
    celda("Fecha_Emision").value = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if estado == "EMITIDA" else None
    celda("Error").value = error
    if estado == "ERROR":
        d = {c: ws.cell(fila, idx[c] + 1).value for c in CAMPOS_VENTA}
        celda("Huella").value = huella(d)
    else:
        celda("Huella").value = None


def guardar(wb, ruta, espera_max_seg=300):
    """Guarda el Excel. Si está abierto en Excel (PermissionError) avisa y reintenta:
    una boleta ya emitida en SUNAT no debe quedar sin anotar."""
    inicio = time.monotonic()
    while True:
        try:
            wb.save(ruta)
            return
        except PermissionError:
            if time.monotonic() - inicio > espera_max_seg:
                raise
            print(f"       ! No puedo guardar {ruta}: ciérralo en Excel (reintento en 5 s)")
            time.sleep(5)
