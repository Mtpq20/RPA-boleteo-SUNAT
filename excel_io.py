"""Lectura/escritura del Excel. Una fila de 'Ventas' = un producto vendido.
Las filas con la misma Ref_Odoo forman UNA boleta (una orden con varios productos);
una fila sin Ref_Odoo es una boleta de un solo producto."""
import hashlib
import os
import re
import time
from copy import copy
from datetime import date, datetime
from pathlib import Path
import yaml
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
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


def preparar_libro(wb, cuentas, limites=None):
    """Deja el libro al día con config.yaml cada vez que se abre:
    columnas técnicas, lista desplegable de empresas, filtros en Ventas y hojas Resumen y Por mes.
    limites = {cuenta: tope mensual S/} (limites_de(cfg)). Devuelve idx {columna: índice 0-based}."""
    ws = wb["Ventas"]
    enc = [c.value for c in ws[1]]
    asegurar_columna(ws, enc, "Ref_Odoo")
    asegurar_columna(ws, enc, "Huella")   # columna técnica para detectar filas corregidas
    idx = {n: i for i, n in enumerate(enc)}
    # Primero lo escrito en las hojas de empresa (si no, se perdería al rehacerlas abajo)
    nuevas = traer_de_hojas_empresa(wb, ws, idx, cuentas)
    if nuevas:
        print(f"Hojas de empresa: {nuevas} fila(s) nuevas o cambiadas pasadas a Ventas")
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
    escribir_por_mes(wb, ws, cuentas, idx, limites or {})
    escribir_hojas_empresa(wb, cuentas, idx)
    return idx


def nombre_empresa(cuenta):
    """Nombre comercial de la cuenta (EMPRESA_A_NOMBRE=Chery en .env); si no hay, la cuenta."""
    return _texto(os.environ.get(f"{cuenta}_NOMBRE")) or cuenta


def usa_odoo(cuenta):
    """False si en config.yaml la cuenta tiene `odoo: false` (su hoja no muestra Ref_Odoo)."""
    cfg = yaml.safe_load((Path(__file__).with_name("config.yaml")).read_text(encoding="utf-8"))
    return ((cfg.get("cuentas") or {}).get(cuenta) or {}).get("odoo", True) is not False


def _nombre_hoja(cuenta):
    # Excel: máx. 31 caracteres y sin \ / ? * [ ] :
    nombre = re.sub(r"[\\/?*\[\]:]", "", nombre_empresa(cuenta))[:31].strip() or cuenta
    return cuenta if nombre in ("Ventas", "Resumen", HOJA_MES) else nombre


def _fecha(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(_texto(v), fmt)
        except ValueError:
            pass
    return None


COLOR_HOJA_EMPRESA = "2E75B6"   # marca las hojas de empresa (se rehacen en cada corrida)
FILAS_HOJA_EMPRESA = 1000       # filas con la fórmula de Total lista para escribir ventas nuevas
# Lo que se puede escribir o corregir en una hoja de empresa y pasa a Ventas
# (Serie_Numero, Fecha_Emision y Error los escribe el bot; Total es fórmula)
EDITABLES = ["Fecha", "Tipo_Doc", "Nro_Doc", "Cliente", "Descripcion", "Cantidad", "Precio_Unit", "Estado", "Ref_Odoo"]


def _igual(a, b):
    """Compara lo de la hoja de empresa con Ventas sin ver cambios donde no los hay
    (fecha como texto o como fecha, 35 y 35.0, 45123456 y '45123456')."""
    if _texto(a) == _texto(b):
        return True
    fa, fb = _fecha(a), _fecha(b)
    if fa and fb:
        return fa == fb
    try:
        return float(_texto(a)) == float(_texto(b))
    except ValueError:
        return False


def traer_de_hojas_empresa(wb, ws, idx, cuentas):
    """Pasa a Ventas lo escrito en las hojas de empresa (Chery, Jose...), antes de rehacerlas:
    - fila sin ID (columna A, oculta) = venta nueva: se agrega a Ventas con la Cuenta de esa hoja
      y queda pendiente (o con el Estado escrito, p. ej. NO EMITIR);
    - fila con ID = venta que ya está en Ventas: se copian los cambios (DNI, precio, Estado...).
      Una EMITIDA o EN_PROCESO no se cambia; de una REVISAR solo se toma el Estado.
    Devuelve cuántas filas se agregaron o cambiaron."""
    col = lambda n: idx[n] + 1
    fila_de_id, ultima, max_id = {}, 1, 0
    for r, (v,) in enumerate(ws.iter_rows(min_row=2, min_col=col("ID"), max_col=col("ID"), values_only=True), start=2):
        if _texto(v):
            fila_de_id[_texto(v)] = r
            ultima = r
            try:
                max_id = max(max_id, int(float(v)))
            except ValueError:
                pass
    L = {n: get_column_letter(col(n)) for n in ("Cantidad", "Precio_Unit")}
    cuenta_de_id = {_texto(v[col("ID") - 1]): _texto(v[col("Cuenta") - 1])
                    for v in ws.iter_rows(min_row=2, max_col=len(idx), values_only=True)}
    cambios = 0
    for hs in list(wb.worksheets):
        enc = [c.value for c in hs[1]]
        color = hs.sheet_properties.tabColor
        if "ID" not in enc or color is None or not str(color.rgb).endswith(COLOR_HOJA_EMPRESA):
            continue   # no es hoja de empresa (o es de una versión anterior, solo copia)
        # Empresa de la hoja: la guardada en la hoja, o por su nombre, o (si se renombró la empresa)
        # por las ventas que ya tiene
        cuenta = hs.sheet_properties.codeName if hs.sheet_properties.codeName in cuentas else next(
            (c for c in cuentas if _nombre_hoja(c) == hs.title), None)
        if cuenta is None:
            ids = [_texto(v) for (v,) in hs.iter_rows(min_row=2, max_col=1, values_only=True) if _texto(v)]
            cuenta = next((cuenta_de_id.get(i) for i in ids if cuenta_de_id.get(i) in cuentas), None)
        if cuenta is None:
            print(f"  ! Hoja {hs.title!r}: no sé de qué empresa es; sus ventas nuevas no se pasaron a Ventas")
            continue
        for valores in hs.iter_rows(min_row=2, max_col=len(enc), values_only=True):
            d = dict(zip(enc, valores))
            # Una fórmula escrita a mano (p. ej. =TODAY() en Fecha) no se copia: Fecha toma el día de hoy
            datos = {n: d[n] for n in EDITABLES if n in d and not _texto(d[n]).startswith("=")}
            if all(_texto(datos.get(n)) == "" for n in ("Nro_Doc", "Cliente", "Descripcion", "Cantidad", "Precio_Unit")):
                continue   # sin datos de venta (vacía, o solo Fecha/Estado/fórmulas)
            if _texto(d.get("ID")) == "":
                max_id += 1
                ultima += 1
                ws.cell(ultima, col("ID")).value = max_id
                ws.cell(ultima, col("Cuenta")).value = cuenta
                for n, v in datos.items():
                    ws.cell(ultima, col(n)).value = v
                if not _fecha(datos.get("Fecha")):
                    ws.cell(ultima, col("Fecha")).value = datetime.now().strftime("%Y-%m-%d")
                if _texto(datos.get("Estado")) == "":
                    ws.cell(ultima, col("Estado")).value = "PENDIENTE"   # que se vea que falta emitir
                ws.cell(ultima, col("Total")).value = f"=ROUND({L['Cantidad']}{ultima}*{L['Precio_Unit']}{ultima},2)"
                cambios += 1
                continue
            r = fila_de_id.get(_texto(d["ID"]))
            estado = ws.cell(r, col("Estado")).value if r else None
            if not r or estado in ("EMITIDA", "EN_PROCESO"):
                continue   # ya emitida (o emitiéndose): lo de Ventas manda
            cambio = False
            for n in (["Estado"] if estado == "REVISAR" else datos):
                if n in datos and not _igual(datos[n], ws.cell(r, col(n)).value):
                    ws.cell(r, col(n)).value = datos[n]
                    cambio = True
            cambios += cambio
    return cambios


def escribir_hojas_empresa(wb, cuentas, idx):
    """Una hoja por empresa con sus filas de Ventas (sin Cuenta), donde también se pueden escribir
    ventas nuevas: traer_de_hojas_empresa() las pasa a Ventas en la siguiente corrida.
    La columna A (ID) va oculta: une cada fila con su fila de Ventas. Total es fórmula.
    Ventas y Resumen quedan ocultas; el Excel se abre en Por mes."""
    for hoja in list(wb.worksheets):   # borrar las generadas antes (también si cambió el nombre)
        color = hoja.sheet_properties.tabColor
        if hoja.title != HOJA_MES and color is not None and str(color.rgb).endswith(COLOR_HOJA_EMPRESA):
            wb.remove(hoja)
    ventas = wb["Ventas"]
    nombres = list(idx)
    # Huella es técnica; Cuenta sobra (la hoja ya es de una empresa)
    columnas = ["ID"] + [n for n in nombres if n not in ("Huella", "ID", "Cuenta")]
    filas = [dict(zip(nombres, v)) for v in ventas.iter_rows(min_row=2, max_col=len(nombres), values_only=True)
             if v[0] is not None]
    for cuenta in cuentas:
        encabezado = [c for c in columnas if c != "Ref_Odoo" or usa_odoo(cuenta)]
        letra = {n: get_column_letter(i) for i, n in enumerate(encabezado, start=1)}
        hs = wb.create_sheet(_nombre_hoja(cuenta))
        hs.sheet_properties.tabColor = COLOR_HOJA_EMPRESA
        hs.sheet_properties.codeName = cuenta   # empresa de la hoja, aunque cambie su nombre
        hs.append(encabezado)
        _estilo_encabezado(hs[1])
        for d in filas:
            if _texto(d["Cuenta"]) == cuenta:
                hs.append([_fecha(d[c]) or d[c] if c == "Fecha" else d[c] for c in encabezado])
        # Total = Cantidad × Precio en todas las filas, también en las vacías para ventas nuevas
        for r in range(2, max(hs.max_row, FILAS_HOJA_EMPRESA) + 1):
            cant, precio = f"{letra['Cantidad']}{r}", f"{letra['Precio_Unit']}{r}"
            hs[f"{letra['Total']}{r}"] = f'=IF(OR({cant}="",{precio}=""),"",ROUND({cant}*{precio},2))'
        for n, l in letra.items():
            hs.column_dimensions[l].width = (ventas.column_dimensions[get_column_letter(idx[n] + 1)].width
                                             if n in idx else None) or 13
            for celda in hs[l][1:]:
                if n == "Fecha":
                    celda.number_format = "DD/MM/YYYY"
                elif n in ("Precio_Unit", "Total"):
                    celda.number_format = '#,##0.00'
        # Estado: lista para elegir en las ventas nuevas (el bot escribe los demás estados)
        estados = DataValidation(type="list", formula1='"PENDIENTE"', allow_blank=True,
                                 showErrorMessage=False)
        hs.add_data_validation(estados)
        estados.add(f"{letra['Estado']}2:{letra['Estado']}{hs.max_row}")
        hs.column_dimensions["A"].hidden = True   # ID: no se ve, pero no borrar esa columna
        hs.auto_filter.ref = f"A1:{get_column_letter(len(encabezado))}{hs.max_row}"
        hs.freeze_panes = "B2"

    # Al abrir el Excel no se ven Ventas ni Resumen (siguen ahí; en Excel: clic derecho > Mostrar)
    for nombre in ("Ventas", "Resumen"):
        if nombre in wb.sheetnames:
            wb[nombre].sheet_state = "hidden"
    primera = next(h for h in wb.worksheets if h.sheet_state == "visible")   # Por mes
    for h in wb.worksheets:
        h.sheet_view.tabSelected = h is primera
    wb.active = wb.worksheets.index(primera)


def crear_plantilla(ruta, cuentas, filas_demo=(), limites=None):
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
    preparar_libro(wb, cuentas, limites)
    wb.save(ruta)


def escribir_resumen(rs, cuentas, idx=None):
    """Hoja Resumen por cuenta. Una boleta puede ocupar varias filas (una por producto):
    se cuentan números de boleta distintos, no filas."""
    idx = idx or {n: i for i, n in enumerate(COLUMNAS)}
    L = {n: get_column_letter(idx[n] + 1) for n in ("Cuenta", "Total", "Estado", "Serie_Numero")}
    rango = lambda n: f"Ventas!${L[n]}$2:${L[n]}${FILAS_MAX}"
    rs.delete_rows(1, rs.max_row)
    rs.append(["Cuenta", "N° boletas emitidas", "Total boleteado (S/)", "Productos pendientes",
               "Productos con error", "Total pendiente (S/)", "Empresa"])
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
            nombre_empresa(cuenta),
        ])
    t = len(cuentas) + 2
    rs.append(["TOTAL"] + [f"=SUM({get_column_letter(c)}2:{get_column_letter(c)}{t-1})" for c in range(2, 7)])
    for c in rs[t]:
        c.font = Font(bold=True)
    for fila in rs.iter_rows(min_row=2):
        fila[2].number_format = fila[5].number_format = '#,##0.00'
    for i in range(1, 8):
        rs.column_dimensions[get_column_letter(i)].width = 24


MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
HOJA_MES = "Por mes"


def _mes(v):
    """'2026-11-10 12:00:00' o fecha de Excel -> '2026-11' (None si no se reconoce)."""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m")
    t = _texto(v)[:7]
    return t if len(t) == 7 and t[4] == "-" and t[:4].isdigit() and t[5:].isdigit() else None


ALERTA_PORC = 0.8   # desde este % del tope mensual se avisa "POR LLEGAR AL TOPE"


def _relleno(color):
    # En formato condicional Excel toma el color de relleno de bgColor (end_color)
    return PatternFill("solid", start_color=color, end_color=color)


def escribir_por_mes(wb, ws, cuentas, idx, limites=None):
    """Primera hoja del libro: total boleteado (boletas EMITIDAS) por mes de emisión y por empresa.
    Lista los meses que ya tienen emisiones más el mes actual (las boletas de hoy caen ahí).
    Con tope mensual (limites), para el mes actual: tope, lo que está en REVISAR, disponible y una
    fila de alerta por empresa (✓ OK / ⚠ POR LLEGAR AL TOPE / ⛔ TOPE ALCANZADO), más un aviso
    general en la fila 1 que se ve apenas se abre el Excel. Todo es fórmula: se actualiza solo."""
    limites = limites or {}
    con_tope = [cu for cu in cuentas if limites.get(cu)]
    rs = wb[HOJA_MES] if HOJA_MES in wb.sheetnames else wb.create_sheet(HOJA_MES)
    L = {n: get_column_letter(idx[n] + 1) for n in ("Cuenta", "Total", "Estado", "Fecha_Emision")}
    rango = lambda n: f"Ventas!${L[n]}$2:${L[n]}${FILAS_MAX}"
    mes_actual = datetime.now().strftime("%Y-%m")
    meses = {mes_actual}
    for (v,) in ws.iter_rows(min_row=2, min_col=idx["Fecha_Emision"] + 1,
                             max_col=idx["Fecha_Emision"] + 1, values_only=True):
        if _mes(v):
            meses.add(_mes(v))
    meses = sorted(meses)
    col = lambda c: get_column_letter(c + 2)          # columna de la empresa número c (B, C, ...)
    fin = col(len(cuentas) - 1)                       # última columna de empresa
    ultima = get_column_letter(len(cuentas) + 2)      # columna TOTAL

    # Filas: 1 aviso, 3 encabezado, 4.. meses, TOTAL, (vacía), tope, en REVISAR, disponible, alerta
    ENC = 3
    primero = ENC + 1
    total = primero + len(meses)
    actual = primero + meses.index(mes_actual)
    tope, revisar, disp, alerta = total + 2, total + 3, total + 4, total + 5

    for rango_unido in list(rs.merged_cells.ranges):
        rs.unmerge_cells(str(rango_unido))
    rs.delete_rows(1, rs.max_row)
    rs.conditional_formatting = ConditionalFormattingList()   # no acumular reglas en cada corrida
    rs.cell(ENC, 1, "Mes")
    for c, cu in enumerate(cuentas):
        rs[f"{col(c)}{ENC}"] = nombre_empresa(cu)   # Chery, Prime... (las fórmulas usan la cuenta)
    rs[f"{ultima}{ENC}"] = "TOTAL"
    _estilo_encabezado(rs[ENC])
    for n, mes in enumerate(meses, start=primero):
        rs.cell(n, 1, f"{MESES[int(mes[5:]) - 1].capitalize()} {mes[:4]}")   # Octubre 2026
        for c in range(len(cuentas)):
            # Fecha_Emision es texto 'AAAA-MM-DD hh:mm:ss': sus 7 primeros caracteres son el mes
            rs[f"{col(c)}{n}"] = (f'=SUMPRODUCT((LEFT({rango("Fecha_Emision")},7)="{mes}")'
                                  f'*({rango("Cuenta")}="{cuentas[c]}")*({rango("Estado")}="EMITIDA"),{rango("Total")})')
        rs[f"{ultima}{n}"] = f"=SUM(B{n}:{fin}{n})"
        rs[f"{ultima}{n}"].font = Font(bold=True)
    rs.cell(total, 1, "TOTAL")
    for c in range(len(cuentas) + 1):
        rs[f"{col(c)}{total}"] = f"=SUM({col(c)}{primero}:{col(c)}{total - 1})"
    for c in rs[total]:
        c.font = Font(bold=True)

    if con_tope:
        etiqueta_mes = rs.cell(actual, 1).value
        rs.cell(tope, 1, "Tope mensual")
        # Como el bot: lo que está en REVISAR/EN_PROCESO pudo emitirse y también resta del tope.
        # Es un valor (no fórmula) calculado en la última corrida del bot.
        rs.cell(revisar, 1, "En REVISAR (pudo emitirse)")
        rs.cell(disp, 1, f"Disponible {etiqueta_mes}")
        rs.cell(alerta, 1, f"Alerta {etiqueta_mes}")
        partes = []
        for c, cu in enumerate(cuentas):
            if not limites.get(cu):
                continue
            x = col(c)
            usado = f"({x}{actual}+{x}{revisar})"
            rs[f"{x}{tope}"] = limites[cu]
            rs[f"{x}{revisar}"] = round(boleteado_en_mes(ws, idx, cu, mes_actual)
                                        - _emitido_en_mes(ws, idx, cu, mes_actual), 2)
            rs[f"{x}{disp}"] = f"=MAX(0,{x}{tope}-{usado})"
            rs[f"{x}{alerta}"] = (f'=IF({usado}>={x}{tope},"⛔ TOPE ALCANZADO",'
                                  f'IF({usado}>={ALERTA_PORC}*{x}{tope},"⚠ POR LLEGAR AL TOPE","✓ OK"))')
            partes.append(f'IF(LEFT({x}{alerta},1)="✓","",{x}${ENC}&": "&{x}{alerta}'
                          f'&" (quedan S/ "&FIXED({x}{disp},2)&")   ")')
            # Celdas del mes en amarillo desde el % de alerta y en rojo al llegar al tope
            rs.conditional_formatting.add(f"{x}{primero}:{x}{total - 1}", CellIsRule(
                operator="greaterThanOrEqual", formula=[f"${x}${tope}"], fill=_relleno("F8696B")))
            rs.conditional_formatting.add(f"{x}{primero}:{x}{total - 1}", CellIsRule(
                operator="greaterThanOrEqual", formula=[f"{ALERTA_PORC}*${x}${tope}"], fill=_relleno("FFEB84")))
        for c in rs[disp] + rs[alerta]:
            c.font = Font(bold=True)

        # Fila de alerta por empresa y aviso general (fila 1) con colores según el texto
        avisos = "&".join(partes)
        rs["A1"] = (f'=IF(({avisos})="","✓ {etiqueta_mes}: todas las empresas están dentro del tope",'
                    f'IF(COUNTIF(B{alerta}:{fin}{alerta},"⛔*")>0,"⛔ TOPE ALCANZADO — ","⚠ CERCA DEL TOPE — ")'
                    f'&"{etiqueta_mes}: "&{avisos})')
        rs.merge_cells(f"A1:{ultima}1")
        rs["A1"].font = Font(bold=True, size=13)
        rs["A1"].alignment = Alignment(wrap_text=True, vertical="center")
        rs.row_dimensions[1].height = 36
        for rango_cf, ref in ((f"A1:{ultima}1", "$A$1"), (f"B{alerta}:{fin}{alerta}", f"B{alerta}")):
            for simbolo, color in (("⛔", "F8696B"), ("⚠", "FFC000"), ("✓", "C6EFCE")):
                rs.conditional_formatting.add(rango_cf, FormulaRule(
                    formula=[f'LEFT({ref},1)="{simbolo}"'], fill=_relleno(color)))
    else:
        rs["A1"] = "Boleteado por mes y por empresa (boletas EMITIDAS)"
        rs["A1"].font = Font(bold=True, size=13)

    for r in rs.iter_rows(min_row=primero, max_row=disp if con_tope else total):
        for c in r[1:]:
            c.number_format = '"S/ "#,##0.00'
    rs.column_dimensions["A"].width = 30
    for i in range(2, len(cuentas) + 3):
        rs.column_dimensions[get_column_letter(i)].width = 22
    rs.freeze_panes = f"B{primero}"

    # Que sea la primera hoja y la que se ve al abrir el Excel
    wb.move_sheet(rs, offset=-wb.index(rs))
    for hoja in wb.worksheets:
        hoja.sheet_view.tabSelected = hoja is rs
    wb.active = 0


def limites_de(cfg):
    """{cuenta: tope mensual en S/} desde config.yaml; sin tope = None."""
    general = cfg.get("limite_mensual")
    return {n: (c or {}).get("limite_mensual", general) or None for n, c in cfg["cuentas"].items()}


def _sumar_mes(ws, idx, cuenta, mes, con_revisar):
    nombres = list(idx)
    total = 0.0
    for valores in ws.iter_rows(min_row=2, max_col=len(nombres), values_only=True):
        d = dict(zip(nombres, valores))
        if _texto(d["Cuenta"]) != cuenta:
            continue
        if (d["Estado"] == "EMITIDA" and _mes(d["Fecha_Emision"]) == mes) or \
           (con_revisar and d["Estado"] in ("REVISAR", "EN_PROCESO") and _mes(d["Fecha"]) == mes):
            total += _total(d)
    return round(total, 2)


def _emitido_en_mes(ws, idx, cuenta, mes):
    return _sumar_mes(ws, idx, cuenta, mes, con_revisar=False)


def boleteado_en_mes(ws, idx, cuenta, mes):
    """Lo boleteado por la cuenta en el mes 'AAAA-MM': EMITIDAS por fecha de emisión, más las que
    quedaron en REVISAR/EN_PROCESO con venta de ese mes (pudieron emitirse: se cuentan por prudencia)."""
    return _sumar_mes(ws, idx, cuenta, mes, con_revisar=True)


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


def leer_pendientes(ruta, cuentas, limites=None):
    """Devuelve (wb, ws, idx, pendientes, filas_por_ref).
    Pendientes = filas PENDIENTE/vacías, y filas ERROR cuyos datos cambiaron desde el error
    (ERROR = nunca se llegó a emitir, reintentar es seguro). REVISAR nunca se reintenta solo.
    Se recorre la hoja una sola vez."""
    wb = load_workbook(ruta)
    ws = wb["Ventas"]
    idx = preparar_libro(wb, cuentas, limites)
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
