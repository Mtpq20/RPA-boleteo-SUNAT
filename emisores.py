"""Emisores: SimuladorEmisor (pruebas) y SunatSolEmisor (portal real, Playwright).

SUNAT cambia su web de vez en cuando; todo lo frágil está concentrado en SELECTORES
y en los métodos _login, _ir_a_formulario y _llenar.
"""
import os
import re
import time
import unicodedata
import zipfile
from itertools import count
from pathlib import Path


class ErrorValidacion(Exception):
    pass


def validar(b):
    """b es una boleta (excel_io.agrupar): datos del cliente + lista de items (productos)."""
    if b.get("error"):
        raise ErrorValidacion(b["error"])
    td = str(b["Tipo_Doc"])
    nd = str(b["Nro_Doc"] or "").strip()
    if td == "1" and not re.fullmatch(r"\d{8}", nd):
        raise ErrorValidacion(f"DNI inválido: {nd!r}")
    if td == "6" and not re.fullmatch(r"\d{11}", nd):
        raise ErrorValidacion(f"RUC inválido: {nd!r}")
    if td == "0" and b["Total"] >= 700:
        raise ErrorValidacion("Boleta >= S/700 requiere DNI del cliente")
    for i in b["items"]:
        if not str(i["Descripcion"] or "").strip():
            raise ErrorValidacion("Falta la Descripcion de un producto")
        if i["Total"] <= 0:
            raise ErrorValidacion(f"{i['Descripcion']}: el total del producto debe ser mayor a 0")


class SimuladorEmisor:
    def __init__(self, cuentas):
        self.cuentas = cuentas
        self.n = {c: count(1) for c in cuentas}

    def emitir(self, b):
        validar(b)
        cfg = self.cuentas[b["Cuenta"]]
        return f"{cfg['serie']}-{next(self.n[b['Cuenta']]):08d}"

    def cerrar(self):
        pass


# ---- Calibrado con playwright codegen (26/10/2026). Si SUNAT cambia su web, regrabar y ajustar aquí ----
SELECTORES = {
    "login_url": "https://e-menu.sunat.gob.pe/cl-ti-itmenu/MenuInternet.htm",   # redirige al login
    # Menú lateral: Empresas > Comprobantes de pago > SEE - SOL > Boleta de Venta Electrónica > Emitir
    "menu_pasos": ["Comprobantes de pago", "SEE - SOL", "Boleta de Venta Electrónica"],
    "menu_emitir": "Emitir Boleta de Venta",
    "iframe_aviso": 'iframe[name="ifrVCE"]',          # aviso de confirmar correo/celular tras el login
    "iframe": 'iframe[name="iframeApplication"]',     # el formulario vive en este iframe
    "tipo_doc": '[id="inicio.tipoDocumento"]',
    "nro_doc": '[id="inicio.numeroDocumento"]',
    "cliente": '[id="inicio.razonSocial"]',           # con DNI/RUC lo llena SUNAT solo
    "descripcion": '[id="item.descripcion"]',
    "cantidad": '[id="item.cantidad"]',               # no salió en la grabación: se usa solo si existe
    "precio": '[id="item.precioUnitario"]',
}
# Texto de cada opción de "Tipo de documento" en el portal, según Tipo_Doc del Excel
OPCIONES_DOC = {"0": r"SIN DOCUMENTO", "1": r"NACIONAL DE IDEN", "6": r"CONTRIBUYENTE|R\.?U\.?C"}

BASE = Path(__file__).parent


def precio_texto(precio):
    """30.0 -> "30", 59.9 -> "59.9", 1.25 -> "1.25" (como lo teclearía una persona)."""
    return f"{precio:.2f}".rstrip("0").rstrip(".")


def numero_en_texto(texto):
    """Busca una serie-número de boleta (EB01-12, B001-345) en el texto de la pantalla final."""
    m = re.search(r"\b(E?B[A-Z0-9]{2,3})\s*-\s*(\d{1,8})\b", texto or "")
    return f"{m.group(1)}-{int(m.group(2))}" if m else None


def sin_tildes(texto):
    """'Inversión  en el Perú' -> 'INVERSION EN EL PERU' (para comparar con lo que muestra SUNAT)."""
    t = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return " ".join(t.upper().split())


def a_numero(texto):
    try:
        return float(str(texto).replace(",", "").replace("S/", "").strip())
    except ValueError:
        return None


def leer_numero(archivo):
    """Saca la serie-número (p. ej. EB01-1) del XML/ZIP que descarga SUNAT, o de su nombre."""
    datos = archivo.read_bytes()
    if zipfile.is_zipfile(archivo):
        with zipfile.ZipFile(archivo) as z:
            datos = b"".join(z.read(n) for n in z.namelist() if n.lower().endswith(".xml"))
    m = re.search(r"<cbc:ID>([A-Z]\w{3}-\d+)</cbc:ID>", datos.decode("utf-8", "ignore"))
    if m:
        return m.group(1)
    m = re.search(r"-0?3-([A-Z]\w{3})-(\d+)", archivo.name)   # RUC-03-EB01-1.xml
    return f"{m.group(1)}-{m.group(2)}" if m else None


class SunatSolEmisor:
    """ensayo=True llena la boleta hasta la vista previa y NO presiona Emitir."""

    def __init__(self, cuentas, headless=False, ensayo=False):
        from playwright.sync_api import sync_playwright
        self.cuentas = cuentas
        self.ensayo = ensayo
        self.pw = sync_playwright().start()
        self.browser = self.pw.chromium.launch(headless=headless)
        self.page = None
        self.cuenta_activa = None
        self.login_fallido = set()   # nunca reintentar un login fallido: SUNAT bloquea la clave SOL
        self.descargas = []
        self.ultimo_pdf = None   # PDF de la última boleta emitida (main.py lo copia a Drive)

    @property
    def form(self):
        return self.page.frame_locator(SELECTORES["iframe"])

    def _login(self, nombre):
        from playwright.sync_api import TimeoutError as PWTimeout
        cfg = self.cuentas[nombre]
        clave = os.environ.get(cfg["clave_env"])
        if not clave:
            raise ErrorValidacion(f"Falta {cfg['clave_env']} en el .env")
        if self.page:
            self.page.context.close()
        ctx = self.browser.new_context(accept_downloads=True)
        # Juntar toda descarga (también las de pestañas emergentes) para guardarlas luego
        # (Playwright no acepta métodos nativos como list.append como manejador: usar lambda)
        ctx.on("page", lambda pg: pg.on("download", lambda d: self.descargas.append(d)))
        self.page = ctx.new_page()
        p = self.page
        p.goto(SELECTORES["login_url"])
        p.get_by_role("textbox", name="RUC").fill(cfg["ruc"])
        p.get_by_role("textbox", name="Usuario").fill(cfg["usuario_sol"])
        p.get_by_role("textbox", name="Contraseña").fill(clave)
        p.get_by_role("button", name="Iniciar sesión").click()
        aviso = p.frame_locator(SELECTORES["iframe_aviso"])
        for boton in ("Finalizar", "Continuar sin confirmar"):
            try:
                aviso.get_by_role("button", name=boton).click(timeout=8000)
            except PWTimeout:
                pass   # el aviso no siempre aparece
            except Exception:
                break   # el aviso cerró o cambió la pestaña
        if p.is_closed():
            if not p.context.pages:
                raise ErrorValidacion(f"El navegador se cerró durante el login de {nombre}")
            p = self.page = p.context.pages[-1]
        try:
            p.get_by_role("heading", name="Empresas").wait_for(timeout=20000)
        except PWTimeout:
            self.login_fallido.add(nombre)
            raise ErrorValidacion(f"No se pudo entrar a SOL con {nombre}: revisa RUC/usuario/clave o captcha")
        self.cuenta_activa = nombre

    def _ir_a_formulario(self):
        p, s = self.page, SELECTORES
        emitir = p.get_by_text(s["menu_emitir"]).first
        if not emitir.is_visible():
            p.get_by_role("heading", name="Empresas").click()
            for texto in s["menu_pasos"]:
                p.get_by_text(texto).first.click()
                p.wait_for_timeout(700)
        emitir.click()   # también reinicia el formulario si quedó una boleta a medias
        self.form.locator(s["tipo_doc"]).wait_for(timeout=30000)

    def _elegir_tipo_doc(self, td):
        """La lista de SUNAT tarda en activarse tras cargar el formulario: reintentar el clic."""
        from playwright.sync_api import TimeoutError as PWTimeout
        p, f = self.page, self.form
        campo = f.locator(SELECTORES["tipo_doc"])
        patron = re.compile(OPCIONES_DOC[td], re.I)
        opcion = f.get_by_role("option", name=patron)
        for intento in range(4):
            p.wait_for_timeout(1000)
            campo.click()
            if intento % 2:
                campo.press("ArrowDown")   # otra forma de abrir la lista
            try:
                opcion.first.click(timeout=3000)
                return
            except PWTimeout:
                pass
        # Diagnóstico para recalibrar: captura con la lista abierta y el HTML del formulario
        carpeta = BASE / "capturas"
        carpeta.mkdir(exist_ok=True)
        p.screenshot(path=str(carpeta / "tipo_doc_lista.png"))
        frame = p.frame(name="iframeApplication")
        if frame:
            (carpeta / "tipo_doc_formulario.html").write_text(frame.content(), encoding="utf-8")
        vistas = f.get_by_role("option").all_inner_texts()
        p.keyboard.press("Escape")
        raise ErrorValidacion(f"No se pudo elegir el tipo de documento {td} en SUNAT; opciones vistas: {vistas}")

    def _llenar(self, b):
        p, f, s = self.page, self.form, SELECTORES
        td, nd = b["Tipo_Doc"], b["Nro_Doc"]
        self._elegir_tipo_doc(td)
        if td == "0":
            f.locator(s["cliente"]).fill(str(b["Cliente"] or "CLIENTES VARIOS"))
        else:
            f.locator(s["nro_doc"]).fill(nd)
            f.locator(s["nro_doc"]).press("Tab")   # al salir del campo SUNAT busca el nombre
            for _ in range(20):
                if f.locator(s["cliente"]).input_value().strip():
                    break
                p.wait_for_timeout(500)
            else:
                raise ErrorValidacion(f"SUNAT no encontró el documento {nd}")
        f.get_by_role("button", name="Continuar").click()
        for item in b["items"]:   # un "Adicionar" por producto de la orden
            p.wait_for_timeout(800)
            f.get_by_role("button", name="Adicionar").click()
            f.get_by_role("radio").first.check()   # Bien (no Servicio)
            f.locator(s["descripcion"]).fill(str(item["Descripcion"]))
            if f.locator(s["cantidad"]).count():
                self._escribir_numero(s["cantidad"], float(item["Cantidad"]), "cantidad")
            elif float(item["Cantidad"]) != 1:
                raise ErrorValidacion("No encontré el campo cantidad en SUNAT: usa Cantidad=1 o recalibra")
            self._escribir_numero(s["precio"], float(item["Precio_Unit"]), "precio")
            f.get_by_role("button", name="Aceptar").click()
        for _ in range(2):   # datos adicionales -> vista previa
            p.wait_for_timeout(1000)
            f.get_by_role("button", name="Continuar").click()
        f.get_by_role("button", name="Emitir").wait_for(timeout=30000)
        self._verificar_importe(b)

    def _escribir_numero(self, selector, valor, nombre):
        """Con fill("1.0") el portal guardó 0.001: se teclea como una persona y se relee el campo."""
        p, campo = self.page, self.form.locator(selector)
        texto = precio_texto(valor)
        campo.click()
        campo.press("ControlOrMeta+A")
        campo.press("Delete")
        campo.press_sequentially(texto, delay=80)
        campo.press("Tab")
        p.wait_for_timeout(500)
        visto = a_numero(campo.input_value())
        if visto is None or abs(visto - valor) > 0.005:
            raise ErrorValidacion(f"El campo {nombre} quedó como {campo.input_value()!r} en SUNAT (se escribió {texto})")

    def _verificar_importe(self, b):
        """Última barrera antes de Emitir: cada producto de la vista previa debe tener el importe del Excel."""
        self.form.locator("tr").first.wait_for(timeout=10000)
        textos = self.form.locator("tr").all_inner_texts()
        for item in b["items"]:
            desc = sin_tildes(item["Descripcion"])
            filas = [t for t in textos if desc in sin_tildes(t)]
            numeros = re.findall(r"\d[\d,]*\.\d+", filas[-1]) if filas else []
            importe = a_numero(numeros[-1]) if numeros else None
            if importe is None or abs(importe - item["Total"]) > 0.01:
                raise ErrorValidacion(f"Vista previa: {item['Descripcion']} con importe {importe} y el Excel dice "
                                      f"{item['Total']:.2f}: no se emitió")

    def _captura(self, b, motivo):
        carpeta = BASE / "capturas"
        carpeta.mkdir(exist_ok=True)
        ruta = carpeta / f"fila{b['fila']}_{motivo}_{time.strftime('%Y%m%d_%H%M%S')}.png"
        try:
            self.page.screenshot(path=str(ruta), full_page=True)
        except Exception:
            pass
        return ruta

    def _guardar(self, descarga, b):
        carpeta = BASE / "comprobantes" / b["Cuenta"]
        carpeta.mkdir(parents=True, exist_ok=True)
        ruta = carpeta / f"fila{b['fila']}_{descarga.suggested_filename}"
        descarga.save_as(ruta)
        return ruta

    def emitir(self, b):
        from playwright.sync_api import TimeoutError as PWTimeout
        validar(b)
        if b["Cuenta"] in self.login_fallido:
            raise ErrorValidacion(f"Login de {b['Cuenta']} falló antes en esta corrida; no se reintenta")
        try:
            if self.cuenta_activa != b["Cuenta"]:
                self._login(b["Cuenta"])
            self._ir_a_formulario()
            self._llenar(b)
        except ErrorValidacion:
            self._captura(b, "error")
            raise
        except Exception as e:
            # Aún no se presionó Emitir: seguro que no existe la boleta en SUNAT
            raise ErrorValidacion(f"Portal (antes de emitir): {type(e).__name__}: {e} | ver {self._captura(b, 'error').name}")
        if self.ensayo:
            return f"ENSAYO OK, ver {self._captura(b, 'ensayo').name}"

        # Desde aquí la boleta puede existir en SUNAT: cualquier error queda en REVISAR (main.py)
        try:
            return self._emitir_y_leer(b)
        except Exception:
            self._captura(b, "revisar")
            raise

    def _emitir_y_leer(self, b):
        from playwright.sync_api import TimeoutError as PWTimeout
        p, f = self.page, self.form
        self.descargas.clear()
        f.get_by_role("button", name="Emitir").click()
        f.get_by_role("button", name="Aceptar").click()   # confirmación de SUNAT
        # Pantalla final: la que tiene Descargar PDF / Descargar XML
        f.get_by_role("button", name="Descargar XML").wait_for(timeout=120000)
        self._captura(b, "emitida")
        pantalla = p.frame(name="iframeApplication").inner_text("body")
        try:
            with p.expect_download(timeout=30000):
                f.get_by_role("button", name="Descargar XML").click()
        except PWTimeout:
            pass
        p.wait_for_timeout(1000)
        archivos = [self._guardar(d, b) for d in self.descargas]
        numero = next((n for a in archivos if (n := leer_numero(a))), None) or numero_en_texto(pantalla)
        # El PDF es solo para archivar (Drive): si falla, la boleta igual quedó emitida
        try:
            self.ultimo_pdf = self._descargar_pdf(b, numero)
        except Exception as e:
            print(f"       ! No se pudo obtener el PDF de SUNAT ({type(e).__name__}: {e})")
            self.ultimo_pdf = None
        if not self.ultimo_pdf:
            self.ultimo_pdf = self._pdf_desde_captura(b, numero)
        try:
            f.get_by_role("button", name="Cerrar").click(timeout=5000)
            f.get_by_role("button", name="Aceptar").click(timeout=5000)
        except PWTimeout:
            pass
        if not numero:
            raise RuntimeError("Se emitió pero no pude leer el número; revisa en SUNAT y en capturas/")
        return numero

    def _descargar_pdf(self, b, numero):
        """Botón "Descargar PDF" de la pantalla final. SUNAT puede mandarlo como descarga
        o abrirlo en otra pestaña: se aceptan las dos formas."""
        p, ctx = self.page, self.page.context
        carpeta = BASE / "comprobantes" / b["Cuenta"]
        carpeta.mkdir(parents=True, exist_ok=True)
        ruta = carpeta / f"{numero or 'fila' + str(b['fila'])}.pdf"
        antes, pestanas = len(self.descargas), list(ctx.pages)
        self.form.get_by_role("button", name="Descargar PDF").click(timeout=10000)
        for _ in range(20):
            p.wait_for_timeout(500)
            for d in self.descargas[antes:]:
                if d.suggested_filename.lower().endswith(".pdf"):
                    d.save_as(ruta)
                    return ruta
            for pg in [x for x in ctx.pages if x not in pestanas]:
                if pg.url.startswith("http"):
                    r = ctx.request.get(pg.url)   # usa la misma sesión SOL
                    if r.ok and r.body()[:4] == b"%PDF":
                        ruta.write_bytes(r.body())
                        pg.close()
                        return ruta
        for pg in [x for x in ctx.pages if x not in pestanas]:
            print(f"       ! Pestaña del PDF no reconocida: {pg.url[:120]}")
            pg.close()
        return None

    def _pdf_desde_captura(self, b, numero):
        """Plan B: PDF con la imagen de la pantalla final (no es el formato oficial de SUNAT)."""
        from PIL import Image
        capturas = sorted((BASE / "capturas").glob(f"fila{b['fila']}_emitida_*.png"))
        if not capturas:
            return None
        ruta = BASE / "comprobantes" / b["Cuenta"] / f"{numero or 'fila' + str(b['fila'])}_captura.pdf"
        ruta.parent.mkdir(parents=True, exist_ok=True)
        Image.open(capturas[-1]).convert("RGB").save(ruta, "PDF")
        print(f"       ! Se guardó el PDF desde la captura de pantalla: {ruta.name}")
        return ruta

    def cerrar(self):
        self.browser.close()
        self.pw.stop()


def crear_emisor(modo, cuentas, headless):
    if modo == "simulacion":
        return SimuladorEmisor(cuentas)
    return SunatSolEmisor(cuentas, headless, ensayo=(modo == "ensayo"))
