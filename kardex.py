"""Extracción masiva de Kárdex de Depósitos e Historiales Académicos desde academico.upedagogica.edu.bo."""
import io
import os
import platform
import re
import subprocess
import sys
import unicodedata
import pandas as pd
from playwright.sync_api import sync_playwright
from pypdf import PdfWriter

if platform.system() == "Linux":
    try:
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=False)
    except Exception:
        pass

ARCHIVO_SESION = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sesion_academico.json")

RUTAS_REPORTE = {
    "deposito": "/reportes/deposito/",
    "historial": "/reportes/historial/",
}


def normalizar_texto(texto):
    texto = str(texto).upper().strip()
    texto = "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")
    texto = texto.replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", texto)


def extraer_nombre_limpio_title(attr_title):
    t = normalizar_texto(attr_title)
    if "|" in t:
        t = t.split("|", 1)[1].strip()
    t = re.sub(r"\s*-\s*(VIRTUAL|PRESENCIAL|SEMIPRESENCIAL|EN LINEA).*$", "", t).strip()
    return t


def leer_listado_kardex(archivo):
    if hasattr(archivo, "seek"):
        archivo.seek(0)
    raw = pd.read_excel(archivo, header=None, dtype=object, engine="openpyxl")

    fila_header = 0
    for i in range(min(15, len(raw))):
        vals = {normalizar_texto(c) for c in raw.iloc[i] if pd.notna(c)}
        if "CARNET" in vals and "PROGRAMA" in vals:
            fila_header = i
            break

    df = raw.iloc[fila_header + 1:].copy()
    df.columns = [normalizar_texto(c) if pd.notna(c) else f"COL_{j}" for j, c in enumerate(raw.iloc[fila_header])]

    if "CARNET" not in df.columns or "PROGRAMA" not in df.columns:
        return pd.DataFrame(), "El archivo Excel debe tener las columnas **CARNET** y **PROGRAMA**."

    df = df.dropna(subset=["CARNET", "PROGRAMA"]).copy()
    df["CARNET"] = df["CARNET"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["PROGRAMA"] = df["PROGRAMA"].astype(str).str.strip()
    df = df[(df["CARNET"] != "") & (df["PROGRAMA"] != "")]
    return df, None


def hay_sesion_guardada():
    return os.path.exists(ARCHIVO_SESION)


def procesar_kardex_lote(df, usuario="", password="", tipo_documento="deposito", on_progress=None):
    """
    Busca por CARNET + nombre de PROGRAMA + PAF APROBADO.
    tipo_documento: 'deposito' (Kárdex de pagos) o 'historial' (Historial Académico).
    """
    ruta_objetivo = RUTAS_REPORTE.get(tipo_documento, "/reportes/deposito/")
    merger = PdfWriter()
    resultados_log = []
    total = len(df)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        if os.path.exists(ARCHIVO_SESION):
            context = browser.new_context(storage_state=ARCHIVO_SESION)
        else:
            context = browser.new_context()

        page = context.new_page()

        try:
            page.goto("https://academico.upedagogica.edu.bo/participantes", wait_until="networkidle")

            necesita_login = "login" in page.url.lower() or page.locator("input[type='password']").count() > 0
            if necesita_login:
                if not usuario or not password:
                    if os.path.exists(ARCHIVO_SESION):
                        os.remove(ARCHIVO_SESION)
                    return None, pd.DataFrame(), "NECESITA_LOGIN"

                if on_progress:
                    on_progress(0, total, "🔐 Iniciando sesión por única vez...")

                campo_user = page.locator("input[type='email'], input[name*='user'], input[name*='email'], input[type='text']").first
                campo_pass = page.locator("input[type='password']").first
                campo_user.fill(usuario)
                campo_pass.fill(password)
                page.locator("button[type='submit'], input[type='submit']").first.click()
                page.wait_for_load_state("networkidle")

                if "login" in page.url.lower() or page.locator("input[type='password']").count() > 0:
                    return None, pd.DataFrame(), "❌ Usuario o contraseña incorrectos."

                context.storage_state(path=ARCHIVO_SESION)

            for idx, (_, row) in enumerate(df.iterrows()):
                carnet = row["CARNET"]
                prog_excel = row["PROGRAMA"]
                prog_excel_norm = normalizar_texto(prog_excel)

                if on_progress:
                    on_progress(idx + 1, total, f"Procesando ({idx + 1}/{total}): Carnet **{carnet}**")

                url_busqueda = f"https://academico.upedagogica.edu.bo/participantes?carnet={carnet}&nombre=&paterno=&materno="
                page.goto(url_busqueda, wait_until="networkidle")

                filas = page.locator("table.table-striped tbody tr")
                cantidad_filas = filas.count()
                href_elegido = None
                motivo_fallo = "Carnet sin registros" if cantidad_filas == 0 else "No coincidió programa + PAF APROBADO"
                programas_vistos = []

                for i in range(cantidad_filas):
                    fila = filas.nth(i)
                    texto_fila_norm = normalizar_texto(fila.inner_text())

                    celda_prog = fila.locator("td[title]").first
                    title_raw = celda_prog.get_attribute("title") if celda_prog.count() > 0 else ""
                    title_completo_norm = normalizar_texto(title_raw)
                    title_limpio_norm = extraer_nombre_limpio_title(title_raw)
                    sigla_visible_norm = normalizar_texto(celda_prog.inner_text()) if celda_prog.count() > 0 else ""

                    tiene_paf = "PAF APROBADO" in texto_fila_norm
                    programas_vistos.append(f"{sigla_visible_norm} ({'PAF OK' if tiene_paf else 'Sin PAF'})")

                    coincide_programa = (
                        (prog_excel_norm in title_completo_norm)
                        or (title_limpio_norm != "" and title_limpio_norm in prog_excel_norm)
                        or (prog_excel_norm in sigla_visible_norm)
                    )

                    if coincide_programa and tiene_paf:
                        enlace_doc = fila.locator(f"a.dropdown-item[href*='{ruta_objetivo}']").first
                        if enlace_doc.count() > 0:
                            href_elegido = enlace_doc.get_attribute("href")
                            break
                    elif coincide_programa and not tiene_paf:
                        motivo_fallo = "Programa coincide, pero NO tiene PAF APROBADO"

                if href_elegido:
                    if href_elegido.startswith("/"):
                        href_elegido = f"https://academico.upedagogica.edu.bo{href_elegido}"

                    respuesta_pdf = context.request.get(href_elegido)
                    if respuesta_pdf.status == 200 and "pdf" in respuesta_pdf.headers.get("content-type", "").lower():
                        merger.append(io.BytesIO(respuesta_pdf.body()))
                        resultados_log.append({
                            "Carnet": carnet,
                            "Programa Excel": prog_excel,
                            "Hallado en Web": " | ".join(programas_vistos),
                            "Estado": "✅ Descargado",
                        })
                    else:
                        resultados_log.append({
                            "Carnet": carnet,
                            "Programa Excel": prog_excel,
                            "Hallado en Web": " | ".join(programas_vistos),
                            "Estado": f"⚠️ Error HTTP {respuesta_pdf.status}",
                        })
                else:
                    resultados_log.append({
                        "Carnet": carnet,
                        "Programa Excel": prog_excel,
                        "Hallado en Web": " | ".join(programas_vistos) if programas_vistos else "Sin registros",
                        "Estado": f"❌ {motivo_fallo}",
                    })

        finally:
            context.close()
            browser.close()

    pdf_bytes = None
    if len(merger.pages) > 0:
        out = io.BytesIO()
        merger.write(out)
        merger.close()
        out.seek(0)
        pdf_bytes = out.getvalue()

    return pdf_bytes, pd.DataFrame(resultados_log), None