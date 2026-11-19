import os
import streamlit as st

from unir import contar, exportar_excel, unir_y_auditar
from kardex import ARCHIVO_SESION, hay_sesion_guardada, leer_listado_kardex, procesar_kardex_lote

st.set_page_config(page_title="Herramientas - Gestión Académica", page_icon="🗂️", layout="wide")

st.sidebar.title("Herramientas")
herramienta = st.sidebar.radio(
    "Elija una herramienta",
    [
        "Unir y revisar Excels",
        "Extraer kardex de pagos",
        "Extraer historiales académicos",
    ],
)

# ==============================================================================
# HERRAMIENTA 1: UNIR Y REVISAR EXCELS
# ==============================================================================
if herramienta == "Unir y revisar Excels":
    st.title("Unir y revisar Excels de valores asignados")
    st.markdown(
        "Suba los Excels de todas las personas que asignaron valores. La herramienta los une en uno solo y revisa que "
        "**certificados de notas, historial académico, hoja de registro universitario y título** sean correlativos, "
        "que no haya valores ni participantes repetidos, y le da un mensaje con el resultado."
    )

    archivos = st.file_uploader(
        "Excels (.xlsx). Puede seleccionar varios a la vez.",
        type=["xlsx"],
        accept_multiple_files=True,
    )
    if not archivos:
        st.stop()

    firma = tuple((a.name, a.size) for a in archivos)
    if st.button("Unir y revisar", type="primary"):
        with st.spinner("Uniendo y revisando..."):
            st.session_state["union"] = (firma, unir_y_auditar([(a.name, a) for a in archivos]))

    guardado = st.session_state.get("union")
    if not guardado or guardado[0] != firma:
        st.info("Cuando termine de subir los archivos, presione **Unir y revisar**.")
        st.stop()

    res = guardado[1]
    if res["df"].empty:
        st.error("No se pudo leer ningún archivo.")
        for e in res["errores_lectura"]:
            st.write("•", e)
        st.stop()

    err, avi = contar(res)
    titular = res["resumen"].split("\n")[0]
    if err == 0 and avi == 0:
        st.success(titular)
    elif err == 0:
        st.warning(titular)
    else:
        st.error(titular)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Archivos", res["n_archivos"])
    c2.metric("Filas", len(res["df"]))
    c3.metric("Participantes", res["partic"].get("unico", 0))
    c4.metric("Errores", err)
    c5.metric("Avisos", avi)

    st.text_area("Mensaje de resumen (puede copiarlo)", res["resumen"], height=330)

    if "_excel" not in res:
        res["_excel"] = exportar_excel(res)
    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇️ Descargar Excel unido y revisado",
        res["_excel"],
        "Consolidado_revisado.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    d2.download_button(
        "⬇️ Descargar mensaje de resumen (.txt)",
        res["resumen"].encode("utf-8"),
        "Resumen_revision.txt",
        "text/plain",
    )

    tab_prob, tab_cons = st.tabs(["Problemas encontrados", "Excel unido"])
    with tab_prob:
        p = res["problemas"].drop(columns=["_idx"])
        if p.empty:
            st.success("No hay problemas.")
        else:
            graves = st.multiselect("Mostrar", ["ERROR", "AVISO"], default=["ERROR", "AVISO"])
            st.dataframe(p[p["Gravedad"].isin(graves)], use_container_width=True, height=420)
    with tab_cons:
        st.caption(
            "La columna OBSERVACIÓN indica qué problema tiene cada fila. En el Excel descargado esas filas "
            "salen pintadas: rojo = error, amarillo = aviso."
        )
        st.dataframe(res["df"], use_container_width=True, height=420)
    with st.expander("Hojas leídas"):
        for m in res["lecturas"]:
            st.write("•", m)

# ==============================================================================
# HERRAMIENTAS 2 Y 3: EXTRACCIÓN DE KÁRDEX DE PAGOS E HISTORIALES ACADÉMICOS
# ==============================================================================
elif herramienta in ("Extraer kardex de pagos", "Extraer historiales académicos"):
    es_historial = herramienta == "Extraer historiales académicos"

    if es_historial:
        st.title("Extraer historiales académicos")
        st.markdown(
            "Sube un Excel con las columnas **CARNET** y **PROGRAMA**. El sistema buscará cada participante, "
            "verificará que coincida el programa con **PAF APROBADO** y unirá todos los **historiales académicos** en un solo PDF."
        )
        tipo_doc = "historial"
        clave_estado = "resultado_historial"
        texto_boton = "🚀 Extraer historiales y generar PDF"
        nombre_archivo_pdf = "Historiales_Academicos_Consolidado.pdf"
        etiqueta_unidad = "historiales"
    else:
        st.title("Extraer kárdex de pagos (Depósitos)")
        st.markdown(
            "Sube un Excel con las columnas **CARNET** y **PROGRAMA**. El sistema buscará cada participante, "
            "verificará que coincida el programa con **PAF APROBADO** y unirá todos los **kárdex de depósitos** en un solo PDF."
        )
        tipo_doc = "deposito"
        clave_estado = "resultado_kardex"
        texto_boton = "🚀 Extraer kárdex y generar PDF"
        nombre_archivo_pdf = "Kardex_Depositos_Consolidado.pdf"
        etiqueta_unidad = "kárdex"

    usuario_acad = st.session_state.get("usuario_acad", "")
    pass_acad = st.session_state.get("pass_acad", "")

    if not hay_sesion_guardada() and (not usuario_acad or not pass_acad):
        with st.form("form_login_academico"):
            st.subheader("🔐 Acceso al sistema académico (se guarda una sola vez)")
            c_u, c_p = st.columns(2)
            u_in = c_u.text_input("Usuario")
            p_in = c_p.text_input("Contraseña", type="password")
            if st.form_submit_button("Guardar acceso", type="primary"):
                if u_in and p_in:
                    st.session_state["usuario_acad"] = u_in
                    st.session_state["pass_acad"] = p_in
                    st.rerun()
                else:
                    st.error("Ingresa usuario y contraseña.")
        st.stop()

    archivo_subido = st.file_uploader(
        "Excel de participantes (.xlsx)",
        type=["xlsx", "xls"],
        key=f"uploader_{tipo_doc}",
    )
    if not archivo_subido:
        st.stop()

    df_listado, error_lectura = leer_listado_kardex(archivo_subido)
    if error_lectura:
        st.error(f"❌ {error_lectura}")
        st.stop()

    st.write(f"**Participantes listos para procesar:** {len(df_listado)}")
    st.dataframe(df_listado[["CARNET", "PROGRAMA"]], use_container_width=True, height=220)

    if st.button(texto_boton, type="primary"):
        barra = st.progress(0)
        estado = st.empty()

        def actualizar_progreso(actual, total, mensaje):
            estado.info(mensaje)
            if total > 0:
                barra.progress(min(actual / total, 1.0))

        pdf_bytes, df_log, err_proc = procesar_kardex_lote(
            df=df_listado,
            usuario=usuario_acad,
            password=pass_acad,
            tipo_documento=tipo_doc,
            on_progress=actualizar_progreso,
        )

        if err_proc == "NECESITA_LOGIN":
            st.session_state.pop("usuario_acad", None)
            st.session_state.pop("pass_acad", None)
            st.warning("La sesión caducó. Por favor ingresa tu usuario y contraseña nuevamente.")
            st.rerun()
        elif err_proc:
            estado.empty()
            if os.path.exists(ARCHIVO_SESION):
                os.remove(ARCHIVO_SESION)
            st.session_state.pop("usuario_acad", None)
            st.session_state.pop("pass_acad", None)
            st.error(err_proc)
            st.stop()

        estado.success("✅ Proceso completado.")
        st.session_state[clave_estado] = (pdf_bytes, df_log)

    if clave_estado in st.session_state:
        pdf_bytes, df_log = st.session_state[clave_estado]

        st.subheader("Resumen de extracción")
        c1, c2, c3 = st.columns(3)
        ok_count = int(df_log["Estado"].str.startswith("✅").sum())
        fail_count = len(df_log) - ok_count
        c1.metric("Total solicitados", len(df_log))
        c2.metric("Descargados", ok_count)
        c3.metric("No encontrados / Sin PAF", fail_count)

        st.dataframe(df_log, use_container_width=True)

        if pdf_bytes:
            st.download_button(
                label=f"⬇️ Descargar PDF Consolidado ({ok_count} {etiqueta_unidad} listos para imprimir)",
                data=pdf_bytes,
                file_name=nombre_archivo_pdf,
                mime="application/pdf",
                type="primary",
            )
        else:
            st.warning(f"No se pudo descargar ningún documento. Revisa la columna Estado en la tabla superior.")