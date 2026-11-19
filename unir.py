"""Unir Excels de valores asignados y revisar correlatividad y repeticiones.

Lógica pura (no usa Streamlit): se puede probar o usar desde otro programa.
"""
import io
import re
import unicodedata
from bisect import bisect_left, bisect_right

import pandas as pd

# ---------------------------------------------------------------- nombres de columnas
CANON = {
    "nro": "Nro.", "fecha_deposito": "FECHA DEPÓSITO", "nro_deposito": "No. DEPÓSITO",
    "nombre": "NOMBRES Y APELLIDOS", "monto": "DEPÓSITO",
    **{f"cert{i}": f"CERT.NOTAS {i}" for i in range(1, 7)},
    "historial": "HISTORIAL ACAD.", "hoja_reg": "HOJA DE REG. U", "titulo": "TÍTULO",
}
ALIAS = {
    "NRO": "nro", "FECHADEPOSITO": "fecha_deposito",
    "NODEPOSITO": "nro_deposito", "NRODEPOSITO": "nro_deposito", "NDEPOSITO": "nro_deposito",
    "NUMERODEPOSITO": "nro_deposito",
    "NOMBRESYAPELLIDOS": "nombre", "NOMBREYAPELLIDO": "nombre", "NOMBRES": "nombre", "NOMBRE": "nombre",
    "DEPOSITO": "monto",
    **{f"CERTNOTAS{i}": f"cert{i}" for i in range(1, 7)},
    "HISTORIALACAD": "historial", "HISTORIALACADEMICO": "historial",
    "HOJADEREGU": "hoja_reg", "HOJADEREGISTROUNIVERSITARIO": "hoja_reg",
    "TITULO": "titulo",
}
IMPORTANTES = {"nombre", "historial", "hoja_reg", "titulo"} | {f"cert{i}" for i in range(1, 7)}
SERIES = [
    ("Certificados de notas", [f"cert{i}" for i in range(1, 7)]),  # una sola cadena del 1 al 6
    ("Historial académico", ["historial"]),
    ("Hoja de registro universitario", ["hoja_reg"]),
    ("Título", ["titulo"]),
]
COL_ARCHIVO, COL_FILA = "ARCHIVO ORIGEN", "FILA EXCEL"


def _sin_acentos(t):
    return "".join(c for c in unicodedata.normalize("NFD", str(t)) if unicodedata.category(c) != "Mn")


def _clave(t):
    return re.sub(r"[^A-Z0-9]", "", _sin_acentos(t).upper())


def _limpiar(t):
    return re.sub(r"[^A-Z0-9]+", " ", _sin_acentos(t).upper()).strip()


def _texto(x):
    return x is not None and not pd.isna(x) and str(x).strip() != ""


def _entero(x):
    n = pd.to_numeric(x, errors="coerce")
    return 0 if pd.isna(n) else int(round(n))


def _rangos(nums):
    nums = sorted(nums)
    out, ini, ant = [], None, None
    for n in nums:
        if ini is None:
            ini = ant = n
        elif n == ant + 1:
            ant = n
        else:
            out.append((ini, ant))
            ini = ant = n
    if ini is not None:
        out.append((ini, ant))
    return out


def _txt_rangos(rangos, maximo=8):
    partes = [str(a) if a == b else f"{a}–{b}" for a, b in rangos[:maximo]]
    if len(rangos) > maximo:
        partes.append(f"… y {len(rangos) - maximo} más")
    return ", ".join(partes)


# ---------------------------------------------------------------- lectura y unión
def _buscar_encabezado(raw):
    for i in range(min(15, len(raw))):
        hallados = {ALIAS.get(_clave(c)) for c in raw.iloc[i] if _texto(c)} & IMPORTANTES
        if len(hallados) >= 3:
            return i
    return None


def leer_excels(archivos):
    """archivos: lista de (nombre, archivo). Devuelve (DataFrame unido, mensajes de lectura, problemas de lectura)."""
    partes, mensajes, errores = [], [], []
    for nombre, f in archivos:
        try:
            if hasattr(f, "seek"):
                f.seek(0)
            hojas = pd.read_excel(f, sheet_name=None, header=None, dtype=object, engine="openpyxl")
        except Exception as e:
            errores.append(f"No se pudo leer «{nombre}»: {e}")
            continue
        usadas = 0
        for hoja, raw in hojas.items():
            i = _buscar_encabezado(raw)
            if i is None:
                continue
            columnas, vistos = {}, set()
            for j, c in enumerate(raw.iloc[i]):
                if not _texto(c):
                    continue
                k = ALIAS.get(_clave(c))
                nom = CANON[k] if k else str(c).strip()
                if nom not in vistos:
                    vistos.add(nom)
                    columnas[j] = nom
            datos = raw.iloc[i + 1:, list(columnas)].copy()
            datos.columns = list(columnas.values())
            datos = datos.dropna(how="all")
            if datos.empty:
                continue
            datos.insert(0, COL_FILA, datos.index + 1)  # fila real en Excel
            datos.insert(0, COL_ARCHIVO, nombre if len(hojas) == 1 else f"{nombre} [{hoja}]")
            partes.append(datos)
            usadas += 1
            mensajes.append(f"«{nombre}» hoja «{hoja}»: {len(datos)} filas")
        if usadas == 0:
            errores.append(f"«{nombre}»: no encontré una hoja con encabezados reconocibles "
                           "(CERT.NOTAS, HISTORIAL ACAD., HOJA DE REG. U, TÍTULO…).")
    if not partes:
        return pd.DataFrame(), mensajes, errores
    return pd.concat(partes, ignore_index=True), mensajes, errores


# ---------------------------------------------------------------- revisión
def _clave_participante(nombre):
    """Identifica al participante por el CI que va después de ':' o, si no hay, por el nombre."""
    if not _texto(nombre):
        return None
    s = str(nombre).strip()
    m = re.search(r":\s*(\d[\dA-Za-z-]*)\s*$", s)
    return "CI " + re.sub(r"\W", "", m.group(1)) if m else "N " + _limpiar(s)


def auditar(df):
    problemas, series = [], []
    c = lambda k: CANON[k] if CANON[k] in df.columns else None  # noqa: E731

    def add(grav, tipo, detalle, idx=None):
        problemas.append({
            "Gravedad": grav, "Tipo": tipo, "Detalle": detalle,
            "Archivo": df.at[idx, COL_ARCHIVO] if idx is not None else "",
            "Fila Excel": df.at[idx, COL_FILA] if idx is not None else None, "_idx": idx})

    def donde(idx):
        return f"{df.at[idx, COL_ARCHIVO]} fila {df.at[idx, COL_FILA]}"

    col_nombre = c("nombre")
    df["_pk"] = df[col_nombre].map(_clave_participante) if col_nombre else None
    if not col_nombre:
        add("AVISO", "Falta columna", "No encontré la columna NOMBRES Y APELLIDOS: no se puede revisar a los participantes.")

    # ---- 1. series de valores: correlatividad y repetidos
    for titulo, claves in SERIES:
        cols = [CANON[k] for k in claves if CANON[k] in df.columns]
        if not cols:
            add("AVISO", "Falta columna", f"{titulo}: no encontré la(s) columna(s) en los archivos.")
            series.append({"serie": titulo, "n": 0})
            continue
        partes = []
        for col in cols:
            s = pd.to_numeric(df[col], errors="coerce")
            ok = s.notna() & (s != 0)
            partes.append(pd.DataFrame({"idx": s[ok].index, "valor": s[ok].round().astype("int64").values, "col": col}))
        v = pd.concat(partes, ignore_index=True)
        if v.empty:
            add("AVISO", "Sin datos", f"{titulo}: no hay valores numéricos.")
            series.append({"serie": titulo, "n": 0})
            continue
        v["pk"] = v["idx"].map(df["_pk"])
        unicos = sorted(v["valor"].unique())
        mn, mx = unicos[0], unicos[-1]
        info = {"serie": titulo, "n": len(unicos), "min": mn, "max": mx, "faltan": 0, "rangos": [], "repetidos": 0}

        if mx - mn > 5_000_000:
            add("AVISO", "Valor atípico", f"{titulo}: el rango {mn}–{mx} es enorme; revisa si hay un número mal escrito.")
        else:
            falt = _rangos(set(range(mn, mx + 1)) - set(unicos))
            info["rangos"], info["faltan"] = falt, sum(b - a + 1 for a, b in falt)
            for a, b in falt[:100]:
                ant = unicos[bisect_left(unicos, a) - 1]
                sig = unicos[bisect_right(unicos, b)]
                idx = int(v.loc[v["valor"] == ant, "idx"].iloc[0])
                add("ERROR", "Salto en la secuencia",
                    f"{titulo}: faltan {b - a + 1} número(s): {a if a == b else f'{a}–{b}'} (entre el {ant} y el {sig}).", idx)

        for valor, g in v[v.duplicated("valor", keep=False)].groupby("valor"):
            filas = list(dict.fromkeys(g["idx"]))
            if len(filas) == 1:
                for idx in filas:
                    add("ERROR", "Valor repetido en la misma fila", f"{titulo}: el valor {valor} está en dos columnas de la misma fila.", idx)
                info["repetidos"] += 1
                continue
            if g["pk"].notna().all() and g["pk"].nunique() == 1:
                continue  # mismo participante en varias filas: se informa en "Participante repetido"
            info["repetidos"] += 1
            lugares = "; ".join(donde(i) for i in filas[:4]) + ("…" if len(filas) > 4 else "")
            for idx in filas:
                add("ERROR", "Valor repetido en participantes distintos",
                    f"{titulo}: el valor {valor} aparece en {len(filas)} filas ({lugares}).", idx)
        series.append(info)

    # ---- 2. revisión fila por fila
    certs = [CANON[f"cert{i}"] for i in range(1, 7) if CANON[f"cert{i}"] in df.columns]
    otros = [(CANON[k], n) for k, n in (("historial", "HISTORIAL ACAD."), ("hoja_reg", "HOJA DE REG. U"), ("titulo", "TÍTULO"))
             if CANON[k] in df.columns]
    firmas = {}
    for idx in df.index:
        vals = [_entero(df.at[idx, col]) for col in certs]
        nz = [x for x in vals if x]
        if certs:
            if not nz:
                add("AVISO", "Sin certificados de notas", "La fila no tiene ningún CERT.NOTAS.", idx)
            else:
                if any(x == 0 and any(vals[j] for j in range(i + 1, len(vals))) for i, x in enumerate(vals)):
                    add("AVISO", "Certificados con hueco", f"Hay un CERT.NOTAS vacío entre otros llenos: {vals}.", idx)
                if any(b - a != 1 for a, b in zip(nz, nz[1:])):
                    add("AVISO", "Certificados no consecutivos", f"Los certificados de la fila no son consecutivos: {nz}.", idx)
        falta = [n for col, n in otros if not _entero(df.at[idx, col])]
        if falta:
            add("AVISO", "Falta un valor", f"Falta: {', '.join(falta)}.", idx)
        if col_nombre and df.at[idx, "_pk"] is None:
            add("AVISO", "Fila sin nombre", "La fila no tiene nombre de participante.", idx)
        firmas[idx] = tuple(vals) + tuple(_entero(df.at[idx, col]) for col, _ in otros)

    # ---- 3. participantes repetidos
    col_dep = c("nro_deposito")
    n_igual = n_dist = 0
    if col_nombre:
        for pk, g in df[df["_pk"].notna()].groupby("_pk"):
            if len(g) < 2:
                continue
            nombre = str(g.iloc[0][col_nombre]).strip()
            deps = ", ".join(str(x) for x in g[col_dep].dropna()) if col_dep else ""
            lugares = "; ".join(donde(i) for i in g.index[:4])
            if len({firmas[i] for i in g.index}) == 1:
                n_igual += 1
                for idx in g.index:
                    add("AVISO", "Participante repetido (mismos valores)",
                        f"{nombre} aparece {len(g)} veces con los mismos valores ({lugares}). Si es un depósito adicional "
                        f"es normal; sus valores se cuentan una sola vez. Depósitos: {deps}.", idx)
            else:
                n_dist += 1
                for idx in g.index:
                    add("ERROR", "Participante repetido con valores distintos",
                        f"{nombre} aparece {len(g)} veces con valores diferentes ({lugares}).", idx)

    # ---- 4. mismo depósito en participantes distintos
    if col_dep and col_nombre:
        dep = df[col_dep].map(lambda x: re.sub(r"\.0$", "", str(x).strip()) if _texto(x) else None)
        for d, g in df.groupby(dep):
            if len(g) > 1 and g["_pk"].nunique() > 1:
                for idx in g.index:
                    add("AVISO", "Mismo depósito en participantes distintos",
                        f"El No. DEPÓSITO {d} está en {len(g)} filas de participantes distintos (puede ser un depósito grupal).", idx)

    # ---- 5. columna Nro. dentro de cada archivo
    col_nro = c("nro")
    if col_nro:
        for arch, g in df.groupby(COL_ARCHIVO, sort=False):
            s = pd.to_numeric(g[col_nro], errors="coerce").dropna().astype(int)
            if s.empty:
                continue
            falt = _rangos(set(range(int(s.min()), int(s.max()) + 1)) - set(s))
            dup = sorted(s[s.duplicated()].unique())
            if falt or dup:
                txt = (f"faltan {_txt_rangos(falt)}" if falt else "") + ("; " if falt and dup else "") + \
                      (f"repetidos {dup[:8]}" if dup else "")
                add("AVISO", "Nro. con saltos o repetidos", f"{arch}: la columna Nro. tiene {txt}.", g.index[0])

    resumen_part = {"unico": df["_pk"].nunique() if col_nombre else 0, "igual": n_igual, "distintos": n_dist}
    return pd.DataFrame(problemas, columns=["Gravedad", "Tipo", "Detalle", "Archivo", "Fila Excel", "_idx"]), series, resumen_part


# ---------------------------------------------------------------- resultado y mensaje
def unir_y_auditar(archivos):
    df, mensajes, errores_lectura = leer_excels(archivos)
    if df.empty:
        return {"df": df, "problemas": pd.DataFrame(), "series": [], "partic": {}, "lecturas": mensajes,
                "errores_lectura": errores_lectura, "n_archivos": 0}
    problemas, series, partic = auditar(df)
    # columna de observaciones por fila + gravedad para pintar
    con_fila = problemas.dropna(subset=["_idx"])
    obs = con_fila.groupby("_idx")["Tipo"].agg(lambda s: "; ".join(dict.fromkeys(s)))
    grav = con_fila.groupby("_idx")["Gravedad"].agg(lambda s: "ERROR" if "ERROR" in set(s) else "AVISO")
    df = df.drop(columns=["_pk"])
    df["OBSERVACIÓN"] = df.index.map(obs).fillna("")
    res = {"df": df, "problemas": problemas, "series": series, "partic": partic, "lecturas": mensajes,
           "errores_lectura": errores_lectura, "n_archivos": df[COL_ARCHIVO].str.replace(r" \[.*\]$", "", regex=True).nunique(),
           "gravedad_fila": grav.to_dict()}
    res["resumen"] = redactar_resumen(res)
    return res


def contar(res):
    p = res["problemas"]
    err = int((p["Gravedad"] == "ERROR").sum()) + len(res["errores_lectura"])
    return err, int((p["Gravedad"] == "AVISO").sum())


def redactar_resumen(res):
    err, avi = contar(res)
    if err == 0 and avi == 0:
        veredicto = "✅ TODO CORRECTO: los valores son correlativos y no hay repeticiones."
    elif err == 0:
        veredicto = f"⚠️ SIN ERRORES GRAVES, pero hay {avi} aviso(s) por revisar."
    else:
        veredicto = f"❌ SE ENCONTRARON {err} ERROR(ES) y {avi} aviso(s). Revisa antes de firmar el acta."
    n_filas, n_part = len(res["df"]), res["partic"].get("unico", 0)
    L = [veredicto, "", f"Se unieron {res['n_archivos']} archivo(s): {n_filas} filas y {n_part} participantes distintos.", ""]
    for e in res["errores_lectura"]:
        L.append(f"❌ {e}")
    for s in res["series"]:
        if not s.get("n"):
            L.append(f"➖ {s['serie']}: sin datos.")
            continue
        base = f"{s['serie']}: del {s['min']} al {s['max']} ({s['n']} números)"
        malo = []
        if s["faltan"]:
            malo.append(f"faltan {s['faltan']} ({_txt_rangos(s['rangos'])})")
        if s["repetidos"]:
            malo.append(f"{s['repetidos']} valor(es) repetido(s)")
        L.append(f"❌ {base} — " + "; ".join(malo) + "." if malo else f"✅ {base} — correlativo, sin saltos ni repetidos.")
    p = res["partic"]
    if p.get("distintos"):
        L.append(f"❌ {p['distintos']} participante(s) repetido(s) con valores DISTINTOS.")
    if p.get("igual"):
        L.append(f"⚠️ {p['igual']} participante(s) en más de una fila con los mismos valores (normalmente, un depósito adicional).")
    otros = res["problemas"]
    otros = otros[~otros["Tipo"].isin(["Salto en la secuencia", "Valor repetido en participantes distintos",
                                       "Valor repetido en la misma fila", "Participante repetido (mismos valores)",
                                       "Participante repetido con valores distintos"])]
    if not otros.empty:
        L += ["", "Otros avisos:"]
        for tipo, n in otros["Tipo"].value_counts().items():
            L.append(f"• {tipo}: {n}")
    return "\n".join(L)


def exportar_excel(res):
    """Devuelve los bytes de un Excel con las hojas Resumen, Consolidado (filas con problemas pintadas) y Problemas."""
    from openpyxl.styles import PatternFill
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame({"RESUMEN DE LA REVISIÓN": res["resumen"].split("\n")}).to_excel(w, sheet_name="Resumen", index=False)
        res["df"].to_excel(w, sheet_name="Consolidado", index=False)
        res["problemas"].drop(columns=["_idx"]).to_excel(w, sheet_name="Problemas", index=False)
        rojo, amarillo = PatternFill("solid", fgColor="F8CBAD"), PatternFill("solid", fgColor="FFF2CC")
        ws = w.sheets["Consolidado"]
        for idx, g in res["gravedad_fila"].items():
            for cell in ws[int(idx) + 2]:
                cell.fill = rojo if g == "ERROR" else amarillo
        for ws in w.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                ancho = max((len(str(c.value)) for c in col[:200] if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(max(ancho + 2, 8), 70)
    return buf.getvalue()
