"""Bulk Renamer: renombra archivos a partir de una lista  nombre actual -> nombre nuevo.

Validaciones (sanitize_base, detect_collisions) portadas de bulk-rename-core
(MIT) https://github.com/adityasingh031901-star/bulk-rename-core
"""
import io
import os
import re
import uuid
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Bulk Renamer", page_icon="📁", layout="wide")

COLS = ["Nombre actual", "Nombre nuevo"]
HEADERS = {"old", "old name", "nombre actual", "actual", "original", "nombre", "name"}

# ---------------------------------------------------------------------------
# Port de bulk-rename-core
# ---------------------------------------------------------------------------
ILLEGAL = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
WINDOWS_RESERVED = re.compile(r"^(con|prn|aux|nul|com[1-9]|lpt[1-9])$", re.I)


def get_extension(name: str) -> str:
    dot = name.rfind(".")
    return name[dot:] if 0 < dot < len(name) - 1 else ""


def get_base(name: str) -> str:
    dot = name.rfind(".")
    return name[:dot] if dot > 0 else name


def sanitize_base(base: str) -> str:
    """Quita caracteres ilegales en Windows/macOS, puntos/espacios finales,
    protege nombres reservados (CON, NUL, COM1…) y limita a 180 caracteres."""
    s = ILLEGAL.sub("", base or "")
    s = re.sub(r"[\s.]+$", "", s).strip()
    if WINDOWS_RESERVED.match(s):
        s = "_" + s
    if len(s) > 180:
        s = s[:180].strip()
    return s or "unnamed"


def detect_collisions(names: list[str]) -> set[int]:
    """Índices de nombres duplicados, sin distinguir mayúsculas."""
    seen: dict[str, list[int]] = {}
    for i, n in enumerate(names):
        seen.setdefault(n.lower(), []).append(i)
    return {i for idxs in seen.values() if len(idxs) > 1 for i in idxs}


# ---------------------------------------------------------------------------
# Lógica de la lista
# ---------------------------------------------------------------------------
def looks_like_ext(ext: str) -> bool:
    return 1 < len(ext) <= 6 and ext[1:].isalnum() and any(c.isalpha() for c in ext)


def build_new_name(old: str, new: str, keep_ext: bool) -> str:
    ext = get_extension(new)
    if looks_like_ext(ext):
        base = get_base(new)
    else:
        base, ext = new, ""
        if keep_ext:
            ext = get_extension(old)
    return sanitize_base(base) + ILLEGAL.sub("", ext)[:24]


def parse_list(text: str) -> list[tuple[str, str]]:
    """Una línea por archivo. Separador: tab (pegado de Excel), ; , -> o ,"""
    pares = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        for sep in ("\t", ";", "->", "→", ","):
            if sep in line:
                old, new = line.split(sep, 1)
                break
        else:
            continue
        old, new = old.strip().strip('"'), new.strip().strip('"')
        if old.lower() in HEADERS:
            continue
        pares.append((old, new))
    return pares


def build_plan(files: list[str], pares: list[tuple[str, str]], keep_ext: bool):
    exact = set(files)
    lower = {f.lower(): f for f in files}
    by_base: dict[str, list[str]] = {}
    for f in files:
        by_base.setdefault(get_base(f).lower(), []).append(f)

    rows, targets = [], {}
    for old, new in pares:
        if not old or not new:
            continue
        actual = old if old in exact else lower.get(old.lower())
        if actual is None and len(by_base.get(old.lower(), [])) == 1:
            actual = by_base[old.lower()][0]  # lista sin extensión
        if actual is None:
            rows.append({"Actual": old, "Nuevo": new, "Estado": "❌ No está entre los archivos"})
            continue
        if actual in targets:
            rows.append({"Actual": actual, "Nuevo": new, "Estado": "❌ Repetido en la lista"})
            continue
        final = build_new_name(actual, new, keep_ext)
        targets[actual] = final
        rows.append({"Actual": actual, "Nuevo": final, "Estado": None, "_raw": new})

    finales = [targets.get(f, f) for f in files]
    choques = {files[i] for i in detect_collisions(finales)}

    plan = []
    for r in rows:
        if r["Estado"]:
            continue
        actual, final = r["Actual"], r["Nuevo"]
        if final == actual:
            r["Estado"] = "➖ Sin cambio"
        elif actual in choques:
            r["Estado"] = "❌ Choca con otro nombre (mayúsc./minúsc. incluidas)"
        else:
            raw = r["_raw"]
            ajustado = final not in (raw, raw + get_extension(actual))
            r["Estado"] = "✅ OK (nombre limpiado)" if ajustado else "✅ OK"
            plan.append((actual, final))
    df = pd.DataFrame(rows).drop(columns="_raw", errors="ignore")
    return df, plan


def rename_on_disk(base: Path, plan):
    """Dos pasos vía temporales: permite intercambios A↔B y cambios solo de mayúsculas."""
    temporales = []
    for viejo, nuevo in plan:
        tmp = base / f".__tmp_{uuid.uuid4().hex}"
        os.rename(base / viejo, tmp)
        temporales.append((tmp, nuevo, viejo))
    hechos = []
    for tmp, nuevo, viejo in temporales:
        os.rename(tmp, base / nuevo)
        hechos.append((nuevo, viejo))
    return hechos


def is_local() -> bool:
    try:
        host = st.context.headers.get("Host", "") or ""
    except Exception:
        return True
    return host.split(":")[0] in ("", "localhost", "127.0.0.1")


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("📁 Bulk Renamer")
if "historial" not in st.session_state:
    st.session_state.historial = []

modos = ["Subir archivos (descargas un ZIP)"]
if is_local():
    modos.append("Carpeta de mi PC (renombra en el sitio)")
modo = st.radio("Modo", modos, horizontal=True)
local = modo.startswith("Carpeta")

# 1. Archivos
st.subheader("1. Archivos")
files: list[str] = []
uploads = {}
carpeta = None
if local:
    ruta = st.text_input("Ruta de la carpeta", placeholder=r"C:\Users\Nicolas\Documents\facturas")
    if ruta:
        carpeta = Path(ruta.strip().strip('"'))
        if carpeta.is_dir():
            files = sorted(f.name for f in carpeta.iterdir() if f.is_file() and f.name != "rename_log.csv")
            st.caption(f"{len(files)} archivos en la carpeta.")
        else:
            st.error("La carpeta no existe.")
else:
    subidos = st.file_uploader("Arrastra los archivos", accept_multiple_files=True)
    uploads = {f.name: f for f in subidos or []}
    files = list(uploads)

if files:
    plantilla = pd.DataFrame({COLS[0]: files, COLS[1]: ""}).to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Plantilla CSV con los nombres actuales", plantilla, "plantilla_renombrar.csv")

# 2. Lista
st.subheader("2. Lista de nombre actual → nombre nuevo")
keep_ext = st.checkbox("Mantener la extensión si el nombre nuevo no la trae", value=True)
metodo = st.radio("¿Cómo cargas la lista?", ["Pegar lista", "Tabla", "Excel / CSV"], horizontal=True)

pares: list[tuple[str, str]] = []
if metodo == "Pegar lista":
    texto = st.text_area(
        "Una línea por archivo. Pega dos columnas de Excel o usa coma / punto y coma.",
        height=220,
        placeholder="IMG_001.jpg,factura_enero\nIMG_002.jpg;factura_febrero\nscan 3.pdf\tcontrato_cliente",
    )
    pares = parse_list(texto)
elif metodo == "Tabla":
    df = pd.DataFrame({COLS[0]: files, COLS[1]: [""] * len(files)})
    editado = st.data_editor(df, num_rows="dynamic", width="stretch", key=f"ed_{hash(tuple(files))}")
    pares = [(str(a), str(b)) for a, b in editado.fillna("").itertuples(index=False)]
else:
    hoja = st.file_uploader("Columna A: nombre actual · Columna B: nombre nuevo", type=["xlsx", "xls", "csv"])
    if hoja is not None:
        if hoja.name.lower().endswith(".csv"):
            raw = hoja.getvalue().decode("utf-8-sig", errors="replace")
            pares = parse_list(raw)
        else:
            d = pd.read_excel(hoja, header=None, dtype=str).iloc[:, :2].fillna("")
            pares = [(a.strip(), b.strip()) for a, b in d.itertuples(index=False)
                     if a.strip().lower() not in HEADERS]

# 3. Vista previa
st.subheader("3. Vista previa")
if not files:
    st.info("Primero carga los archivos.")
elif not pares:
    st.info("Agrega al menos una fila a la lista.")
else:
    preview, plan = build_plan(files, pares, keep_ext)
    st.dataframe(preview, width="stretch", hide_index=True)
    errores = int(preview["Estado"].str.startswith("❌").sum())
    st.write(f"**{len(plan)}** listos · **{errores}** con error (no se tocan)")

    log = pd.DataFrame(plan, columns=["Anterior", "Nuevo"])
    if local:
        if st.button("🚀 Renombrar", type="primary", disabled=not plan):
            try:
                hechos = rename_on_disk(carpeta, plan)
                st.session_state.historial.append((str(carpeta), hechos))
                log.to_csv(carpeta / "rename_log.csv", index=False)
                st.success(f"Listo: {len(hechos)} archivos renombrados. Log en rename_log.csv")
            except Exception as e:
                st.error(f"Error: {e}")
    else:
        solo = st.checkbox("Incluir solo los archivos renombrados", value=False)
        if plan:
            mapa = dict(plan)
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                for nombre, f in uploads.items():
                    if solo and nombre not in mapa:
                        continue
                    z.writestr(mapa.get(nombre, nombre), f.getvalue())
                z.writestr("rename_log.csv", log.to_csv(index=False))
            st.download_button("⬇️ Descargar ZIP", buf.getvalue(), "renombrados.zip",
                               "application/zip", type="primary")

if local and st.session_state.historial:
    if st.button("↩️ Deshacer último renombrado"):
        carpeta_h, hechos = st.session_state.historial.pop()
        try:
            rename_on_disk(Path(carpeta_h), hechos)
            st.success("Renombrado revertido.")
        except Exception as e:
            st.error(f"No se pudo deshacer: {e}")
