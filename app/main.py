import io
import math
import os
import tempfile
import warnings
import zipfile
from datetime import datetime

import pandas as pd
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

app = FastAPI(title="ICBF Informes")

# ── Sirve el frontend desde /static ──
BASE_DIR = os.path.dirname(__file__)
STATIC_DIR = os.path.join(BASE_DIR, "static")
os.makedirs(STATIC_DIR, exist_ok=True)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


# ──────────────────────────────────────────────
# LÓGICA ICBF (portada del script v2)
# ──────────────────────────────────────────────
PERFIL = {
    "general": {
        "col_doc_nutricion":  "Numero Documento Beneficiario",
        "col_diagnostico":    "ESTADO PESO TALLA",
        "patron_alerta":      r"desnutrici[oó]n",
        "hoja_alerta":        "Alerta Desnutricion",
        "msg_sin_alerta":     "Sin casos de desnutricion en la ultima toma",
        "tipos_beneficiario": ["MENOR DE SEIS MESES", "NIÑO O NIÑA ENTRE 6 MESES Y 5 AÑOS Y 11 MESES"],
    },
    "gestante": {
        "col_doc_nutricion":  "Número documento beneficiario",
        "col_diagnostico":    "EST.NUTR. GESTANTE",
        "patron_alerta":      r"bajo peso|obesidad|sobrepeso",
        "hoja_alerta":        "Alerta Nutricional",
        "msg_sin_alerta":     "Sin alertas nutricionales en la ultima toma",
        "tipos_beneficiario": ["PERSONA GESTANTE"],
    },
}

REEMPLAZOS = str.maketrans("ÓÍÁÉÚ\n", "OIAEU ")

def norm(col):
    return col.strip().upper().translate(REEMPLAZOS).replace("  ", " ").strip()

def leer_bytes(data: bytes, filename: str, preferir_hoja=None) -> pd.DataFrame:
    ext = os.path.splitext(filename)[1].lower()
    if ext in (".xlsx", ".xls"):
        xf = pd.ExcelFile(io.BytesIO(data))
        hoja = xf.sheet_names[0]
        if preferir_hoja:
            for h in xf.sheet_names:
                if preferir_hoja.lower() in h.lower():
                    hoja = h
                    break
        return pd.read_excel(io.BytesIO(data), sheet_name=hoja, dtype=str)
    texto = data.decode("utf-8", errors="ignore")
    sep = ";" if texto.split("\n")[0].count(";") > texto.split("\n")[0].count(",") else ","
    return pd.read_csv(io.StringIO(texto), sep=sep, on_bad_lines="skip", low_memory=False, dtype=str)

def tomas_esperadas(meses, intervalo):
    return max(1, math.floor(meses / intervalo))

def procesar(df_nut_raw: bytes, fn_nut: str,
             df_act_raw: bytes | None, fn_act: str | None,
             tipo: str, intervalo_meses: int = 3) -> list[dict]:
    """
    Retorna lista de dicts:  { "contrato": str, "bytes": bytes, "filename": str }
    """
    perfil = PERFIL[tipo]
    clave_hoja = "ICBFCUEGeneralPorToma" if tipo == "general" else "GestanteLactantePorToma"

    df_nut = leer_bytes(df_nut_raw, fn_nut, preferir_hoja=clave_hoja)
    df_nut.columns = [norm(c) for c in df_nut.columns]

    COL_DIAG    = perfil["col_diagnostico"]
    COL_FECHA   = "FECHA VALORACION NURICIONAL" if tipo == "general" else "FECHA VALORACION NUTRICIONAL"
    COL_ESTADO  = "ESTADO"
    COL_CONTRATO = "NUMERO CONTRATO"
    COL_UNIDAD  = "NOMBRE UNIDAD"
    COL_DOC     = norm(perfil["col_doc_nutricion"])
    COL_NOMBRE  = "PRIMER NOMBRE BENEFICIARIO"
    COL_APELLIDO = "PRIMER APELLIDO BENEFICIARIO"

    for col in [COL_FECHA, COL_ESTADO, COL_CONTRATO, COL_DOC]:
        if col not in df_nut.columns:
            raise ValueError(f"Columna no encontrada en archivo {tipo}: '{col}'")

    df_nut[COL_ESTADO] = df_nut[COL_ESTADO].str.strip().str.upper()
    df_nut = df_nut[df_nut[COL_ESTADO] == "VINCULADO"].copy()

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df_nut[COL_FECHA] = pd.to_datetime(df_nut[COL_FECHA], errors="coerce", format="mixed", dayfirst=True)

    # ── Activos (opcional) ──
    usa_activos = df_act_raw is not None
    if usa_activos:
        df_act = leer_bytes(df_act_raw, fn_act)
        tipos_upper = [t.upper() for t in perfil["tipos_beneficiario"]]
        df_act["_tipo_upper"] = df_act["Nombre Tipo de beneficiario"].str.strip().str.upper()
        df_act_tipo = df_act[df_act["_tipo_upper"].isin(tipos_upper)].copy()
        df_act_tipo["_fecha_vinc"] = pd.to_datetime(
            df_act_tipo["Fecha de vinculación del beneficiario"], errors="coerce", dayfirst=True)
        df_act_tipo["_doc"] = df_act_tipo["Documento del beneficiario"].str.strip()
        hoy = pd.to_datetime(datetime.now())
        df_act_tipo["_meses_vinculado"] = (
            (hoy.year - df_act_tipo["_fecha_vinc"].dt.year) * 12 +
            (hoy.month - df_act_tipo["_fecha_vinc"].dt.month))
        df_act_tipo["_tomas_esperadas"] = df_act_tipo["_meses_vinculado"].apply(
            lambda m: tomas_esperadas(m, intervalo_meses) if pd.notna(m) else 0)
        conteo = df_nut.groupby(COL_DOC).size().reset_index(name="_tomas_reales")
        df_act_tipo = df_act_tipo.merge(conteo, left_on="_doc", right_on=COL_DOC, how="left")
        df_act_tipo["_tomas_reales"] = df_act_tipo["_tomas_reales"].fillna(0).astype(int)

    # ── Último registro por usuario ──
    COL_NTOMA = next(
        (c for c in df_nut.columns if "NUMERO DE TOMA" in c or "TOMA DEL BENEFICIARIO" in c),
        None)
    df_ord = df_nut.dropna(subset=[COL_FECHA]).copy()
    df_ord["_fdt"] = pd.to_datetime(df_ord[COL_FECHA], errors="coerce")
    if COL_NTOMA:
        df_ord["_nt"] = pd.to_numeric(df_ord[COL_NTOMA], errors="coerce").fillna(0)
        df_ord = df_ord.sort_values([COL_DOC, "_fdt", "_nt"])
    else:
        df_ord = df_ord.sort_values([COL_DOC, "_fdt"])
    df_ultimo = df_ord.drop_duplicates(subset=[COL_DOC], keep="last")

    resultados = []
    for contrato in df_nut[COL_CONTRATO].dropna().unique():
        df_c    = df_nut[df_nut[COL_CONTRATO] == contrato]
        df_ul_c = df_ultimo[df_ultimo[COL_CONTRATO] == contrato]

        # Hoja 1
        hoja_unidades = (df_c.groupby(COL_UNIDAD)[COL_DOC]
                         .nunique().reset_index()
                         .rename(columns={COL_DOC: "TOTAL USUARIOS UNICOS"}))

        # Hoja 2: tomas faltantes
        if usa_activos:
            df_act_c = df_act_tipo[
                df_act_tipo["Número del Contrato"].str.strip() == str(contrato).strip()
            ] if "Número del Contrato" in df_act_tipo.columns else df_act_tipo

            sin_toma   = df_act_c[df_act_c["_tomas_reales"] == 0]
            con_deficit = df_act_c[
                (df_act_c["_tomas_reales"] > 0) &
                (df_act_c["_tomas_reales"] < df_act_c["_tomas_esperadas"])]
            faltantes_src = pd.concat([sin_toma, con_deficit], ignore_index=True)

            filas = []
            for _, row in faltantes_src.iterrows():
                doc = row["_doc"]
                rn  = df_ul_c[df_ul_c[COL_DOC] == doc]
                filas.append({
                    "UNIDAD":             rn[COL_UNIDAD].values[0] if not rn.empty else row.get("Nombre de la unidad de servicio", ""),
                    "DOCUMENTO":          doc,
                    "NOMBRE":             rn[COL_NOMBRE].values[0] if not rn.empty else row.get("Primer Nombre del beneficiario", ""),
                    "APELLIDO":           rn[COL_APELLIDO].values[0] if not rn.empty else row.get("Primer apellido del beneficiario", ""),
                    "TOMAS REALIZADAS":   int(row["_tomas_reales"]),
                    "TOMAS ESPERADAS":    int(row["_tomas_esperadas"]),
                    "MESES VINCULADO":    int(row["_meses_vinculado"]) if pd.notna(row.get("_meses_vinculado")) else 0,
                    "ULTIMO DIAGNOSTICO": rn[COL_DIAG].values[0] if not rn.empty and COL_DIAG in rn.columns else "SIN TOMA",
                    "MOTIVO":             "Sin toma registrada" if row["_tomas_reales"] == 0 else "Tomas insuficientes",
                })
            df_faltantes = pd.DataFrame(filas)
        else:
            # Modo básico: detectar por historial de tomas
            hoy = pd.to_datetime(datetime.now())
            ids_faltantes = set()
            for usuario, grupo in df_ord[df_ord[COL_CONTRATO] == contrato].groupby(COL_DOC):
                fechas = grupo["_fdt"].dt.to_period("M").unique()
                if len(fechas) > 1:
                    for i in range(1, len(fechas)):
                        diff = (fechas[i].year - fechas[i-1].year)*12 + (fechas[i].month - fechas[i-1].month)
                        if diff > intervalo_meses:
                            ids_faltantes.add(usuario)
                            break
                if usuario not in ids_faltantes:
                    ultima = fechas[-1]
                    diff_hoy = (hoy.year - ultima.year)*12 + (hoy.month - ultima.month)
                    if diff_hoy >= intervalo_meses:
                        ids_faltantes.add(usuario)

            cols_f = [COL_UNIDAD, COL_DOC, COL_NOMBRE, COL_APELLIDO]
            if COL_DIAG in df_ul_c.columns:
                cols_f.append(COL_DIAG)
            df_faltantes = (df_ul_c[df_ul_c[COL_DOC].isin(ids_faltantes)][cols_f].copy()
                            if ids_faltantes else pd.DataFrame(columns=cols_f))

        # Hoja 3: alertas
        if COL_DIAG in df_ul_c.columns:
            mask = df_ul_c[COL_DIAG].str.contains(perfil["patron_alerta"], case=False, na=False, regex=True)
            cols_a = [COL_UNIDAD, COL_DOC, COL_NOMBRE, COL_APELLIDO, COL_DIAG, COL_FECHA]
            df_alerta = df_ul_c[mask][cols_a].copy()
            df_alerta[COL_FECHA] = df_alerta[COL_FECHA].dt.strftime("%Y-%m-%d")
        else:
            df_alerta = pd.DataFrame()

        # Serializar a bytes
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            hoja_unidades.to_excel(writer, sheet_name="Usuarios por Unidad", index=False)
            (df_faltantes if not df_faltantes.empty
             else pd.DataFrame({"MENSAJE": ["Sin tomas faltantes"]})).to_excel(
                writer, sheet_name="Tomas Faltantes", index=False)
            (df_alerta if not df_alerta.empty
             else pd.DataFrame({"MENSAJE": [perfil["msg_sin_alerta"]]})).to_excel(
                writer, sheet_name=perfil["hoja_alerta"], index=False)
        buf.seek(0)

        resultados.append({
            "contrato": str(contrato),
            "tipo": tipo,
            "n_vinculados": int(df_c[COL_DOC].nunique()),
            "n_faltantes": len(df_faltantes),
            "n_alertas": len(df_alerta),
            "bytes": buf.read(),
            "filename": f"Informe_{tipo.upper()}_Contrato_{contrato}.xlsx",
        })

    return resultados


# ──────────────────────────────────────────────
# ENDPOINTS
# ──────────────────────────────────────────────
@app.get("/", response_class=HTMLResponse)
async def index():
    with open(os.path.join(STATIC_DIR, "index.html"), encoding="utf-8") as f:
        return f.read()


@app.post("/procesar")
async def procesar_endpoint(
    modo: str,
    general: UploadFile = File(...),
    gestante: UploadFile = File(...),
    activos: UploadFile = File(None),
):
    if modo not in ("basico", "completo"):
        raise HTTPException(400, "modo debe ser 'basico' o 'completo'")
    if modo == "completo" and activos is None:
        raise HTTPException(400, "Modo completo requiere el archivo de activos")

    try:
        general_bytes  = await general.read()
        gestante_bytes = await gestante.read()
        activos_bytes  = (await activos.read()) if activos else None
        activos_fn     = activos.filename if activos else None

        resultados_gen  = procesar(general_bytes,  general.filename,  activos_bytes, activos_fn, "general")
        resultados_gest = procesar(gestante_bytes, gestante.filename, activos_bytes, activos_fn, "gestante")
        todos = resultados_gen + resultados_gest

    except ValueError as e:
        raise HTTPException(422, str(e))
    except Exception as e:
        raise HTTPException(500, f"Error al procesar: {e}")

    # Empaquetar en ZIP
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in todos:
            zf.writestr(r["filename"], r["bytes"])
    zip_buf.seek(0)

    # Resumen para el header JSON
    resumen = [
        {k: v for k, v in r.items() if k != "bytes"}
        for r in todos
    ]
    import json
    resumen_header = json.dumps(resumen, ensure_ascii=False)

    return StreamingResponse(
        zip_buf,
        media_type="application/zip",
        headers={
            "Content-Disposition": "attachment; filename=informes_icbf.zip",
            "X-Resumen": resumen_header,
        },
    )
