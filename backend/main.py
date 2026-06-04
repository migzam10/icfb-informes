"""
ICBF Informes - Backend FastAPI
Modos:
  - basico:    solo General + Gestante (sin archivo de activos)
  - completo:  General + Gestante + BeneficiariosPIActivos (análisis de déficit)
"""

import io
import json
import math
import warnings
import zipfile
from datetime import datetime

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles

app = FastAPI(title="ICBF Informes")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["X-Resumen"],
)

PERFIL = {
    "general": {
        "col_doc":        "Numero Documento Beneficiario",
        "col_diag":       "ESTADO PESO TALLA",
        "patron_alerta":  r"desnutrici[oó]n",
        "hoja_alerta":    "Alerta Desnutricion",
        "msg_sin_alerta": "Sin casos de desnutricion en la ultima toma",
        "tipos_activos":  ["MENOR DE SEIS MESES", "NIÑO O NIÑA ENTRE 6 MESES Y 5 AÑOS Y 11 MESES"],
        "col_fecha":      "FECHA VALORACION NURICIONAL",
        "hoja_excel":     "ICBFCUEGeneralPorToma",
    },
    "gestante": {
        "col_doc":        "Número documento beneficiario",
        "col_diag":       "EST.NUTR. GESTANTE",
        "patron_alerta":  r"bajo peso|obesidad|sobrepeso",
        "hoja_alerta":    "Alerta Nutricional",
        "msg_sin_alerta": "Sin alertas nutricionales en la ultima toma",
        "tipos_activos":  ["PERSONA GESTANTE"],
        "col_fecha":      "FECHA VALORACION NUTRICIONAL",
        "hoja_excel":     "GestanteLactantePorToma",
    },
}

REEMPLAZOS = str.maketrans("ÓÍÁÉÚ\n", "OIAEU ")
MAX_UPLOAD_MB = 30


def norm(col: str) -> str:
    return col.strip().upper().translate(REEMPLAZOS).replace("  ", " ").strip()


def leer_bytes(data: bytes, filename: str, preferir_hoja: str = None) -> pd.DataFrame:
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext in ("xlsx", "xls"):
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


def tomas_esperadas(meses: float, intervalo: int) -> int:
    return max(1, math.floor(meses / intervalo))


def procesar(
    nut_bytes: bytes, nut_fn: str,
    act_bytes: bytes | None, act_fn: str | None,
    tipo: str, intervalo: int,
) -> list[dict]:
    p = PERFIL[tipo]
    df = leer_bytes(nut_bytes, nut_fn, preferir_hoja=p["hoja_excel"])
    df.columns = [norm(c) for c in df.columns]

    COL_DOC      = norm(p["col_doc"])
    COL_DIAG     = p["col_diag"]
    COL_FECHA    = p["col_fecha"]
    COL_ESTADO   = "ESTADO"
    COL_CONTRATO = "NUMERO CONTRATO"
    COL_UNIDAD   = "NOMBRE UNIDAD"
    COL_NOMBRE   = "PRIMER NOMBRE BENEFICIARIO"
    COL_APELLIDO = "PRIMER APELLIDO BENEFICIARIO"

    for col in [COL_FECHA, COL_ESTADO, COL_CONTRATO, COL_DOC]:
        if col not in df.columns:
            raise ValueError(f"Columna no encontrada en archivo {tipo}: '{col}'. ¿Es el archivo correcto?")

    df[COL_ESTADO] = df[COL_ESTADO].str.strip().str.upper()

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df[COL_FECHA] = pd.to_datetime(df[COL_FECHA], errors="coerce", format="mixed", dayfirst=True)

    COL_NTOMA = next(
        (c for c in df.columns if "TOMA DEL BENEFICIARIO" in c or "NUMERO DE TOMA" in c), None
    )
    df_ord = df.dropna(subset=[COL_FECHA]).copy()
    if COL_NTOMA:
        df_ord["_n"] = pd.to_numeric(df_ord[COL_NTOMA], errors="coerce").fillna(0)
        df_ord = df_ord.sort_values([COL_DOC, COL_FECHA, "_n"])
    else:
        df_ord = df_ord.sort_values([COL_DOC, COL_FECHA])
    df_ultimo = df_ord.drop_duplicates(subset=[COL_DOC], keep="last")

    ahora = pd.to_datetime(datetime.now())
    resultados = []

    for contrato in df[COL_CONTRATO].dropna().unique():
        df_c   = df[df[COL_CONTRATO] == contrato]
        df_ult = df_ultimo[df_ultimo[COL_CONTRATO] == contrato]

        hoja_unidades = (
            df_c.groupby(COL_UNIDAD)[COL_DOC]
            .nunique().reset_index()
            .rename(columns={COL_DOC: "TOTAL USUARIOS UNICOS"})
        )

        if act_bytes:
            df_act = leer_bytes(act_bytes, act_fn)
            tipos_u = [t.upper() for t in p["tipos_activos"]]
            df_act = df_act[df_act["Nombre Tipo de beneficiario"].str.strip().str.upper().isin(tipos_u)].copy()
            df_act = df_act[df_act["Número del Contrato"].str.strip() == str(contrato).strip()].copy()
            df_act["_doc"]      = df_act["Documento del beneficiario"].str.strip()
            df_act["_vinc"]     = pd.to_datetime(df_act["Fecha de vinculación del beneficiario"], errors="coerce", dayfirst=True)
            df_act["_meses"]    = (ahora.year - df_act["_vinc"].dt.year) * 12 + (ahora.month - df_act["_vinc"].dt.month)
            df_act["_esperadas"] = df_act["_meses"].apply(lambda m: tomas_esperadas(m, intervalo) if pd.notna(m) else 0)
            conteo = df_c.groupby(COL_DOC).size().reset_index(name="_reales")
            df_act = df_act.merge(conteo, left_on="_doc", right_on=COL_DOC, how="left")
            df_act["_reales"] = df_act["_reales"].fillna(0).astype(int)
            faltantes_src = df_act[df_act["_reales"] < df_act["_esperadas"]]

            filas = []
            for _, row in faltantes_src.iterrows():
                doc = row["_doc"]
                reg = df_ult[df_ult[COL_DOC] == doc]
                filas.append({
                    "ESTADO":             reg[COL_ESTADO].values[0] if not reg.empty else "",
                    "UNIDAD":             reg[COL_UNIDAD].values[0] if not reg.empty else row.get("Nombre de la unidad de servicio", ""),
                    "DOCUMENTO":          doc,
                    "NOMBRE":             reg[COL_NOMBRE].values[0] if not reg.empty and COL_NOMBRE in reg.columns else row.get("Primer Nombre del beneficiario", ""),
                    "APELLIDO":           reg[COL_APELLIDO].values[0] if not reg.empty and COL_APELLIDO in reg.columns else row.get("Primer apellido del beneficiario", ""),
                    "TOMAS REALIZADAS":   int(row["_reales"]),
                    "TOMAS ESPERADAS":    int(row["_esperadas"]),
                    "MESES VINCULADO":    int(row["_meses"]) if pd.notna(row["_meses"]) else 0,
                    "ULTIMO DIAGNOSTICO": reg[COL_DIAG].values[0] if not reg.empty and COL_DIAG in reg.columns else "SIN TOMA",
                    "MOTIVO":             "Sin toma registrada" if row["_reales"] == 0 else "Tomas insuficientes",
                })
            df_faltantes = pd.DataFrame(filas)
        else:
            ids_faltantes = set()
            for usuario, grupo in df_ord[df_ord[COL_CONTRATO] == contrato].groupby(COL_DOC):
                fechas = grupo[COL_FECHA].dt.to_period("M").unique()
                if len(fechas) > 1:
                    for i in range(1, len(fechas)):
                        diff = (fechas[i].year - fechas[i-1].year) * 12 + (fechas[i].month - fechas[i-1].month)
                        if diff > intervalo:
                            ids_faltantes.add(usuario)
                            break
                if usuario not in ids_faltantes:
                    ultima = fechas[-1]
                    diff_hoy = (ahora.year - ultima.year) * 12 + (ahora.month - ultima.month)
                    if diff_hoy >= intervalo:
                        ids_faltantes.add(usuario)

            cols_f = [c for c in [COL_ESTADO, COL_UNIDAD, COL_DOC, COL_NOMBRE, COL_APELLIDO, COL_DIAG] if c in df_ult.columns]
            df_faltantes = df_ult[df_ult[COL_DOC].isin(ids_faltantes)][cols_f].copy() if ids_faltantes else pd.DataFrame(columns=cols_f)

        if COL_DIAG in df_ult.columns:
            mask = df_ult[COL_DIAG].str.contains(p["patron_alerta"], case=False, na=False, regex=True)
            cols_a = [c for c in [COL_ESTADO, COL_UNIDAD, COL_DOC, COL_NOMBRE, COL_APELLIDO, COL_DIAG, COL_FECHA] if c in df_ult.columns]
            df_alerta = df_ult[mask][cols_a].copy()
            if COL_FECHA in df_alerta.columns:
                df_alerta[COL_FECHA] = df_alerta[COL_FECHA].dt.strftime("%Y-%m-%d")
        else:
            df_alerta = pd.DataFrame()

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            hoja_unidades.to_excel(writer, sheet_name="Usuarios por Unidad", index=False)
            (df_faltantes if not df_faltantes.empty
             else pd.DataFrame({"MENSAJE": ["Sin tomas faltantes registradas"]})).to_excel(
                writer, sheet_name="Tomas Faltantes", index=False)
            (df_alerta if not df_alerta.empty
             else pd.DataFrame({"MENSAJE": [p["msg_sin_alerta"]]})).to_excel(
                writer, sheet_name=p["hoja_alerta"], index=False)
        buf.seek(0)

        resultados.append({
            "contrato":  str(contrato),
            "tipo":      tipo,
            "vinculados": int(df_c[df_c[COL_ESTADO] == "VINCULADO"][COL_DOC].nunique()),
            "faltantes":  len(df_faltantes),
            "alertas":    len(df_alerta),
            "unidades":   len(hoja_unidades),
            "filename":   f"Informe_{tipo.upper()}_Contrato_{contrato}.xlsx",
            "bytes":      buf.read(),
        })

    return resultados


# ─────────────────────────────────────────────
# ENDPOINTS
# ─────────────────────────────────────────────

@app.get("/health")
def health():
    return {"status": "ok", "timestamp": datetime.now().isoformat()}


@app.post("/procesar")
async def procesar_archivos(
    modo: str = Form(...),
    general: UploadFile = File(...),
    gestante: UploadFile = File(...),
    activos: UploadFile = File(None),
    intervalo: int = Form(3),
):
    if modo not in ("basico", "completo"):
        raise HTTPException(400, "modo debe ser 'basico' o 'completo'")

    general_bytes  = await general.read()
    gestante_bytes = await gestante.read()
    act_bytes = (await activos.read()) if activos and activos.filename else None
    act_fn    = activos.filename if activos and activos.filename else None

    for nombre, data in [("general", general_bytes), ("gestante", gestante_bytes)]:
        if len(data) > MAX_UPLOAD_MB * 1024 * 1024:
            raise HTTPException(413, f"El archivo {nombre} supera los {MAX_UPLOAD_MB}MB permitidos.")
    if act_bytes and len(act_bytes) > MAX_UPLOAD_MB * 1024 * 1024:
        raise HTTPException(413, f"El archivo de activos supera los {MAX_UPLOAD_MB}MB permitidos.")

    if modo == "completo" and not act_bytes:
        raise HTTPException(400, "Modo completo requiere el archivo de activos.")

    try:
        todos = []
        for tipo, data, fn in [
            ("general",  general_bytes,  general.filename),
            ("gestante", gestante_bytes, gestante.filename),
        ]:
            todos.extend(procesar(data, fn, act_bytes, act_fn, tipo, intervalo))
    except ValueError as e:
        raise HTTPException(422, str(e))
    except Exception as e:
        raise HTTPException(500, f"Error procesando archivos: {str(e)}")

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in todos:
            zf.writestr(r["filename"], r["bytes"])
    zip_buf.seek(0)

    resumen = [{k: v for k, v in r.items() if k != "bytes"} for r in todos]

    return StreamingResponse(
        zip_buf,
        media_type="application/zip",
        headers={
            "Content-Disposition": "attachment; filename=informes_icbf.zip",
            "X-Resumen": json.dumps(resumen, ensure_ascii=False),
        },
    )


app.mount("/", StaticFiles(directory="/app/frontend", html=True), name="frontend")
