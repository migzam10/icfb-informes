"""
ICBF Informes - Backend FastAPI
Modos:
  - basico:    solo General + Gestante (sin archivo de activos)
  - completo:  General + Gestante + BeneficiariosPIActivos (análisis de déficit)
"""

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

import pandas as pd
import warnings
import math
import os
import uuid
import shutil
from datetime import datetime
from pathlib import Path

app = FastAPI(title="ICBF Informes")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOADS = Path("/tmp/icbf_uploads")
OUTPUTS = Path("/tmp/icbf_outputs")
UPLOADS.mkdir(exist_ok=True)
OUTPUTS.mkdir(exist_ok=True)

# ─────────────────────────────────────────────
# CONFIGURACIÓN POR TIPO
# ─────────────────────────────────────────────
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

def norm(col: str) -> str:
    return col.strip().upper().translate(REEMPLAZOS).replace("  ", " ").strip()

def leer_excel(ruta: Path, preferir_hoja: str = None) -> pd.DataFrame:
    hojas = pd.ExcelFile(ruta).sheet_names
    hoja = hojas[0]
    if preferir_hoja:
        for h in hojas:
            if preferir_hoja.lower() in h.lower():
                hoja = h
                break
    return pd.read_excel(ruta, sheet_name=hoja, dtype=str)

def tomas_esperadas(meses: float, intervalo: int) -> int:
    return max(1, math.floor(meses / intervalo))

# ─────────────────────────────────────────────
# LÓGICA CENTRAL
# ─────────────────────────────────────────────
def procesar(ruta_nut: Path, tipo: str, intervalo: int,
             ruta_activos: Path = None) -> dict:
    p = PERFIL[tipo]
    df = leer_excel(ruta_nut, preferir_hoja=p["hoja_excel"])
    df.columns = [norm(c) for c in df.columns]

    COL_DOC    = norm(p["col_doc"])
    COL_DIAG   = p["col_diag"]
    COL_FECHA  = p["col_fecha"]
    COL_ESTADO = "ESTADO"
    COL_CONTRATO = "NUMERO CONTRATO"
    COL_UNIDAD   = "NOMBRE UNIDAD"
    COL_NOMBRE   = "PRIMER NOMBRE BENEFICIARIO"
    COL_APELLIDO = "PRIMER APELLIDO BENEFICIARIO"

    for col in [COL_FECHA, COL_ESTADO, COL_CONTRATO, COL_DOC]:
        if col not in df.columns:
            raise ValueError(f"Columna no encontrada: '{col}'. ¿Es el archivo correcto?")

    df[COL_ESTADO] = df[COL_ESTADO].str.strip().str.upper()
    df = df[df[COL_ESTADO] == "VINCULADO"].copy()
    total_vinculados = len(df)

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df[COL_FECHA] = pd.to_datetime(df[COL_FECHA], errors="coerce", format="mixed", dayfirst=True)

    # Último registro por usuario (desempate por nro de toma)
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

    resultados = []
    contratos = df[COL_CONTRATO].dropna().unique()
    ahora = pd.to_datetime(datetime.now())

    for contrato in contratos:
        df_c    = df[df[COL_CONTRATO] == contrato]
        df_ult  = df_ultimo[df_ultimo[COL_CONTRATO] == contrato]

        # Hoja 1: usuarios por unidad
        hoja_unidades = (
            df_c.groupby(COL_UNIDAD)[COL_DOC]
            .nunique().reset_index()
            .rename(columns={COL_DOC: "TOTAL USUARIOS UNICOS"})
        )

        # Hoja 2: tomas faltantes
        if ruta_activos:
            df_act = leer_excel(ruta_activos)
            tipos_u = [t.upper() for t in p["tipos_activos"]]
            df_act = df_act[df_act["Nombre Tipo de beneficiario"].str.strip().str.upper().isin(tipos_u)].copy()
            df_act = df_act[df_act["Número del Contrato"].str.strip() == str(contrato).strip()].copy()
            df_act["_doc"]  = df_act["Documento del beneficiario"].str.strip()
            df_act["_vinc"] = pd.to_datetime(df_act["Fecha de vinculación del beneficiario"], errors="coerce", dayfirst=True)
            df_act["_meses"] = (ahora.year - df_act["_vinc"].dt.year) * 12 + (ahora.month - df_act["_vinc"].dt.month)
            df_act["_esperadas"] = df_act["_meses"].apply(lambda m: tomas_esperadas(m, intervalo) if pd.notna(m) else 0)
            conteo = df_c.groupby(COL_DOC).size().reset_index(name="_reales")
            df_act = df_act.merge(conteo, left_on="_doc", right_on=COL_DOC, how="left")
            df_act["_reales"] = df_act["_reales"].fillna(0).astype(int)
            faltantes = df_act[(df_act["_reales"] < df_act["_esperadas"])].copy()

            filas = []
            for _, row in faltantes.iterrows():
                doc = row["_doc"]
                reg = df_ult[df_ult[COL_DOC] == doc]
                filas.append({
                    "UNIDAD":           reg[COL_UNIDAD].values[0] if not reg.empty else row.get("Nombre de la unidad de servicio", ""),
                    "DOCUMENTO":        doc,
                    "NOMBRE":           reg[COL_NOMBRE].values[0] if not reg.empty and COL_NOMBRE in reg.columns else row.get("Primer Nombre del beneficiario", ""),
                    "APELLIDO":         reg[COL_APELLIDO].values[0] if not reg.empty and COL_APELLIDO in reg.columns else row.get("Primer apellido del beneficiario", ""),
                    "TOMAS REALIZADAS": int(row["_reales"]),
                    "TOMAS ESPERADAS":  int(row["_esperadas"]),
                    "MESES VINCULADO":  int(row["_meses"]) if pd.notna(row["_meses"]) else 0,
                    "ULTIMO DIAGNOSTICO": reg[COL_DIAG].values[0] if not reg.empty and COL_DIAG in reg.columns else "SIN TOMA",
                    "MOTIVO":           "Sin toma registrada" if row["_reales"] == 0 else "Tomas insuficientes",
                })
            df_faltantes = pd.DataFrame(filas)
        else:
            # Modo básico: lógica por huecos históricos
            ids_faltantes = set()
            for usuario, grupo in df_ord[df_ord[COL_CONTRATO] == contrato].groupby(COL_DOC):
                fechas = grupo[COL_FECHA].dt.to_period("M").unique()
                if len(fechas) > 1:
                    for i in range(1, len(fechas)):
                        diff = (fechas[i].year - fechas[i-1].year)*12 + (fechas[i].month - fechas[i-1].month)
                        if diff > intervalo:
                            ids_faltantes.add(usuario)
                            break
                if usuario not in ids_faltantes:
                    ultima = fechas[-1]
                    diff_hoy = (ahora.year - ultima.year)*12 + (ahora.month - ultima.month)
                    if diff_hoy >= intervalo:
                        ids_faltantes.add(usuario)

            cols_f = [COL_UNIDAD, COL_DOC, COL_NOMBRE, COL_APELLIDO, COL_DIAG]
            cols_f = [c for c in cols_f if c in df_ult.columns]
            df_faltantes = df_ult[df_ult[COL_DOC].isin(ids_faltantes)][cols_f].copy() if ids_faltantes else pd.DataFrame(columns=cols_f)

        # Hoja 3: alertas nutricionales
        if COL_DIAG in df_ult.columns:
            mask = df_ult[COL_DIAG].str.contains(p["patron_alerta"], case=False, na=False, regex=True)
            df_alerta = df_ult[mask][[COL_UNIDAD, COL_DOC, COL_NOMBRE, COL_APELLIDO, COL_DIAG, COL_FECHA]].copy()
            df_alerta[COL_FECHA] = df_alerta[COL_FECHA].dt.strftime("%Y-%m-%d")
        else:
            df_alerta = pd.DataFrame()

        resultados.append({
            "contrato":     contrato,
            "df_unidades":  hoja_unidades,
            "df_faltantes": df_faltantes,
            "df_alerta":    df_alerta,
            "stats": {
                "vinculados":  total_vinculados,
                "faltantes":   len(df_faltantes),
                "alertas":     len(df_alerta),
                "unidades":    len(hoja_unidades),
            }
        })

    return resultados


def exportar_excel(resultados: list, tipo: str, carpeta: Path) -> list:
    p = PERFIL[tipo]
    archivos = []
    for r in resultados:
        nombre = carpeta / f"Informe_{tipo.upper()}_Contrato_{r['contrato']}.xlsx"
        with pd.ExcelWriter(nombre, engine="openpyxl") as writer:
            r["df_unidades"].to_excel(writer, sheet_name="Usuarios por Unidad", index=False)
            if not r["df_faltantes"].empty:
                r["df_faltantes"].to_excel(writer, sheet_name="Tomas Faltantes", index=False)
            else:
                pd.DataFrame({"MENSAJE": ["Sin tomas faltantes registradas"]}).to_excel(
                    writer, sheet_name="Tomas Faltantes", index=False)
            if not r["df_alerta"].empty:
                r["df_alerta"].to_excel(writer, sheet_name=p["hoja_alerta"], index=False)
            else:
                pd.DataFrame({"MENSAJE": [p["msg_sin_alerta"]]}).to_excel(
                    writer, sheet_name=p["hoja_alerta"], index=False)
        archivos.append(str(nombre))
    return archivos


# ─────────────────────────────────────────────
# ENDPOINTS
# ─────────────────────────────────────────────
@app.get("/health")
def health():
    return {"status": "ok", "timestamp": datetime.now().isoformat()}


@app.post("/procesar")
async def procesar_archivos(
    modo: str = Form(...),                          # "basico" | "completo"
    general: UploadFile = File(...),
    gestante: UploadFile = File(...),
    activos: UploadFile = File(None),
    intervalo: int = Form(3),
):
    job_id = uuid.uuid4().hex[:8]
    carpeta = OUTPUTS / job_id
    carpeta.mkdir(parents=True)
    tmp = UPLOADS / job_id
    tmp.mkdir(parents=True)

    async def guardar(f: UploadFile, nombre: str) -> Path:
        ruta = tmp / nombre
        with open(ruta, "wb") as out:
            out.write(await f.read())
        return ruta

    try:
        ruta_general  = await guardar(general,  "general.xlsx")
        ruta_gestante = await guardar(gestante, "gestante.xlsx")
        ruta_activos  = await guardar(activos,  "activos.xlsx") if activos and activos.filename else None

        if modo == "completo" and not ruta_activos:
            raise HTTPException(400, "Modo completo requiere el archivo de activos.")

        archivos_generados = []
        resumen = []

        for tipo, ruta_nut in [("general", ruta_general), ("gestante", ruta_gestante)]:
            resultados = procesar(
                ruta_nut, tipo, intervalo,
                ruta_activos if modo == "completo" else None
            )
            archivos = exportar_excel(resultados, tipo, carpeta)
            archivos_generados.extend(archivos)
            for r in resultados:
                resumen.append({"tipo": tipo, "contrato": str(r["contrato"]), **r["stats"]})

        # Empaquetar en ZIP si hay más de un archivo
        if len(archivos_generados) > 1:
            import zipfile
            zip_path = OUTPUTS / f"Informes_{job_id}.zip"
            with zipfile.ZipFile(zip_path, "w") as zf:
                for f in archivos_generados:
                    zf.write(f, Path(f).name)
            descarga_url = f"/descargar/{job_id}/zip"
        else:
            descarga_url = f"/descargar/{job_id}/{Path(archivos_generados[0]).name}"

        return JSONResponse({
            "ok": True,
            "job_id": job_id,
            "modo": modo,
            "resumen": resumen,
            "descarga": descarga_url,
            "archivos": [Path(f).name for f in archivos_generados],
        })

    except ValueError as e:
        raise HTTPException(422, str(e))
    except Exception as e:
        raise HTTPException(500, f"Error procesando archivos: {str(e)}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


@app.get("/descargar/{job_id}/{nombre}")
def descargar(job_id: str, nombre: str):
    if nombre == "zip":
        ruta = OUTPUTS / f"Informes_{job_id}.zip"
        return FileResponse(ruta, filename=f"Informes_{job_id}.zip",
                            media_type="application/zip")
    ruta = OUTPUTS / job_id / nombre
    if not ruta.exists():
        raise HTTPException(404, "Archivo no encontrado")
    return FileResponse(ruta, filename=nombre,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# Servir frontend estático
app.mount("/", StaticFiles(directory="/app/frontend", html=True), name="frontend")
