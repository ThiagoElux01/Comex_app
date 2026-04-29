from io import BytesIO
import gc
import pandas as pd
import fitz  # PyMuPDF
from typing import List, Optional
import streamlit as st

from services.externos_utils import (
    identificar_Proveedor,
    adicionar_provedor_iscala,
    extrair_factura,
    ajustar_factura,
    extrair_fecha,
    ajustar_coluna_fecha,
    adicionar_colunas_fixas,
    adicionar_tipo_doc,
    adicionar_amount,
    ajustar_amount,
    adicionar_erro,
    organizar_colunas_externos,
    adicionar_cod_autorizacion_ext,
    adicionar_tip_fac_ext,
    remover_duplicatas_source_file,
    op_gravada_negativo_CN_externos,
)


# ==============================
# Utils
# ==============================
def adicionar_coluna_tasa_externos(df, cambio_df):
    if cambio_df is None or cambio_df.empty or "Fecha de Emisión" not in df.columns:
        return df
    dft = df.copy()
    dft["Fecha_tmp"] = pd.to_datetime(
        dft["Fecha de Emisión"], errors="coerce", dayfirst=True
    )
    tasa = cambio_df.copy()
    tasa["Data"] = pd.to_datetime(tasa["Data"], errors="coerce", dayfirst=True)
    dft = dft.merge(
        tasa[["Data", "Venta"]],
        how="left",
        left_on="Fecha_tmp",
        right_on="Data",
    )
    dft.rename(columns={"Venta": "Tasa"}, inplace=True)
    dft.drop(columns=["Fecha_tmp", "Data"], inplace=True)
    return dft


def _extract_text(pdf_bytes: bytes) -> str:
    try:
        with fitz.open(stream=BytesIO(pdf_bytes), filetype="pdf") as doc:
            return "".join(page.get_text() for page in doc)
    except Exception:
        return ""


# ==============================
# MAIN
# ==============================
def process_externos_streamlit(
    uploaded_files: List,
    progress_widget=None,
    status_widget=None,
    cambio_df: Optional[pd.DataFrame] = None,
):
    if not uploaded_files:
        return None

    BATCH_SIZE = 5  # ✅ seguro para Streamlit Cloud
    total = len(uploaded_files)

    dfs_finais = []

    for start in range(0, total, BATCH_SIZE):
        batch = uploaded_files[start:start + BATCH_SIZE]
        rows = []

        for idx, f in enumerate(batch, start=start + 1):
            pdf_bytes = f.getvalue()
            texto = _extract_text(pdf_bytes)

            rows.append({
                "source_file": f.name,
                "conteudo_pdf": texto
            })

            del pdf_bytes
            del texto

            if progress_widget:
                progress_widget.progress(
                    int(idx / total * 100),
                    text=f"Lendo {f.name} ({idx}/{total})"
                )
            if status_widget:
                status_widget.write(f"📄 {f.name}")

        df = pd.DataFrame(rows)

        # ================= PIPELINE =================
        df = identificar_Proveedor(df)
        df = adicionar_provedor_iscala(df)
        df = extrair_factura(df)
        df = ajustar_factura(df)
        df = extrair_fecha(df)
        df = ajustar_coluna_fecha(df)
        df = adicionar_colunas_fixas(df)
        df = adicionar_tipo_doc(df)
        df = adicionar_amount(df)
        df = ajustar_amount(df)
        df = op_gravada_negativo_CN_externos(df)
        df = adicionar_erro(df)

        df = adicionar_coluna_tasa_externos(df, cambio_df)
        df = adicionar_cod_autorizacion_ext(df)
        df = adicionar_tip_fac_ext(df)

        # Heurística Lineaabajo (ANTES de dropar texto)
        MAP_LINEA = {
            "REFRIGERATOR": 36,
            "FREEZER": 35,
            "STOVE": 38,
            "AIR CONDITIONER": 41,
            "WASHING MACHINE": 25,
        }

        def detectar_linea(txt):
            if not isinstance(txt, str):
                return None
            txt = txt.upper()
            for k, v in MAP_LINEA.items():
                if k in txt:
                    return v
            return None

        df["Lineaabajo"] = df["conteudo_pdf"].apply(detectar_linea)

        # 🔥 CRUCIAL: descartar texto IMEDIATAMENTE
        df.drop(columns=["conteudo_pdf"], inplace=True)

        df = organizar_colunas_externos(df)
        df = remover_duplicatas_source_file(df)

        dfs_finais.append(df)

        # 🔥 LIMPEZA PESADA
        del rows
        del df
        gc.collect()

    df_final = pd.concat(dfs_finais, ignore_index=True)
    del dfs_finais
    gc.collect()

    if progress_widget:
        progress_widget.progress(100, text="Concluído (Externos).")
    if status_widget:
        status_widget.success("Pipeline Externos finalizado.")

    return df_final
