# services/adicionales_service.py
from io import BytesIO
import gc
from typing import List, Optional

import pandas as pd
import fitz  # PyMuPDF
import streamlit as st

# Utils
from services.adicionales_utils import (
    extrair_ruc,
    extrair_facturas,
    remover_ruc_indesejado,
    criar_coluna_proveedor_iscala,
    extrair_fecha_emision,
    normalizar_data,
    extrair_moneda,
    ajustar_e_padronizar_moneda,
    codificar_moneda,
    extrair_op_gravada,
    limpar_op_gravada,
    formatar_op_gravada,
    op_gravada_negativo_CN,
    extrair_tipo_doc,
    padronizar_tipo_doc,
    adicionar_cod_autorizacion_adicionales,
    adicionar_tip_doc_adicionales,
    organizar_colunas_adicionales,
    remover_duplicatas_source_file,
    adicionar_sharepoint_adicionales,
)

from services.duas_utils import adicionar_coluna_tasa


# ------------------------------------------------------------
# Leitura segura do PDF
# ------------------------------------------------------------
def _extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    try:
        with fitz.open(stream=BytesIO(pdf_bytes), filetype="pdf") as doc:
            return "".join(page.get_text() for page in doc)
    except Exception:
        return ""


# ------------------------------------------------------------
# FUNÇÃO PRINCIPAL (STREAMLIT)
# ------------------------------------------------------------
def process_adicionales_streamlit(
    uploaded_files: List,
    progress_widget=None,
    status_widget=None,
    cambio_df: Optional[pd.DataFrame] = None,
) -> pd.DataFrame:
    """
    ⚠️ IMPORTANTE:
    Esta função SEMPRE retorna um DataFrame (mesmo vazio).
    Isso evita UnboundLocalError no process_pdfs.py sem precisar alterá-lo.
    """

    # ✅ GARANTIA ABSOLUTA
    df_final = pd.DataFrame()

    if not uploaded_files:
        return df_final

    rows = []
    total = len(uploaded_files)

    try:
        for i, f in enumerate(uploaded_files, start=1):
            fname = getattr(f, "name", f"arquivo_{i}.pdf")
            texto = _extract_text_from_pdf_bytes(f.getvalue())

            rows.append(
                {
                    "source_file": fname,
                    "conteudo_pdf": texto,
                }
            )

            if progress_widget:
                pct = int(i / total * 100)
                progress_widget.progress(pct, text=f"Lendo {fname} ({i}/{total})")
            if status_widget:
                status_widget.write(f"📄 Lido: **{fname}**")

        if not rows:
            return df_final

        df = pd.DataFrame(rows)

        # -----------------------------
        # PIPELINE ADICIONALES
        # -----------------------------
        df["R.U.C"] = df["conteudo_pdf"].apply(extrair_ruc)
        df = remover_ruc_indesejado(df)
        df["Factura"] = df["conteudo_pdf"].apply(extrair_facturas)
        df = criar_coluna_proveedor_iscala(df)

        df["Fecha de Emisión"] = (
            df["conteudo_pdf"]
            .apply(extrair_fecha_emision)
            .apply(normalizar_data)
        )

        df["Moneda"] = (
            df["conteudo_pdf"]
            .apply(extrair_moneda)
            .apply(ajustar_e_padronizar_moneda)
        )

        df["Cod. Moneda"] = df["Moneda"].apply(codificar_moneda)

        df["Tipo Doc"] = df.apply(extrair_tipo_doc, axis=1)
        df = padronizar_tipo_doc(df)

        df["Op. Gravada"] = df.apply(extrair_op_gravada, axis=1)
        df["Op. Gravada"] = df["Op. Gravada"].apply(limpar_op_gravada)
        df["Op. Gravada"] = df["Op. Gravada"].apply(formatar_op_gravada)
        df = op_gravada_negativo_CN(df)

        # -----------------------------
        # TASA (reaproveita DUAS)
        # -----------------------------
        df = adicionar_coluna_tasa(df, cambio_df=cambio_df)
        if "Cod. Moneda" in df.columns:
            df.loc[df["Cod. Moneda"] == "00", "Tasa"] = 1

        # -----------------------------
        # SHAREPOINT (opcional)
        # -----------------------------
        sharepoint_df = st.session_state.get("sharepoint_df")
        df = adicionar_sharepoint_adicionales(df, sharepoint_df)

        # -----------------------------
        # CÓDIGOS SUNAT
        # -----------------------------
        df = adicionar_cod_autorizacion_adicionales(df)
        df = adicionar_tip_doc_adicionales(df)

        # -----------------------------
        # ORGANIZAÇÃO FINAL
        # -----------------------------
        df = organizar_colunas_adicionales(df)
        df = remover_duplicatas_source_file(df)

        df_final = df.copy()

    except Exception as e:
        if status_widget:
            status_widget.error("❌ Erro no processamento de Gastos Adicionales.")
            status_widget.exception(e)

    finally:
        gc.collect()

    # ✅ SEMPRE retorna um DataFrame
    return df_final
