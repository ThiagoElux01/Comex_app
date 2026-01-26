
# services/adicionales_service.py
from io import BytesIO
from typing import List, Optional
import fitz  # PyMuPDF
import pandas as pd

from services.adicionales_utils import (
    extrair_ruc, extrair_facturas, remover_ruc_indesejado, criar_coluna_proveedor_iscala,
    extrair_fecha_emision, normalizar_data, extrair_moneda, ajustar_e_padronizar_moneda,
    codificar_moneda, extrair_op_gravada, limpar_op_gravada, formatar_op_gravada,
    atribuir_cuenta, error, adicionar_coluna_tasa, organizar_colunas_adicionales,
    extrair_tipo_doc, padronizar_tipo_doc, Ajustar_nro_nota_credito,
    adicionar_cod_autorizacion_adicionales, adicionar_tip_doc_adicionales,
    remover_duplicatas_source_file, op_gravada_negativo_CN
)

def _extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    try:
        with fitz.open(stream=BytesIO(pdf_bytes), filetype="pdf") as doc:
            text = "".join(page.get_text() for page in doc)
            return text if text.strip() else "[PDF baseado em imagem - sem texto extraível]"
    except Exception:
        return "[Erro ao abrir/ler o PDF]"

def process_adicionales_streamlit(
    uploaded_files: List,
    progress_widget=None,
    status_widget=None,
    cambio_df: Optional[pd.DataFrame] = None,
) -> Optional[pd.DataFrame]:
    if not uploaded_files:
        return None

    rows = []
    total = len(uploaded_files)
    for i, f in enumerate(uploaded_files, start=1):
        fname = getattr(f, "name", f"arquivo_{i}.pdf")
        text = _extract_text_from_pdf_bytes(f.getvalue())
        rows.append({"source_file": fname, "conteudo_pdf": text})

        if progress_widget:
            pct = int(i / total * 100)
            progress_widget.progress(pct, text=f"Lendo {fname} ({i}/{total})")
        if status_widget:
            status_widget.write(f"📄 Lido: **{fname}**")

    df = pd.DataFrame(rows)

    # --- Pipeline ---
    df["R.U.C"] = df["conteudo_pdf"].apply(extrair_ruc)
    df = remover_ruc_indesejado(df)
    df = criar_coluna_proveedor_iscala(df)

    df["Factura"] = df["conteudo_pdf"].apply(extrair_facturas)
    df["Fecha de Emisión"] = df["conteudo_pdf"].apply(extrair_fecha_emision).apply(normalizar_data)

    df["Moneda"] = df["conteudo_pdf"].apply(extrair_moneda).apply(ajustar_e_padronizar_moneda)
    df["Cod. Moneda"] = df["Moneda"].apply(codificar_moneda)

    df["Tipo Doc"] = df.apply(extrair_tipo_doc, axis=1)
    df = padronizar_tipo_doc(df)

    df["Op. Gravada"] = df.apply(extrair_op_gravada, axis=1).apply(limpar_op_gravada).apply(formatar_op_gravada)
    df = op_gravada_negativo_CN(df)

    df["Cuenta"] = df["Cod. Moneda"].apply(atribuir_cuenta)
    df["Error"] = df["Proveedor Iscala"].apply(error)

    # Tasa (merge por data)
    df = adicionar_coluna_tasa(df, cambio_df=cambio_df)
    if "Cod. Moneda" in df.columns:
        df.loc[df["Cod. Moneda"] == "00", "Tasa"] = 1

    df = Ajustar_nro_nota_credito(df)
    df = adicionar_cod_autorizacion_adicionales(df)
    df = adicionar_tip_doc_adicionales(df)

    df = organizar_colunas_adicionales(df)
    df = remover_duplicatas_source_file(df)

    if progress_widget:
        progress_widget.progress(100, text="Concluído (Gastos Adicionales).")
    if status_widget:
        status_widget.success("Pipeline Adicionales finalizado.")

    return df
