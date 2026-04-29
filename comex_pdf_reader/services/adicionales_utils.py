from io import BytesIO
import gc
import pandas as pd
import fitz
import streamlit as st

from services.adicionales_utils import *

def _extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    try:
        with fitz.open(stream=BytesIO(pdf_bytes), filetype="pdf") as doc:
            return "".join(page.get_text() for page in doc)
    except Exception:
        return ""

def process_adicionales_streamlit(
    uploaded_files,
    progress_widget=None,
    status_widget=None,
    cambio_df=None,
):
    if not uploaded_files:
        return None

    BATCH_SIZE = 10   # ✅ igual ao Externos
    total = len(uploaded_files)
    dfs_resultado = []

    for start in range(0, total, BATCH_SIZE):
        batch = uploaded_files[start:start + BATCH_SIZE]
        rows = []

        for i, f in enumerate(batch, start=start + 1):
            fname = getattr(f, "name", f"arquivo_{i}.pdf")
            texto = _extract_text_from_pdf_bytes(f.getvalue())
            rows.append({
                "source_file": fname,
                "conteudo_pdf": texto
            })

            if progress_widget:
                pct = int(i / total * 100)
                progress_widget.progress(pct, text=f"Lendo {fname} ({i}/{total})")
            if status_widget:
                status_widget.write(f"📄 Lido: **{fname}**")

        df = pd.DataFrame(rows)

        # -------- PIPELINE ADICIONALES --------
        df["R.U.C"] = df["conteudo_pdf"].apply(extrair_ruc)
        df["Factura"] = df["conteudo_pdf"].apply(extrair_facturas)

        df = criar_coluna_proveedor_iscala(df)

        df["Fecha de Emisión"] = df["conteudo_pdf"].apply(extrair_fecha_emision)
        df["Fecha de Emisión"] = df["Fecha de Emisión"].apply(normalizar_data)

        df["Moneda"] = df["conteudo_pdf"].apply(extrair_moneda)
        df["Moneda"] = df["Moneda"].apply(ajustar_e_padronizar_moneda)
        df["Cod. Moneda"] = df["Moneda"].apply(codificar_moneda)
        df["Cuenta"] = df["Cod. Moneda"].apply(atribuir_cuenta)

        df["Tipo Doc"] = df.apply(extrair_tipo_doc, axis=1)
        df = padronizar_tipo_doc(df)

        df["Op. Gravada"] = df.apply(extrair_op_gravada, axis=1)
        df["Op. Gravada"] = df["Op. Gravada"].apply(limpar_op_gravada)
        df["Op. Gravada"] = df["Op. Gravada"].apply(formatar_op_gravada)
        df = op_gravada_negativo_CN(df)

        df = Ajustar_nro_nota_credito(df)

        # SharePoint
        sharepoint_df = st.session_state.get("sharepoint_df")
        df = adicionar_sharepoint_adicionales(df, sharepoint_df)

        # Códigos
        df = adicionar_cod_autorizacion_adicionales(df)
        df = adicionar_tip_doc_adicionales(df)

        df["Error"] = df["Factura"].apply(error)

        # Organiza
        df = organizar_colunas_adicionales(df)
        df = remover_duplicatas_source_file(df)

        # ✅ remove texto bruto cedo
        df = df.drop(columns=["conteudo_pdf"], errors="ignore")

        dfs_resultado.append(df)

        del rows, df
        gc.collect()

    df_final = pd.concat(dfs_resultado, ignore_index=True)

    if progress_widget:
        progress_widget.progress(100, text="Concluído (Adicionales).")
    if status_widget:
        status_widget.success("Pipeline Adicionales finalizado.")

    return df_final
