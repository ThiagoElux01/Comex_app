# services/externos_service.py
from io import BytesIO
import pandas as pd
import fitz  # PyMuPDF
from typing import List, Optional
import streamlit as st

# Import correto das funções auxiliares
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


def _extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    """Extrai texto de todas as páginas do PDF via PyMuPDF."""
    try:
        with fitz.open(stream=BytesIO(pdf_bytes), filetype="pdf") as doc:
            text = "".join(page.get_text() for page in doc)
            return text if text.strip() else "[PDF baseado em imagem - sem texto extraível]"
    except Exception:
        return "[Erro ao abrir/ler o PDF]"


def process_externos_streamlit(
    uploaded_files: List,
    progress_widget=None,
    status_widget=None,
    cambio_df: Optional[pd.DataFrame] = None,
):
    if not uploaded_files:
        return None

    rows = []
    total = len(uploaded_files)

    # -------------------------------
    # Leitura dos PDFs para DataFrame
    # -------------------------------
    for i, f in enumerate(uploaded_files, start=1):
        fname = getattr(f, "name", f"arquivo_{i}.pdf")
        text = _extract_text_from_pdf_bytes(f.getvalue())
        rows.append({"source_file": fname, "conteudo_pdf": text})

        if progress_widget:
            pct = int(i / total * 100)
            progress_widget.progress(
                pct, text=f"Lendo {fname} ({i}/{total})"
            )
            progress_widget.progress(pct, text=f"Lendo {fname} ({i}/{total})")

        if status_widget:
            status_widget.write(f"📄 Lido: **{fname}**")

    df = pd.DataFrame(rows)

    # ========== PIPELINE PRINCIPAL ==========
    # ------------------------
    # PIPELINE PRINCIPAL
    # ------------------------

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

    df = adicionar_coluna_tasa_externos(df, cambio_df=cambio_df)

    if "Cod. Moneda" in df.columns:
        df.loc[df["Cod. Moneda"] == "00", "Tasa"] = 1

    df = adicionar_cod_autorizacion_ext(df)
    df = adicionar_tip_fac_ext(df)

    df = organizar_colunas_externos(df)
    df = remover_duplicatas_source_file(df)

    # ===========================================================
    # ADICIONA PEC VINDO DO SHAREPOINT (SE EXISTIR)
    # ===========================================================
    from services.externos_utils import adicionar_pec_sharepoint
    sharepoint_df = st.session_state.get("sharepoint_df")
    df = adicionar_pec_sharepoint(df, sharepoint_df)
    MAP_LINEA = {
        "REFRIGERATOR": 36,
        "CHEST FREEZER": 35,
        "STOVE": 38,
        "WASHER": 25,
        "WASHING MACHINE": 25,
        "MICROWAVE OVEN": 22,
        "COCINA": 22,
        "VACUUM CLEANER": 10,
        "AIR FRYER": 22,
        "SECADORA": 45,
        "GAS OVEN": 38,
        "GAS HOB": 38,
        "DISHWASHER": 24,
        "SPARE PARTS": 34,
        "COOKER": 22,
        "ELECTRIC OVEN": 22,
    }

    
    from typing import Optional  # (já existe no arquivo)
    def detectar_linea(texto_pdf: str) -> Optional[int]:
        if not isinstance(texto_pdf, str):
            return None
        up = texto_pdf.upper()
        for ref, num in MAP_LINEA.items():
            if ref in up:
                return num
        return None

    # 👇 CRIAR ANTES DE ORGANIZAR AS COLUNAS
    df["Lineaabajo"] = df["conteudo_pdf"].apply(detectar_linea)

    # --------------------------------------------------------
    # CRIAÇÃO DA COLUNA Lineaabajo COM BASE NO CONTEÚDO DO PDF
    # --------------------------------------------------------

    # Agora sim remover o conteúdo do PDF
    # Agora organizar colunas
    df = organizar_colunas_externos(df)
    df = remover_duplicatas_source_file(df)

    # Agora sim apagar o conteúdo bruto do PDF
    df = df.drop(columns=["conteudo_pdf"], errors="ignore")

    # --------------------------------------------------------
    # COMPLEMENTAR CAMPOS VAZIOS COM SHAREPOINT
    # --------------------------------------------------------
    def preencher_vazio(dest_col, src_col):
        if dest_col in df.columns and src_col in df.columns:
            df[dest_col] = df[dest_col].fillna("").replace("", None)
            df[src_col] = df[src_col].fillna("").replace("", None)
            df[dest_col] = df[dest_col].combine_first(df[src_col])

    preencher_vazio("Proveedor Iscala", "proveedor")
    preencher_vazio("Factura", "numero_de_documento")
    preencher_vazio("Tipo Doc", "tipo_doc")
    preencher_vazio("Fecha de Emisión", "Fecha_Emision")
    preencher_vazio("Moneda", "moneda")
    preencher_vazio("Amount", "importe_documento")
    preencher_vazio("Tasa", "Tasa_Sharepoint")

    # Mensagens finais
    # --------------------------------------------------------
    # FIM DO PIPELINE
    # --------------------------------------------------------
    if progress_widget:
        progress_widget.progress(100, text="Concluído (Externos).")

    if status_widget:
        status_widget.success("Pipeline Externos finalizado.")

    return df
