# services/externos_service.py
from io import BytesIO
import pandas as pd
import fitz  # PyMuPDF
from typing import List, Optional
import streamlit as st

# Import das funções auxiliares
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

    # ------------------------------
    # Leitura dos PDFs para DataFrame
    # ------------------------------
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

    # ========== PIPELINE PRINCIPAL ==========
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

    # Tasa (opcional)
    df = adicionar_coluna_tasa_externos(df, cambio_df=cambio_df)
    if "Cod. Moneda" in df.columns:
        df.loc[df["Cod. Moneda"] == "00", "Tasa"] = 1

    # Códigos
    df = adicionar_cod_autorizacion_ext(df)
    df = adicionar_tip_fac_ext(df)

    # =============================
    # Merge PEC / dados do SharePoint
    # =============================
    from services.externos_utils import adicionar_pec_sharepoint

    sharepoint_df = st.session_state.get("sharepoint_df")

    # defensivo: a função no utils pode retornar tuple por engano; normalizamos para DF
    df_sp = adicionar_pec_sharepoint(df, sharepoint_df)
    if isinstance(df_sp, tuple):
        df = df_sp[0]
    else:
        df = df_sp

    # =============================
    # Heurística "Lineaabajo"
    # =============================
    MAP_LINEA = {
        # --- Itens canônicos principais ---
        "REFRIGERATOR": 36,
        "CHEST FREEZER": 35,
        "FREEZER": 35,
        "STOVE": 38,
        "COOKER": 22,            # categoria genérica de cozinha
        "OVEN": 38,              # genérico; temos também variantes específicas abaixo
        "ELECTRIC OVEN": 22,
        "GAS OVEN": 38,
        "MICROWAVE OVEN": 22,    # mantendo seu valor original (lista tinha 38)
        "WASHING MACHINE": 25,
        "WASHER": 25,
        "DRYER": 45,             # equivalente a SECADORA
        "SECADORA": 45,
        "VACUUM CLEANER": 10,
        "ROBOTIC VACUUM CLEANERS": 10,
        "STEAM IRON": 34,
        "GARMENT STEAMER": 24,   # >>> adicionado conforme pedido (24)
        "HANDHELD GARMENT STEAMER": 34,
        "DISHWASHER": 24,
        "GAS HOB": 38,           # mantendo seu valor original (lista tinha 22)
        "COOKER HOOD": 22,
        "COCINA": 22,
        "AIR FRYER": 22,
        "SPLIT AIR CONDITIONER": 41,
        "AIR CONDITIONER": 41,   # >>> adicionado conforme pedido (41)
        "WATER DISPENSER": 22,
        "WINE COOLER": 22,
        "RICE COOKER": 22,
        "JUICER": 22,
        "BLENDER": 34,
        "KETTLE": 34,
        "ELECTRICAL COOKING": 22,
        "SPARE PARTS": 34,
    
        # --- Itens que você já usava e fazem sentido manter ---
        "SEC ELEC": 25,
        "KE4CT": 22,
    }

    def detectar_linea(texto_pdf: str) -> Optional[int]:
        if not isinstance(texto_pdf, str):
            return None
        up = texto_pdf.upper()
        for ref, num in MAP_LINEA.items():
            if ref in up:
                return num
        return None

    # Criar antes de organizar as colunas
    df["Lineaabajo"] = df["conteudo_pdf"].apply(detectar_linea)

    # Organiza e remove duplicatas (uma única vez)
    df = organizar_colunas_externos(df)
    df = remover_duplicatas_source_file(df)

    # Agora sim, remover o conteúdo bruto do PDF
    df = df.drop(columns=["conteudo_pdf"], errors="ignore")

    # Mensagens finais
    if progress_widget:
        progress_widget.progress(100, text="Concluído (Externos).")
    if status_widget:
        status_widget.success("Pipeline Externos finalizado.")
    
    return df
