
# services/externos_service.py
from io import BytesIO
import pandas as pd
import fitz  # PyMuPDF
from typing import List, Optional

# Suas funções já existentes (cole-as em utils/dataframe_utils.py, ou ajuste estes imports)
from utils.dataframe_utils import (
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
    adicionar_tip_fAC_ext,  # cuidado com o nome: verifique a grafia exata no seu módulo
    remover_duplicatas_source_file,
    op_gravada_negativo_CN_externos,
)


def adicionar_coluna_tasa_externos(df, cambio_df):
    if cambio_df is None or cambio_df.empty or "Fecha de Emisión" not in df.columns:
        return df
    dft = df.copy()
    dft["Fecha_tmp"] = pd.to_datetime(dft["Fecha de Emisión"], errors="coerce", dayfirst=True)
    tasa = cambio_df.copy()
    tasa["Data"] = pd.to_datetime(tasa["Data"], errors="coerce", dayfirst=True)
    dft = dft.merge(tasa[["Data", "Venta"]], how="left", left_on="Fecha_tmp", right_on="Data")
    dft.rename(columns={"Venta": "Tasa"}, inplace=True)
    dft.drop(columns=["Fecha_tmp","Data"], inplace=True)
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
    cambio_df: Optional[pd.DataFrame] = None,   # Tasa consolidada (opcional)
):
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

    # === Pipeline (as suas funções) ===
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

    # Negativar Amount para Credit Note
    df = op_gravada_negativo_CN_externos(df)

    # Erro se não achou fornecedor
    df = adicionar_erro(df)

    # Se você já tiver um util que injeta Tasa a partir de "Fecha de Emisión", chame aqui.
    # Ex.: df = adicionar_coluna_tasa(df, cambio_df=cambio_df)

    # Regra opcional: se for moeda "00", força Tasa=1
    # ATENÇÃO: no seu código "Cod. Moneda" = '01' por padrão; então essa condição pode nunca ativar.
    # Ajuste para o código correto da moeda USD no seu cenário.
    if "Cod. Moneda" in df.columns:
        df.loc[df["Cod. Moneda"] == "00", "Tasa"] = 1

    df = adicionar_cod_autorizacion_ext(df)
    df = adicionar_tip_fAC_ext(df)  # ajuste o nome exato da função conforme o seu arquivo
    df = organizar_colunas_externos(df)

    df = remover_duplicatas_source_file(df)

    if progress_widget:
        progress_widget.progress(100, text="Concluído (Externos).")
    if status_widget:
        status_widget.success("Pipeline Externos finalizado.")

    return df
