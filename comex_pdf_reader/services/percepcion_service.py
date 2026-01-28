# services/percepcion_service.py
from io import BytesIO
from typing import List, Optional
import re
from datetime import datetime
import unicodedata
import fitz  # PyMuPDF
import pandas as pd

def _extract_first_page_lines_to_df(pdf_bytes: bytes) -> pd.DataFrame:
    """
    Extrai as linhas (spans) da primeira página e retorna DataFrame com a primeira coluna 'Text'
    e colunas subsequentes 'Col_1', 'Col_2', ...
    """
    try:
        doc = fitz.open(stream=BytesIO(pdf_bytes), filetype="pdf")
        page = doc.load_page(0)
        tdict = page.get_text("dict")
        blocks = tdict.get("blocks", [])
        linhas = []
        for bloco in blocks:
            for linha in bloco.get("lines", []):
                spans = [span.get("text", "") for span in linha.get("spans", [])]
                if spans:
                    linhas.append(spans)
        if not linhas:
            return pd.DataFrame()
        df = pd.DataFrame(linhas)
        cols = ["Text"] + [f"Col_{i}" for i in range(1, df.shape[1])]
        df.columns = cols[:df.shape[1]]
        return df
    except Exception:
        return pd.DataFrame()

def _add_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Replica a lógica do EXE: No_Liquidacion, CDA, Fecha, Monto (+ limpezas).
    """

    # ---------------------------
    # Helpers de normalização
    # ---------------------------
    def _clean_invisibles(s: str) -> str:
        if s is None:
            return ""
        s = str(s)
        # remove invisíveis comuns
        s = s.replace("\u200b", "").replace("\u00a0", " ")
        # colapsa múltiplos espaços
        s = re.sub(r"\s+", " ", s)
        return s.strip()

    def _upper_no_accents(s: str) -> str:
        s = _clean_invisibles(s)
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
        return s.upper()

    # ---------------------------
    # No_Liquidacion
    # ---------------------------
    def extrair_valor(row):
        texto = _upper_no_accents(row.get("Text", ""))
        if "NUMERO DE LIQUIDACION" in texto and ":" in texto:
            # pega o que está à direita dos dois pontos
            parts = texto.split(":", 1)
            return parts[1].strip() if len(parts) > 1 else ""
        elif "NÚMERO DE LIQU" in texto or "NUMERO DE LIQU" in texto:
            return _clean_invisibles(row.get("Col_1", "")) or ""
        return ""

    # ---------------------------
    # CDA (robusta, com fallback)
    # ---------------------------
    def extrair_valor_cda(row, next_row=None):
        """
        Extrai CDA com tolerância:
          - detecta 'CDA', 'C.D.A', 'C.D.A.' e 'C D A'
          - lê Col_1..Col_3 se existirem
          - tenta após ':' na mesma linha
          - fallback: próxima linha (se o valor estiver abaixo)
        """
        texto = _upper_no_accents(row.get("Text", ""))

        # Detecta o marcador CDA na linha
        has_cda = bool(re.search(r"\bC\.?\s*D\.?\s*A\.?\b", texto))
        if not has_cda:
            return ""

        # 1) Primeiro tenta colunas auxiliares
        for k in ("Col_1", "Col_2", "Col_3", "Col_4"):
            if k in row:
                cand = _clean_invisibles(row.get(k, ""))
                if cand and cand.strip():
                    out = cand.replace(" ", "")
                    out = re.sub(r"\s*-\s*", "-", out)
                    return out

        # 2) Tenta após ':' na mesma linha (cobre 'C.D.A.:', 'C.D.A :', etc.)
        m = re.search(r":\s*(.+)$", texto)
        if m:
            out = _clean_invisibles(m.group(1))
            out = out.replace(" ", "")
            out = re.sub(r"\s*-\s*", "-", out)
            return out

        # 3) Fallback: próxima linha
        if next_row is not None:
            nxt_text = _clean_invisibles(next_row.get("Text", ""))
            if nxt_text and not re.search(r"\bC\.?\s*D\.?\s*A\.?\b", _upper_no_accents(nxt_text)):
                out = nxt_text.replace(" ", "")
                out = re.sub(r"\s*-\s*", "-", out)
                return out

            for k in ("Col_1", "Col_2", "Col_3"):
                if k in next_row:
                    cand = _clean_invisibles(next_row.get(k, ""))
                    if cand.strip():
                        out = cand.replace(" ", "")
                        out = re.sub(r"\s*-\s*", "-", out)
                        return out

        return ""

    # ---------------------------
    # Fecha
    # ---------------------------
    def extrair_fecha(row):
        texto = _upper_no_accents(row.get("Text", ""))
        col1 = _clean_invisibles(row.get("Col_1", ""))

        m = re.search(r"DE FECHA\s*:\s*([\d]{2}[/-][\d]{2}[/-][\d]{4})", texto)
        if m:
            try:
                return datetime.strptime(m.group(1), "%d/%m/%Y").strftime("%d/%m/%y")
            except ValueError:
                return datetime.strptime(m.group(1), "%d-%m-%Y").strftime("%d/%m/%y")

        m2 = re.search(r"\b(\d{8})\b", col1)
        if m2:
            try:
                return datetime.strptime(m2.group(1), "%Y%m%d").strftime("%d/%m/%y")
            except ValueError:
                return ""

        return ""

    # ---------------------------
    # Monto (linha após 'SUNAT PERCEPCION IGV')
    # ---------------------------
    def extrair_monto(df_lines: pd.DataFrame):
        out = []
        for i in range(len(df_lines)):
            texto = _upper_no_accents(df_lines.at[i, "Text"])
            if "SUNAT PERCEPCION IGV" in texto:
                out.append(df_lines.at[i + 1, "Text"] if i + 1 < len(df_lines) else "")
            else:
                out.append("")
        return out

    # Aplicações
    df["No_Liquidacion"] = df.apply(extrair_valor, axis=1)
    df["Fecha"] = df.apply(extrair_fecha, axis=1)
    df["Monto"] = extrair_monto(df)

    # CDA com acesso à linha seguinte (fallback)
    cda_vals = []
    for i in range(len(df)):
        row = df.iloc[i].to_dict()
        next_row = df.iloc[i + 1].to_dict() if (i + 1) < len(df) else None
        cda_vals.append(extrair_valor_cda(row, next_row))
    df["CDA"] = cda_vals

    # Limpeza básica
    for col in ["No_Liquidacion", "CDA", "Monto", "Fecha"]:
        df[col] = df[col].apply(lambda x: str(x).strip() if pd.notna(x) else "")

    # Converte Monto → float (usa '.' como decimal após remover ',')
    def to_float(v):
        s = str(v).replace(",", "").strip()
        return round(float(s), 2) if s and s.replace(".", "", 1).isdigit() else None
    df["Monto"] = df["Monto"].apply(to_float)

    # Ajuste do CDA (ex.: "<xx> ... <dddddd>" → "xx-dddddd")
    def ajustar_cda(v):
        m = re.search(r"\b(\d{2,3})\D+.*?(\d{6,})\b", str(v))
        return f"{m.group(1)}-{m.group(2)}" if m else v
    df["CDA"] = df["CDA"].apply(ajustar_cda)

    # Remover sufixos indesejados do No_Liquidacion
    padroes_remover = ["-25", "-26", "-24", "-23", "-27"]
    regex = re.compile(r"(" + "|".join(map(re.escape, padroes_remover)) + r")\b")
    df["No_Liquidacion"] = df["No_Liquidacion"].apply(
        lambda x: regex.sub("", str(x)) if pd.notna(x) else x
    )

    return df

def _consolidar_por_arquivo(df_lines: pd.DataFrame) -> pd.DataFrame:
    """
    Consolida por Source_File pegando o primeiro valor não vazio de cada campo.
    """
    dados = []
    for src in df_lines["Source_File"].unique():
        dfa = df_lines[df_lines["Source_File"] == src]
        pick = lambda s: s.dropna().replace("", pd.NA).dropna()
        no_liq = pick(dfa["No_Liquidacion"])
        cda    = pick(dfa["CDA"])
        fecha  = pick(dfa["Fecha"])
        monto  = pick(dfa["Monto"])
        dados.append([
            src,
            no_liq.iloc[0] if not no_liq.empty else "",
            cda.iloc[0] if not cda.empty else "",
            fecha.iloc[0] if not fecha.empty else "",
            monto.iloc[0] if not monto.empty else ""
        ])
    return pd.DataFrame(dados, columns=["Source_File", "No_Liquidacion", "CDA", "Fecha", "Monto"])

def process_percepcion_streamlit(
    uploaded_files: List,
    progress_widget=None,
    status_widget=None,
) -> Optional[pd.DataFrame]:
    """
    Pipeline Percepciones para Streamlit.
    Lê PDFs do uploader, aplica as regras e retorna o DataFrame final (sem salvar em disco).
    Mantém Tasa = 1.00 conforme lógica original do seu EXE.
    """
    if not uploaded_files:
        return None

    if progress_widget:
        progress_widget.progress(0, text="Lendo PDFs (Percepciones)...")

    dfs = []
    total = len(uploaded_files)
    for i, f in enumerate(uploaded_files, start=1):
        fname = getattr(f, "name", f"arquivo_{i}.pdf")
        lines_df = _extract_first_page_lines_to_df(f.getvalue())
        if not lines_df.empty:
            lines_df.insert(0, "Source_File", fname)
            dfs.append(lines_df)
        if progress_widget:
            progress_widget.progress(int(i / total * 100), text=f"Lendo {fname} ({i}/{total})")
        if status_widget:
            status_widget.write(f"📄 Primeira página lida: **{fname}**")

    if not dfs:
        return None

    df_all = pd.concat(dfs, ignore_index=True)
    df_all = _add_columns(df_all)

    # Consolidação por arquivo
    df_rel = _consolidar_por_arquivo(df_all)

    # Pós-processo (mesma lógica do EXE)
    df_rel["Error"] = df_rel["No_Liquidacion"].apply(
        lambda x: "Can't read the file" if pd.isna(x) or str(x).strip() == "" else ""
    )
    df_rel["Fecha"] = df_rel["Fecha"].astype(str).str.replace("/", "", regex=False)

    # Colunas fixas
    df_rel["Tasa"] = 1.00
    df_rel["COD PROVEEDOR"] = "13131295"
    df_rel["COD MONEDA"] = "00"
    df_rel["Cód. de Autorización"] = "54"
    df_rel["Cuenta"] = "421201"
    df_rel["Tipo de Factura"] = "12"

    # Ordem final
    df_rel = df_rel[
        [
            "Source_File",
            "COD PROVEEDOR",
            "No_Liquidacion",
            "Fecha",
            "CDA",
            "Monto",
            "Tasa",
            "COD MONEDA",
            "Cód. de Autorización",
            "Tipo de Factura",
            "Cuenta",
            "Error",
        ]
    ]

    if progress_widget:
        progress_widget.progress(100, text="Concluído (Percepciones).")
    return df_rel
