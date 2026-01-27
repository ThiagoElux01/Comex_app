
import pandas as pd
import re
from datetime import datetime

def adicionar_tasa_sharepoint(df, tasa_df):
    """
    Adiciona a coluna Tasa ao DataFrame SharePoint baseado na coluna Fecha_Emision.
    """

    if tasa_df is None or tasa_df.empty:
        df["Tasa"] = ""
        return df

    # Converte as datas do DF SharePoint
    df_tmp = df.copy()
    df_tmp["Fecha_Emision_tmp"] = pd.to_datetime(
        df_tmp["Fecha_Emision"], format="%d/%m/%Y", errors="coerce"
    )

    # Prepara TASA
    tasa = tasa_df.copy()
    tasa["Data"] = pd.to_datetime(tasa["Data"], errors="coerce")

    # Merge baseado na data
    df_tmp = df_tmp.merge(
        tasa[["Data", "Venta"]],
        left_on="Fecha_Emision_tmp",
        right_on="Data",
        how="left"
    )

    # Renomeia
    df_tmp.rename(columns={"Venta": "Tasa"}, inplace=True)

    # Limpa colunas auxiliares
    df_tmp = df_tmp.drop(columns=["Fecha_Emision_tmp", "Data"], errors="ignore")

    # Converte Tasa para número
    df_tmp["Tasa"] = pd.to_numeric(df_tmp["Tasa"], errors="coerce")

    return df_tmp
# ============================================================
# FUNÇÃO UNIVERSAL PARA CORRIGIR DATAS DO SHAREPOINT
# ============================================================
def corrigir_data_sharepoint(valor):
    """
    Converte datas de qualquer formato irregular do SharePoint para dd/mm/yyyy.
    Caso não seja possível converter, retorna ''.
    """

    if valor is None:
        return ""

    s = str(valor).strip()

    # Remover caracteres invisíveis
    s = s.replace("\u200b", "")      # zero-width space
    s = s.replace("\u00a0", " ")     # no-break space

    s = s.strip()
    if s == "":
        return ""

    # ---------------------------
    # 1) Extrair trecho que pareça data
    # ---------------------------
    padrao = re.compile(
        r'(\d{1,4}[-/]\d{1,2}[-/]\d{1,4})'
        r'|(\d{1,2}\s+[A-Za-zÁÉÍÓÚáéíóúñÑçÇâêôãõ]{3,15}\s+\d{2,4})'
    )
    m = padrao.search(s)
    if m:
        s = m.group(0)

    # ---------------------------
    # 2) Formatos diretos
    # ---------------------------
    formatos = [
        "%d/%m/%Y", "%d-%m-%Y",
        "%Y/%m/%d", "%Y-%m-%d",
        "%m/%d/%Y", "%m-%d-%Y",
        "%d/%m/%y", "%d-%m-%y",
        "%d %b %Y", "%d %B %Y",
        "%d %b %y", "%d %B %y",
    ]

    for fmt in formatos:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%d/%m/%Y")
        except:
            pass

    # ---------------------------
    # 3) Tentativa com meses PT/ES convertidos para inglês
    # ---------------------------
    meses = {
        # PT
        "jan": "Jan", "janeiro": "Jan",
        "mar": "mar", "março": "Mar",
        "abr": "Apr", "abril": "Apr",
        "mai": "May", "maio": "May",
        "jun": "Jun", "junho": "Jun",
        "jul": "Jul", "julho": "Jul",
        "ago": "Aug", "agosto": "Aug",
        "set": "Sep", "setembro": "Sep",
        "out": "Oct", "outubro": "Oct",
        "nov": "Nov", "novembro": "Nov",
        "dez": "Dec", "dezembro": "Dec",

        # ES
        "ene": "Jan", "enero": "Jan",
        "feb": "Feb", "febrero": "Feb",
        "mar": "Mar", "marzo": "Mar",
        "abr": "Apr", "abril": "Apr",
        "may": "May", "mayo": "May",
        "jun": "Jun", "junio": "Jun",
        "jul": "Jul", "julio": "Jul",
        "ago": "Aug", "agosto": "Aug",
        "sep": "Sep", "sept": "Sep", "septiembre": "Sep",
        "oct": "Oct", "octubre": "Oct",
        "nov": "Nov", "noviembre": "Nov",
        "dic": "Dec", "diciembre": "Dec",
    }

    s_proc = s.lower()
    for mes_local, mes_en in meses.items():
        s_proc = re.sub(rf"\b{mes_local}\b", mes_en, s_proc)

    s_proc = s_proc.title()

    for fmt in ["%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%d-%B-%Y"]:
        try:
            dt = datetime.strptime(s_proc, fmt)
            return dt.strftime("%d/%m/%Y")
        except:
            pass

    # ---------------------------
    # 4) Não converteu → return ""
    # ---------------------------
    return ""


# ============================================================
# AJUSTAR SHAREPOINT DF
# ============================================================
def ajustar_sharepoint_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # 1) Normalizar nomes de colunas
    df.columns = (
        df.columns
        .str.strip()
        .str.replace(" ", "_")
        .str.replace("-", "_")
        .str.lower()
    )

    # 2) IMPORTES NUMÉRICOS
    possiveis_nomes_importe = [
        "importe_documento",
        "importe_del_documento",
        "importe",
    ]

    def clean_number(value):
        if value is None:
            return None
        s = str(value).strip()
        s = re.sub(r"[^\d,.-]", "", s)

        if "." in s and "," in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except:
            return None

    for col in possiveis_nomes_importe:
        if col in df.columns:
            df[col] = df[col].apply(clean_number)

    # 3) DATAS
    possiveis_nomes_data = [
        "fecha_de_emision_del_documento",
        "fecha_emision_documento",
        "fecha",
        "fecha_de_emisipn_del_documento"
    ]

    for col in possiveis_nomes_data:
        if col in df.columns:
            df[col] = df[col].apply(corrigir_data_sharepoint)

    # 4) PROVEEDOR → texto antes do '-'
    if "proveedor" in df.columns:
        df["proveedor"] = (
            df["proveedor"]
            .astype(str)
            .str.split("-", n=1)
            .str[0]
            .str.strip()
        )

    return df
