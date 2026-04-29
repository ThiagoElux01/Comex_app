import pandas as pd
import re
from datetime import datetime
import streamlit as st

# ============================================================
# ADICIONAR TASA SHAREPOINT (MERGE COM TASA SUNAT)
# ============================================================

def adicionar_tasa_sharepoint(df: pd.DataFrame, tasa_df: pd.DataFrame | None) -> pd.DataFrame:
    """
    Adiciona a coluna Tasa_Sharepoint ao DataFrame SharePoint,
    fazendo merge seguro baseado apenas na data (ignorando hora/timezone).
    Retorna SEMPRE Tasa_Sharepoint como STRING.
    """
    df = df.copy()

    # Garante coluna
    
    if "tasa_sharepoint" not in df.columns:
    df["tasa_sharepoint"] = ""

    # Se não houver Tasa SUNAT carregada, retorna como está
    if tasa_df is None or tasa_df.empty:
        df["Tasa_Sharepoint"] = df["Tasa_Sharepoint"].astype("string")
        return df

    # ------------------------------------------------------------
    # 1) Converter Fecha_Emision para datetime (robusto)
    # ------------------------------------------------------------
    df["Fecha_Emision_tmp"] = pd.to_datetime(
        df["Fecha_Emision"],
        dayfirst=True,
        errors="coerce"
    ).dt.normalize()

    # ------------------------------------------------------------
    # 2) Preparar Tasa SUNAT
    # ------------------------------------------------------------
    tasa = tasa_df.copy()
    tasa["Data"] = pd.to_datetime(
        tasa["Data"],
        dayfirst=True,
        errors="coerce"
    ).dt.normalize()

    # ------------------------------------------------------------
    # 3) Merge seguro
    # ------------------------------------------------------------
    df = df.merge(
        tasa[["Data", "Venta"]],
        left_on="Fecha_Emision_tmp",
        right_on="Data",
        how="left"
    )

    # ------------------------------------------------------------
    # 4) Renomear e limpar
    # ------------------------------------------------------------
    df.rename(columns={"Venta": "tasa_sharepoint"}, inplace=True)
    df.drop(columns=["Fecha_Emision_tmp", "Data"], inplace=True, errors="ignore")

    # ------------------------------------------------------------
    # 5) FORÇAR string (crítico p/ pandas + pyarrow)
    # ------------------------------------------------------------
    df["tasa_sharepoint"] = df["tasa_sharepoint"].astype("string)
                                                         
    # Padronizar nome final esperado pelos fluxos
    if "tasa_sharepoint" in df.columns:
        df.rename(columns={"tasa_sharepoint": "Tasa_Sharepoint"}, inplace=True)
                                                         
    return df


# ============================================================
# FUNÇÃO UNIVERSAL PARA CORRIGIR DATAS DO SHAREPOINT
# ============================================================

def corrigir_data_sharepoint(valor) -> str:
    """
    Converte datas de qualquer formato irregular do SharePoint para dd/mm/yyyy.
    Caso não seja possível converter, retorna "".
    """
    if valor is None:
        return ""

    s = str(valor).strip()
    s = s.replace("\u200b", "").replace("\u00a0", " ").strip()

    if s == "":
        return ""

    # Extrair possível trecho de data
    padrao = re.compile(
        r'(\d{1,4}[-/]\d{1,2}[-/]\d{1,4})'
        r'|(\d{1,2}\s+[A-Za-zÁÉÍÓÚáéíóúñÑçÇâêôãõ]{3,15}\s+\d{2,4})'
    )
    m = padrao.search(s)
    if m:
        s = m.group(0)

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
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except:
            pass

    meses = {
        "jan": "Jan", "janeiro": "Jan",
        "fev": "Feb", "febrero": "Feb",
        "mar": "Mar", "março": "Mar",
        "abr": "Apr", "abril": "Apr",
        "mai": "May", "maio": "May",
        "jun": "Jun", "junho": "Jun",
        "jul": "Jul", "julho": "Jul",
        "ago": "Aug", "agosto": "Aug",
        "set": "Sep", "septiembre": "Sep",
        "out": "Oct", "octubre": "Oct",
        "nov": "Nov", "noviembre": "Nov",
        "dez": "Dec", "diciembre": "Dec",
    }

    s_proc = s.lower()
    for mes_local, mes_en in meses.items():
        s_proc = re.sub(rf"\b{mes_local}\b", mes_en, s_proc)

    s_proc = s_proc.title()

    for fmt in ["%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%d-%B-%Y"]:
        try:
            return datetime.strptime(s_proc, fmt).strftime("%d/%m/%Y")
        except:
            pass

    return ""


# ============================================================
# AJUSTAR SHAREPOINT DF (FUNÇÃO PRINCIPAL)
# ============================================================

def ajustar_sharepoint_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # ------------------------------------------------------------
    # 0) Garantir Tasa_Sharepoint como STRING desde o início
    # ------------------------------------------------------------
    #if "Tasa_Sharepoint" not in df.columns:
    #    df["Tasa_Sharepoint"] = ""
    #df["Tasa_Sharepoint"] = df["Tasa_Sharepoint"].astype("string")

    # ------------------------------------------------------------
    # 1) Normalizar nomes das colunas
    # ------------------------------------------------------------
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace(" ", "_")
        .str.replace("-", "_")
        .str.lower()
    )

    # ------------------------------------------------------------
    # 2) IMPORTES NUMÉRICOS
    # ------------------------------------------------------------
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

    # ------------------------------------------------------------
    # 3) DATA → Fecha_Emision
    # ------------------------------------------------------------
    colunas_data_possiveis = [
        "fecha_de_emisipn_del_documento",
        "fecha_de_emision_del_documento",
        "fecha_emision_documento",
        "fecha",
        "fech_emision",
        "fechadeemision",
        "emision",
    ]

    col_data_original = None
    for c in colunas_data_possiveis:
        if c in df.columns:
            col_data_original = c
            break

    if col_data_original:
        df["Fecha_Emision"] = df[col_data_original].apply(corrigir_data_sharepoint)
    else:
        df["Fecha_Emision"] = ""

    # ------------------------------------------------------------
    # 4) PROVEEDOR → texto antes do "-"
    # ------------------------------------------------------------
    if "proveedor" in df.columns:
        df["proveedor"] = (
            df["proveedor"]
            .astype(str)
            .str.split("-", n=1)
            .str[0]
            .str.strip()
        )

    # ------------------------------------------------------------
    # 5) MERGE COM TASA SUNAT
    # ------------------------------------------------------------
    tasa_df = st.session_state.get("tasa_df")
    df = adicionar_tasa_sharepoint(df, tasa_df)

    # ------------------------------------------------------------
    # 6) REGRA DE NEGÓCIO → PEN = 1
    # ------------------------------------------------------------
    if "moneda" in df.columns:
        df["Tasa_Sharepoint"] = df["Tasa_Sharepoint"].astype("string")
        df.loc[
            df["moneda"].astype(str).str.upper().str.strip() == "PEN",
            "Tasa_Sharepoint"
        ] = "1"

    return df
