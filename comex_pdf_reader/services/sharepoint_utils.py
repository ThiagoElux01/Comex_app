import pandas as pd
import re
from datetime import datetime
import streamlit as st


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
        except Exception:
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
        except Exception:
            pass

    return ""


# ============================================================
# ADICIONAR TASA SHAREPOINT (MERGE COM TASA SUNAT)
# ============================================================

def adicionar_tasa_sharepoint(df: pd.DataFrame, tasa_df: pd.DataFrame | None) -> pd.DataFrame:
    """
    Adiciona/preenche a coluna 'tasa_sharepoint' no DataFrame SharePoint,
    fazendo merge seguro baseado apenas na data (ignorando hora/timezone).
    Mantém sempre 'tasa_sharepoint' como string e NÃO cria colunas duplicadas.
    """
    df = df.copy()

    COL_TASA = "tasa_sharepoint"
    COL_FECHA = "fecha_emision"

    # garante coluna destino como string
    if COL_TASA not in df.columns:
        df[COL_TASA] = pd.Series("", index=df.index, dtype="string")
    else:
        df[COL_TASA] = df[COL_TASA].astype("string")

    # sem tasa SUNAT → retorna
    if tasa_df is None or tasa_df.empty:
        return df

    # se não houver data, não tem como mergear
    if COL_FECHA not in df.columns:
        return df

    # 1) normaliza fecha emision no df
    fecha_tmp = pd.to_datetime(df[COL_FECHA], dayfirst=True, errors="coerce").dt.normalize()

    # 2) prepara tasa SUNAT
    tasa = tasa_df.copy()
    tasa_date = pd.to_datetime(tasa["Data"], dayfirst=True, errors="coerce").dt.normalize()
    tasa_out = pd.DataFrame({"_data": tasa_date, "_venta": tasa["Venta"]})

    # 3) merge com suffix para não colidir
    df["_fecha_tmp"] = fecha_tmp
    df = df.merge(
        tasa_out,
        left_on="_fecha_tmp",
        right_on="_data",
        how="left"
    )

    # 4) coalesce: se tasa_sharepoint estiver vazia/NA, usa a do SUNAT
    #    (convertendo para string)
    df[COL_TASA] = df[COL_TASA].replace("", pd.NA)
    df[COL_TASA] = df[COL_TASA].fillna(df["_venta"].astype("string"))
    df[COL_TASA] = df[COL_TASA].fillna("").astype("string")

    # 5) limpa colunas temporárias
    df.drop(columns=["_fecha_tmp", "_data", "_venta"], inplace=True, errors="ignore")

    return df


# ============================================================
# AJUSTAR SHAREPOINT DF (FUNÇÃO PRINCIPAL)
# ============================================================

def ajustar_sharepoint_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # ------------------------------------------------------------
    # 1) Normalizar nomes das colunas PRIMEIRO (chave para não duplicar!)
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
    # 2) Garantir tasa_sharepoint como STRING (canônico)
    # ------------------------------------------------------------
    if "tasa_sharepoint" not in df.columns:
        df["tasa_sharepoint"] = pd.Series("", index=df.index, dtype="string")
    else:
        df["tasa_sharepoint"] = df["tasa_sharepoint"].astype("string")

    # ------------------------------------------------------------
    # 3) IMPORTES NUMÉRICOS
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
        except Exception:
            return None

    for col in possiveis_nomes_importe:
        if col in df.columns:
            df[col] = df[col].apply(clean_number)

    # ------------------------------------------------------------
    # 4) DATA → fecha_emision (canônico)
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

    col_data_original = next((c for c in colunas_data_possiveis if c in df.columns), None)

    if col_data_original:
        df["fecha_emision"] = df[col_data_original].apply(corrigir_data_sharepoint)
    else:
        df["fecha_emision"] = ""

    # ------------------------------------------------------------
    # 5) PROVEEDOR → texto antes do "-"
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
    # 6) MERGE COM TASA SUNAT (sem duplicar colunas)
    # ------------------------------------------------------------
    tasa_df = st.session_state.get("tasa_df")
    df = adicionar_tasa_sharepoint(df, tasa_df)

    # ------------------------------------------------------------
    # 7) REGRA DE NEGÓCIO → PEN = 1
    # ------------------------------------------------------------
    if "moneda" in df.columns:
        df.loc[
            df["moneda"].astype(str).str.upper().str.strip() == "PEN",
            "tasa_sharepoint"
        ] = "1"

    # reforça dtype para evitar surprises no arrow
    df["tasa_sharepoint"] = df["tasa_sharepoint"].astype("string")

    return df
