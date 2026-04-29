# services/adicionales_utils.py
import re
import unicodedata
from datetime import datetime
import pandas as pd
import streamlit as st

# ============================================================
# HELPERS GERAIS
# ============================================================

def _norm_text(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.upper().strip()


# ============================================================
# RUC / FACTURA
# ============================================================

def extrair_ruc(texto: str) -> str:
    for pat in [
        r'R\.U\.C.*?(\d{11})',
        r'RUC:\s*(\d{11})',
        r'RUC N°\s*(\d{11})',
    ]:
        m = re.search(pat, texto)
        if m:
            return m.group(1)
    return ""


def extrair_facturas(texto: str) -> str:
    patterns = [
        r'F\d{3}[-\s]*\d{9}',
        r'F\d{3}[-\s]*\d{8}',
        r'F\d{3}[-\s]*\d{5,7}',
        r'F\d{2}[-\s]*\d{5,7}',
        r'\bPECLLP\d{9}\b',
        r'INV-[A-Z]+-\d{8}',
    ]
    for p in patterns:
        m = re.search(p, texto)
        if m:
            return m.group(0).replace(" ", "")
    return ""


# ============================================================
# PROVEEDOR ISCALΑ
# ============================================================

def criar_coluna_proveedor_iscala(df: pd.DataFrame) -> pd.DataFrame:
    def resolver(row):
        txt = row.get("conteudo_pdf", "")
        if "EVERGREEN" in txt:
            return "EVERGREEN"
        if "MSC MEDITERRANEAN" in txt:
            return "MSC"
        if "WANHAI" in txt or "WAN HAI" in txt:
            return "WAN HAI"
        ruc = row.get("R.U.C", "")
        return ruc[2:-1] if isinstance(ruc, str) and len(ruc) >= 5 else ""
    df["Proveedor Iscala"] = df.apply(resolver, axis=1)
    return df


# ============================================================
# FECHA DE EMISIÓN
# ============================================================

def extrair_fecha_emision(texto: str) -> str:
    linhas = texto.splitlines()
    for ln in linhas:
        m = re.search(r'\b(\d{2}[-/]\d{2}[-/]\d{4}|\d{4}[-/]\d{2}[-/]\d{2})\b', ln)
        if m:
            return m.group(1)
    return ""


def normalizar_data(data: str) -> str:
    formatos = [
        "%d/%m/%Y", "%d-%m-%Y",
        "%Y/%m/%d", "%Y-%m-%d",
        "%d-%b-%Y", "%d-%B-%Y",
    ]
    for f in formatos:
        try:
            return datetime.strptime(data.strip(), f).strftime("%d/%m/%Y")
        except Exception:
            pass
    return data


# ============================================================
# MOEDA
# ============================================================

def ajustar_e_padronizar_moneda(valor: str) -> str:
    t = _norm_text(valor)
    if "USD" in t or "DOLAR" in t or "US$" in t:
        return "USD"
    if "PEN" in t or "SOLES" in t or "S/" in t:
        return "PEN"
    return valor


def codificar_moneda(valor: str) -> str:
    if valor == "USD":
        return "01"
    if valor == "PEN":
        return "00"
    return ""


# ============================================================
# OP. GRAVADA
# ============================================================

def limpar_op_gravada(v):
    if isinstance(v, str):
        return re.sub(r"[^0-9.,-]", "", v)
    return v


def formatar_op_gravada(v):
    try:
        v = str(v).replace(",", ".")
        return float(v)
    except Exception:
        return None


def op_gravada_negativo_CN(df: pd.DataFrame) -> pd.DataFrame:
    if "Tipo Doc" in df.columns and "Op. Gravada" in df.columns:
        mask = df["Tipo Doc"].astype(str).str.upper().str.contains("CREDITO")
        df.loc[mask, "Op. Gravada"] = -abs(df.loc[mask, "Op. Gravada"])
    return df


# ============================================================
# TIPO DOC
# ============================================================

def padronizar_tipo_doc(df: pd.DataFrame) -> pd.DataFrame:
    subs = {
        "INVOICE": "FACTURA",
        "FACTURA ELECTRONICA": "FACTURA",
        "NOTA DE CREDITO": "NOTA DE CRÉDITO",
        "CREDIT NOTE": "NOTA DE CRÉDITO",
    }
    df["Tipo Doc"] = df["Tipo Doc"].replace(subs)
    return df


# ============================================================
# CÓDIGOS SUNAT (ADICIONALES)
# ============================================================

def adicionar_cod_autorizacion_adicionales(df: pd.DataFrame) -> pd.DataFrame:
    if "Cód. de Autorización" not in df.columns:
        df["Cód. de Autorización"] = ""
    tipo = df["Tipo Doc"].astype(str).str.upper()
    df.loc[(df["Cód. de Autorización"] == "") & (tipo == "FACTURA"), "Cód. de Autorización"] = "01"
    df.loc[(df["Cód. de Autorización"] == "") & (tipo.str.contains("CREDITO")), "Cód. de Autorización"] = "07"
    return df


def adicionar_tipo_factura_adicionales(df: pd.DataFrame) -> pd.DataFrame:
    if "Tipo de Factura" not in df.columns:
        df["Tipo de Factura"] = ""
    df.loc[df["Tipo de Factura"] == "", "Tipo de Factura"] = "01"
    return df


# ============================================================
# TASA (REAPROVEITADA DO DUAS)
# ============================================================

from services.duas_utils import adicionar_coluna_tasa


# ============================================================
# SHAREPOINT – ADICIONALES
# ============================================================

def merge_sharepoint_adicionales(df_adic: pd.DataFrame, df_sp: pd.DataFrame) -> pd.DataFrame:
    if df_sp is None or df_sp.empty:
        return df_adic

    df_adic = df_adic.copy()
    df_sp = df_sp.copy()

    df_sp.columns = df_sp.columns.astype(str).str.lower()

    def norm(s):
        if s is None:
            return ""
        s = str(s)
        s = s.replace("\u200b", "").replace("\u00a0", " ")
        return re.sub(r"\s+", " ", s).lower().strip()

    df_adic["_k"] = df_adic["source_file"].apply(norm)
    df_sp["_k"] = df_sp["name"].apply(norm)

    df_sp_sel = df_sp[[
        c for c in [
            "tasa_sharepoint", "numero_de_documento", "importe_documento",
            "moneda", "fecha_emision"
        ] if c in df_sp.columns
    ] + ["_k"]]

    df_final = df_adic.merge(df_sp_sel, on="_k", how="left")
    df_final = df_final.drop(columns=["_k"], errors="ignore")

    return df_final


def adicionar_sharepoint_adicionales(df_adic: pd.DataFrame, df_sp: pd.DataFrame) -> pd.DataFrame:
    return merge_sharepoint_adicionales(df_adic, df_sp)


# ============================================================
# ORGANIZAÇÃO FINAL
# ============================================================

def organizar_colunas_adicionales(df: pd.DataFrame) -> pd.DataFrame:
    desejadas = [
        "source_file", "R.U.C", "Proveedor Iscala", "Factura", "Tipo Doc",
        "Cód. de Autorización", "Tipo de Factura", "Fecha de Emisión",
        "Moneda", "Cod. Moneda", "Op. Gravada", "Tasa", "Cuenta", "Error"
    ]
    base = [c for c in desejadas if c in df.columns]
    resto = [c for c in df.columns if c not in base]
    return df[base + resto]


def remover_duplicatas_source_file(df: pd.DataFrame) -> pd.DataFrame:
    return df.drop_duplicates(subset="source_file", keep="first")
