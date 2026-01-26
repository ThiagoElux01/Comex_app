
# services/adicionales_utils.py
import re
import unicodedata
from datetime import datetime
import pandas as pd

# --- EXTRAÇÕES BÁSICAS ---

def extrair_ruc(texto: str) -> str:
    match1 = re.search(r'R\.U\.C.*?(\d{11})', texto)
    if match1:
        return match1.group(1).strip()
    match2 = re.search(r'RUC:\s*(\d{11})', texto)
    if match2:
        return match2.group(1).strip()
    match3 = re.search(r'RUC N°\s*(\d{11})', texto)
    if match3:
        return match3.group(1).strip()
    return ""

def extrair_facturas(texto: str) -> str:
    for pattern in [
        r'F\d{3}[-\s]*\d{9}',
        r'F\d{3}[-\s]*\d{8}',
        r'F\d{3}[-\s]*\d{5,7}',
        r'F\d{2}[-\s]*\d{5,7}',
        r'INV-[A-Z]+-\d{8}',
        r'Número de Invoice\(Invoice No\.\)\s*:\s*([A-Z]{4}\d{9})',
        r'\bPECLLP\d{9}\b',
        r'F\d{3}[-\s]*\d{4}',
    ]:
        m = re.search(pattern, texto)
        if m:
            g = m.group(1) if m.lastindex else m.group(0)
            return g.replace(" ", "").strip()
    return ""

def remover_ruc_indesejado(df: pd.DataFrame, ruc_indesejado="20100073308") -> pd.DataFrame:
    df["R.U.C"] = df["R.U.C"].apply(lambda x: "" if x == ruc_indesejado else x)
    return df

def criar_coluna_proveedor_iscala(df: pd.DataFrame) -> pd.DataFrame:
    def definir_valor(row):
        txt = row["conteudo_pdf"]
        if "EVERGREEN LINE" in txt:
            return "EVERGREEN"
        elif "MSC Mediterranean Shipping Company S.A." in txt:
            return "MSC"
        elif "WANHAI" in txt:
            return "WAN HAI"
        elif row["R.U.C"]:
            return row["R.U.C"][2:-1]
        else:
            return ""
    df["Proveedor Iscala"] = df.apply(definir_valor, axis=1)
    return df


# --- FECHA DE EMISIÓN ---

def extrair_fecha_emision(texto: str) -> str:
    linhas = texto.splitlines()
    for i in range(len(linhas)):
        linha = linhas[i].strip()
        linha_up = linha.upper()

        m_emision = re.search(r'F\.?\s*DE\s+EMISI[ÓO]N\s*[:\-]?\s*(\d{4}[-/]\d{2}[-/]\d{2})', linhas[i], re.IGNORECASE)
        if m_emision:
            return m_emision.group(1)

        if linha_up == "F. DE" and i + 1 < len(linhas):
            proxima = linhas[i + 1].strip()
            m = re.search(r'[:\-]?\s*(\d{4}[-/]\d{2}[-/]\d{2})', proxima)
            if m:
                return m.group(1)

        if "FECHA DE EMISIÓN" in linha_up or "FECHA DE EMISION" in linha_up:
            m_inline = re.search(r'FECHA DE EMISI[ÓO]N[:\s]*([0-9]{2}[-/][0-9]{2}[-/][0-9]{4})', linha_up)
            if m_inline:
                return m_inline.group(1)
            if i > 0:
                prev = linhas[i - 1].strip()
                if re.match(r'\d{2}[-/]\d{2}[-/]\d{4}', prev) or re.match(r'\d{4}[-/]\d{2}[-/]\d{2}', prev):
                    return prev

        if linha_up == "FECHA" and i + 2 < len(linhas):
            if linhas[i + 1].strip().upper() == "EMISIÓN":
                data_line = linhas[i + 2].strip()
                m = re.search(r'\d{2}[-/]\d{2}[-/]\d{4}', data_line)
                if m:
                    return m.group(0)

        if "R.U.C. N°" in linha_up and i + 1 < len(linhas):
            prox = linhas[i + 1].strip()
            if re.match(r'\d{4}[-/]\d{2}[-/]\d{2}', prox):
                return prox

        if "DOLARES AMERICANOS" in linha_up and i >= 2:
            cand = linhas[i - 2].strip()
            if re.match(r'\d{2}[-/]\d{2}[-/]\d{4}', cand):
                return cand

        if linha_up in ("FECHA DE EMISIÓN", "FECHA DE EMISION") and i > 0:
            prev = linhas[i - 1].strip()
            if re.match(r'\d{4}[-/]\d{2}[-/]\d{2}', prev):
                return prev

        if "FECHA EMISIÓN:" in linha_up or "FECHA DE EMISIÓN:" in linha_up:
            if i > 0:
                acima = linhas[i - 1].strip()
                m = re.search(r'\d{2}[-/]\d{2}[-/]\d{4}|\d{4}[-/]\d{2}[-/]\d{2}', acima)
                if m:
                    return m.group(0)

        if "FECHA DE EMISIÓN" in linha_up or "FECHA DE EMISION" in linha_up:
            for offset in range(1, 17):
                if i + offset < len(linhas):
                    ld = linhas[i + offset].strip()
                    if re.match(r'\d{4}[-/]\d{2}[-/]\d{2}', ld):
                        return ld

    for i in range(1, len(linhas)):
        if linhas[i].strip().upper() in ["FECHA:", "FECHA"]:
            ant = linhas[i - 1].strip()
            m = re.search(r'\d{2}/\d{2}/\d{4}', ant)
            if m:
                return m.group(0)

    for i in range(len(linhas) - 3):
        if "FACTURA" in linhas[i].strip().upper():
            ld = linhas[i + 3].strip()
            m = re.match(r'\d{2}-[A-Z][a-z]{2}-\d{4}', ld)
            if m:
                return m.group(0)

    m = re.search(r'FECHA EMISI[ÓO]N\(ISSUE DATE\)\s*[:\-]?\s*(\d{4}[-/]\d{2}[-/]\d{2})', texto.upper())
    if m:
        return m.group(1)
    return ""

def normalizar_data(data):
    formatos = ["%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d", "%Y-%m-%d", "%d-%b-%Y", "%d-%B-%Y"]
    if not isinstance(data, str):
        return data
    for fmt in formatos:
        try:
            dt = datetime.strptime(data.strip(), fmt)
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            continue
    return data


# --- MOEDA ---

def extrair_moneda(texto: str) -> str:
    linhas = texto.splitlines()
    palavras_chave = ["MONEDA", "CURRENCY", "TIPO DE CAMBIO", "WAN HAI", "GRAN TOTAL:"]
    padroes_moeda = ["DÓLAR", "DOLAR", "USD", "US DÓLARES", "SOLES", "PEN", "EUROS", "EUR"]

    for i, linha in enumerate(linhas):
        up = linha.upper()

        if any(p in up for p in palavras_chave):
            m_inline = re.search(r'(MONEDA|CURRENCY)\s*[:\-]?\s*([A-Z\s]+)', up)
            if m_inline:
                moeda = m_inline.group(2).strip()
                if any(m in moeda for m in padroes_moeda):
                    return moeda.title()

            for j in range(-5, 6):
                if j == 0:
                    continue
                idx = i + j
                if 0 <= idx < len(linhas):
                    prox = linhas[idx].strip().upper()
                    if any(m in prox for m in padroes_moeda):
                        return prox.title()
    return ""

def ajustar_e_padronizar_moneda(valor: str) -> str:
    if not isinstance(valor, str):
        return valor
    val = ''.join(c for c in unicodedata.normalize('NFD', valor) if unicodedata.category(c) != 'Mn').upper()
    m = re.search(r'(DOLARES.*)', val)
    if m:
        val = m.group(1).strip()
    subs = {
        "DOLARES": "USD", "DOLAR AMERICANO": "USD", "DOLARES AMERICANOS": "USD",
        "CTA CTE BBVA - USD": "USD", "USD": "USD", "DOLARES AMERICANOS (US$)": "USD",
        "SOLES (S/)": "PEN", "PEN": "PEN", "SOLES": "PEN", "S/": "PEN", "US$": "USD",
    }
    for k, v in subs.items():
        if k in val:
            return v
    return valor.strip().title()

def codificar_moneda(valor: str) -> str:
    if valor == "USD":
        return "01"
    elif valor == "PEN":
        return "00"
    return ""


# --- OP. GRAVADA ---

def extrair_op_gravada(row) -> str:
    linhas = row['conteudo_pdf'].splitlines()
    prov = row['Proveedor Iscala']
    tipo = str(row.get('Tipo Doc', '')).upper()

    if prov == '25206207' and tipo == 'NOTA DE CRÉDITO':
        for i, linha in enumerate(linhas):
            if "OP. GRAVADA" in linha.upper() and i >= 7:
                return linhas[i - 7].strip()

    if prov == '10001013':
        for i, linha in enumerate(linhas):
            if "SON:" in linha and i >= 8:
                return linhas[i - 8].strip()

    elif prov == '25981421':
        for i, linha in enumerate(linhas):
            if "SON:" in linha and i >= 10:
                return linhas[i - 10].strip()

    elif prov == '34528608':
        for i, linha in enumerate(linhas):
            if "Total Gravado" in linha and i + 1 < len(linhas):
                return linhas[i + 1].strip()

    elif prov == '60342509':
        for i, linha in enumerate(linhas):
            if "Total Valor de Venta - Operaciones Gravadas:" in linha and i + 1 < len(linhas):
                return linhas[i + 1].strip()

    elif prov == '25206207':
        for i, linha in enumerate(linhas):
            if "OP. INAFECTA" in linha and i >= 1:
                return linhas[i - 1].strip()

    elif prov == '51346238':
        for i, linha in enumerate(linhas):
            if "OP. GRAVADAS:" in linha and i >= 2:
                return linhas[i - 2].strip()

    elif prov == '60037433':
        for i, linha in enumerate(linhas):
            if "SON:" in linha and i + 8 < len(linhas):
                return linhas[i + 8].strip()

    elif prov == '10001021':
        for i, linha in enumerate(linhas):
            if "OP. GRAVADAS:" in linha and i >= 2:
                return linhas[i - 2].strip()

    elif prov == '51092775':
        for i, linha in enumerate(linhas):
            if "Operación gravada" in linha and i >= 1:
                return linhas[i - 1].strip()

    elif prov == '34764689':
        for i, linha in enumerate(linhas):
            if "Son: " in linha:
                if i + 1 < len(linhas):
                    return linhas[i + 1].strip()

    elif prov == 'WAN HAI':
        for i, linha in enumerate(linhas):
            if "Son:" in linha and i >= 2:
                return linhas[i - 2].strip()

    if prov == 'EVERGREEN':
        for linha in linhas:
            if "Total Amount(Monto total): " in linha:
                return linha.strip()

    elif prov == 'MSC':
        for i, linha in enumerate(linhas):
            if "SON:" in linha and i >= 5:
                return linhas[i - 5].strip()

    elif prov == '61092558':
        for i, linha in enumerate(linhas):
            if "Total Valor de Venta - Operaciones Gravadas:" in linha and i + 1 < len(linhas):
                return linhas[i + 1].strip()

    elif prov == '54308388':
        for i, linha in enumerate(linhas):
            if "Total Valor de Venta - Operaciones Gravadas:" in linha and i + 1 < len(linhas):
                return linhas[i + 1].strip()

    return ""

def limpar_op_gravada(valor):
    if isinstance(valor, str):
        return re.sub(r'[^0-9,\.]', '', valor)
    return valor

def formatar_op_gravada(valor):
    if isinstance(valor, str):
        if ',' in valor and '.' in valor:
            valor = valor.replace(',', '')
        valor = valor.replace('.', ',')  # visual
        valor = valor.replace(',', '.')  # numérico
        try:
            return float(valor)
        except ValueError:
            return None
    return valor

def op_gravada_negativo_CN(df: pd.DataFrame) -> pd.DataFrame:
    if 'Tipo Doc' in df.columns and 'Op. Gravada' in df.columns:
        df['Op. Gravada'] = df.apply(
            lambda r: -abs(r['Op. Gravada']) if str(r['Tipo Doc']).strip().upper() in ('NOTA DE CRÉDITO', 'NOTA DE CREDITO') else r['Op. Gravada'],
            axis=1
        )
    return df


# --- TIPO DOC ---

def extrair_tipo_doc(row) -> str:
    texto = row["conteudo_pdf"]
    fornecedor = row["Proveedor Iscala"]
    linhas = texto.splitlines()

    if fornecedor == "10001013" and len(linhas) >= 3:
        return linhas[2].strip()
    elif fornecedor == "34528608" and len(linhas) >= 8:
        return linhas[7].strip()
    elif fornecedor == "25981421" and len(linhas) >= 3:
        return linhas[2].strip()
    elif fornecedor == "60342509" and len(linhas) >= 3:
        return linhas[2].strip()
    elif fornecedor == "25206207":
        if len(linhas) >= 4 and "FACTURA" in linhas[3].upper():
            return linhas[3].strip()
        elif len(linhas) >= 6:
            return linhas[5].strip()
    elif fornecedor == "51346238":
        idxs = [i for i, ln in enumerate(linhas) if "FECHA EMISION" in ln.upper() or "FECHA EMISIÓN" in ln.upper()]
        if idxs:
            idx = idxs[1] if len(idxs) >= 2 else idxs[0]
            if idx >= 2:
                return linhas[idx - 2].strip()
            elif idx > 0:
                return linhas[idx - 1].strip()
            return linhas[idx].strip()
    elif fornecedor == "60037433" and len(linhas) >= 5:
        return linhas[4].strip()
    elif fornecedor == "10001021":
        idxs = [i for i, ln in enumerate(linhas) if "FECHA EMISION" in ln.upper() or "FECHA EMISIÓN" in ln.upper()]
        if idxs:
            idx = idxs[1] if len(idxs) >= 2 else idxs[0]
            for off in (3, 2, 1, 0):
                if idx - off >= 0:
                    return linhas[idx - off].strip()
    elif fornecedor == "51092775" and len(linhas) >= 11:
        return linhas[10].strip()
    elif fornecedor == "34764689" and len(linhas) >= 1:
        return linhas[0].strip()
    elif fornecedor == "WAN HAI" and len(linhas) >= 1:
        return linhas[0].strip()
    elif fornecedor == "EVERGREEN":
        for i, ln in enumerate(linhas):
            if "FECHA EMISION" in ln.upper() or "FECHA EMISIÓN" in ln.upper():
                if i >= 1:
                    return linhas[i - 1].strip()
    elif fornecedor == "MSC" and len(linhas) >= 1:
        return linhas[0].strip()
    elif fornecedor == "61092558" and len(linhas) >= 3:
        return linhas[2].strip()
    elif fornecedor == "54308388" and len(linhas) >= 7:
        return linhas[6].strip()
    return ""

def padronizar_tipo_doc(df: pd.DataFrame) -> pd.DataFrame:
    subs = {
        "FACTURA ELECTRÓNICA": "FACTURA",
        "FACTURA ELECTRONICA": "FACTURA",
        "FACTURA  ELECTRÓNICA": "FACTURA",
        "ELECTRONIC INVOICE": "FACTURA",
        "INVOICE": "FACTURA",
        "NOTA DE CRÉDITO ELECTRÓNICA": "NOTA DE CRÉDITO",
        "NOTA DE CREDITO": "NOTA DE CRÉDITO",
        "NOTA DE CRÉDITO": "NOTA DE CRÉDITO",
        "Factura": "FACTURA",
    }
    df["Tipo Doc"] = df["Tipo Doc"].replace(subs)
    return df


# --- OUTRAS REGRAS ---

def Ajustar_nro_nota_credito(df: pd.DataFrame) -> pd.DataFrame:
    def limpar_valor(v: str) -> str:
        return (v.replace("Nro", "").replace("N°", "").replace(".", "").replace(" ", "").strip())

    def get_factura(row):
        proveedor = str(row['Proveedor Iscala']).strip()
        tipo_doc = str(row.get('Tipo Doc', '')).strip().upper()
        linhas_pdf = row['conteudo_pdf'].splitlines()

        if tipo_doc == 'NOTA DE CRÉDITO':
            if proveedor == '10001013' and len(linhas_pdf) > 1:
                return limpar_valor(linhas_pdf[1])
            elif proveedor == '25206207' and len(linhas_pdf) > 0:
                return limpar_valor(linhas_pdf[0])
            elif proveedor == '10001021':
                for i, linha in enumerate(linhas_pdf):
                    if 'NOTA DE CREDITO' in linha.upper() and i > 0:
                        return limpar_valor(linhas_pdf[i - 1])
            elif proveedor == '61092558' and len(linhas_pdf) > 2:
                return limpar_valor(linhas_pdf[1])

        return limpar_valor(str(row.get('Factura', '')))

    df['Factura'] = df.apply(get_factura, axis=1)
    return df

def atribuir_cuenta(cod_moneda: str) -> str:
    if cod_moneda == "01":
        return "421202"
    elif cod_moneda == "00":
        return "421201"
    return ""

def error(valor: str) -> str:
    return "Can't read the file" if valor == "" else ""

# Reaproveita a lógica de merge com Tasa do DUAS
from services.duas_utils import adicionar_coluna_tasa  # [2](https://electrolux-my.sharepoint.com/personal/thiago_farias_electrolux_com/Documents/Microsoft%20Copilot%20Chat%20Files/process_pdfs.py)

def adicionar_cod_autorizacion_adicionales(df: pd.DataFrame) -> pd.DataFrame:
    if 'Tipo Doc' in df.columns:
        df['Tipo de Factura'] = df['Tipo Doc'].apply(
            lambda x: "01" if str(x).strip().upper() in ("FACTURA", "NOTA DE CRÉDITO", "NOTA DE CREDITO") else None
        )
    return df

def adicionar_tip_doc_adicionales(df: pd.DataFrame) -> pd.DataFrame:
    if 'Tipo Doc' in df.columns:
        df['Cód. de Autorización'] = df['Tipo Doc'].apply(
            lambda x: "01" if str(x).strip().upper() == "FACTURA"
            else "07" if str(x).strip().upper() in ("NOTA DE CRÉDITO", "NOTA DE CREDITO")
            else None
        )
    return df

def organizar_colunas_adicionales(df: pd.DataFrame) -> pd.DataFrame:
    desejadas = [
        'source_file', 'conteudo_pdf', 'R.U.C', 'Proveedor Iscala', 'Factura', 'Tipo Doc',
        'Cód. de Autorización', 'Tipo de Factura', 'Fecha de Emisión', 'Moneda',
        'Cod. Moneda', 'Op. Gravada', 'Tasa', 'Cuenta', 'Error'
    ]
    presentes = [c for c in desejadas if c in df.columns]
    return df[presentes + [c for c in df.columns if c not in presentes]]

def remover_duplicatas_source_file(df: pd.DataFrame) -> pd.DataFrame:
    return df.drop_duplicates(subset='source_file', keep='first')
