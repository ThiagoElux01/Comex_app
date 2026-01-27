
# services/externos_utils.py  (ACRÉSCIMO DE IMPORTS NO TOPO)
import re
import pandas as pd
from datetime import datetime

def identificar_Proveedor(df):
    # Ordem de prioridade: os mais específicos primeiro
    fornecedores = [
        "Electrolux Intressenter AB",
        "Electrolux S.E.A. Pte",
        "ELECTROLUX DE CHILE",
        "HOMA APPLIANCES CO",
        "ELECTROLUX HOME PRODUCTS",
        "MIDEA ELECTRIC TRADING",
        "NINGBO XINLE HOUSEHOLD APPLIANCES CO",
        "ELECTROLUX DO BRASIL",
        "GUANGDONG GALANZ",
        "NINGBO HUACAI ELECTRIC APPLIANCES CO",
        "Hefei Snowky Electric",
        "Trade Air System",
        "JIANGMEN JINHUAN",
        "FOSHAN SHUNDE MIDEA"
    ]

    def selecionar_Proveedor(texto):
        texto = str(texto).upper()
        for fornecedor in fornecedores:
            if fornecedor.upper() in texto:
                return fornecedor
        return ""

    df['Proveedor'] = df['conteudo_pdf'].apply(selecionar_Proveedor)
    return df

def adicionar_provedor_iscala(df):
    depara = {
        "Electrolux Intressenter AB": "SEI",
        "HOMA APPLIANCES CO": "5BE",
        "ELECTROLUX HOME PRODUCTS": "US1239",
        "ELECTROLUX DE CHILE": "CLH",
        "MIDEA ELECTRIC TRADING": "5WY",
        "NINGBO XINLE HOUSEHOLD APPLIANCES CO": "5DL",
        "ELECTROLUX DO BRASIL": "BRR",
        "GUANGDONG GALANZ": "5DU",
        "NINGBO HUACAI ELECTRIC APPLIANCES CO": "NINGBO HUA",
        "Electrolux S.E.A. Pte": "SGE",
        "Hefei Snowky Electric": "SNOWKY",
        "Trade Air System": "7DQ",
        "JIANGMEN JINHUAN": "5JU",
        "FOSHAN SHUNDE MIDEA": "7JR"
    }

    def mapear_provedor(provedor):
        for chave in depara:
            if chave in str(provedor):
                return depara[chave]
        return ""

    df['Proveedor Iscala'] = df['Proveedor'].apply(mapear_provedor)
    return df

def extrair_factura(df):
    def buscar_factura(row):
        texto = row['conteudo_pdf']
        provedor = row['Proveedor Iscala']
        linhas = texto.splitlines()

        try:
            if provedor == "SEI":
                for i, linha in enumerate(linhas):
                    linha_upper = linha.upper()
                    if "INVOICE DATE" in linha_upper:
                        return linhas[i - 2].strip() if i >= 2 else ""
                    elif "CREDIT NOTE DATE" in linha_upper:
                        if i >= 2:
                            linha_alvo = linhas[i - 2]
                            numeros = re.findall(r'\d+', linha_alvo)
                            return ' '.join(numeros) if numeros else linha_alvo.strip()
                    elif "DOCUMENT DATE:" in linha_upper:
                        return linhas[i - 1].strip() if i >= 1 else ""

            elif provedor == "SNOWKY":
                for i, linha in enumerate(linhas):
                    if "INVOICE NO." in linha:
                        return linhas[i + 1] if i + 1 < len(linhas) else ""

            elif provedor == "5BE":
                for linha in linhas:
                    if "INVOICE NO." in linha:
                        partes = linha.split("INVOICE NO.")
                        return partes[1].strip() if len(partes) > 1 else linha.strip()


            elif provedor == "US1239":
                for i, linha in enumerate(linhas):
                    if "RUC:" in linha.upper() and i + 2 < len(linhas):
                        linha_alvo = linhas[i + 2].strip()
                        match = re.search(r'\bEH\d{8}\b', linha_alvo)
                        if match:
                            return match.group(0)

                for i, linha in enumerate(linhas):
                    if "INVOICE AND PACKING LIST" in linha.upper():
                        if i + 1 < len(linhas):
                            return linhas[i + 1].strip()

                for i, linha in enumerate(linhas):
                    if "ELECTROLUX HOME PRODUCTS INTERNATIONAL" in linha.upper():
                        if i - 1 >= 0:
                            return linhas[i - 1].strip()

                return ""  # Caso nenhuma das lógicas encontre algo


            elif provedor == "CLH":
                for i, linha in enumerate(linhas):
                    if "ELECTRONIC EXPORT INVOICE" in linha.upper():
                        return linhas[i + 2] if i + 2 < len(linhas) else ""
                    elif "NOTA DE CRÉDITO" in linha.upper():
                        return linhas[i + 5] if i + 5 < len(linhas) else ""

            elif provedor == "7DQ":
                linha_customer = ""
                for i, linha in enumerate(linhas):
                    if "Customer Number" in linha:
                        linha_customer = linhas[i + 1] if i + 1 < len(linhas) else ""
                        break
                linha_quinta = linhas[4] if len(linhas) > 4 else ""
                linha_primeira = linhas[0] if linhas else ""
                return f"{linha_customer} {linha_quinta} {linha_primeira}".strip()

            elif provedor == "5JU":
                for i, linha in enumerate(linhas):
                    if "Invoice #.:" in linha:
                        return linhas[i + 1] if i + 1 < len(linhas) else ""
                    elif "CN No.:" in linha:
                        partes = linha.split("CN No.:")
                        if len(partes) > 1:
                            return partes[1].strip()

            elif provedor == "5WY":
                for linha in linhas:
                    if "MDOK" in linha or "MDR" in linha:
                        return linha.strip()

            elif provedor == "7JR":
                for linha in linhas:
                    if "MD" in linha:
                        match = re.search(r"(MD.*)", linha)
                        if match:
                            return match.group(1).strip()
                        return linha.strip()
                    
            elif provedor == "5DL":
                for i, linha in enumerate(linhas):
                    if "Invoice No." in linha:
                        return linhas[i + 3] if i + 3 < len(linhas) else ""

            elif provedor == "BRR":
                for i, linha in enumerate(linhas):
                    if "FACTURA COMERCIAL" in linha.upper() or "CREDIT NOTE" in linha.upper():
                        partes = re.split(r"FACTURA COMERCIAL|CREDIT NOTE", linha, flags=re.IGNORECASE)
                        resultado = partes[1].strip() if len(partes) > 1 else ""
                        if resultado:
                            return resultado
                        elif i + 1 < len(linhas):
                            return linhas[i + 1].strip()


            elif provedor == "5DU":
                for linha in linhas:
                    if "INV. NO:" in linha:
                        partes = linha.split("INV. NO:")
                        return partes[1].strip() if len(partes) > 1 else linha.strip()

            elif provedor == "NINGBO HUA":
                for i, linha in enumerate(linhas):
                    linha_upper = linha.upper()
                    if "INVOICE NO" in linha_upper:
                        partes = linha_upper.split("INVOICE NO")
                        return partes[1].strip() if len(partes) > 1 else linha.strip()
                    elif "CREDIT NOTE" in linha_upper:
                        indice_desejado = i + 6
                        if indice_desejado < len(linhas):
                            return linhas[indice_desejado].strip()

            elif provedor == "SGE":
                for i, linha in enumerate(linhas):
                    if "Invoice Date" in linha:
                        return linhas[i - 2].strip() if i >= 2 else ""
                    elif "Credit Note Date" in linha:
                        if i >= 2:
                            linha_alvo = linhas[i - 2]
                            numeros = re.findall(r'\d+', linha_alvo)
                            return ' '.join(numeros) if numeros else linha_alvo.strip()

        except Exception as e:
            return f"[Erro ao extrair: {e}]"

        return ""

    df['Factura'] = df.apply(buscar_factura, axis=1)
    return df

def ajustar_factura(df):
    import re

    def limpar(texto):
        if not isinstance(texto, str):
            return texto
        texto = texto.upper()
        texto = re.sub(r'\b(Nº|NO|N°|Nº\.|N°\.|Nº:|NO:)\b', '', texto)  # Remove prefixos
        texto = texto.replace(":", "")
        texto = texto.replace("：", "")
        texto = texto.replace(".", "")
        texto = texto.replace(",", "")
        texto = texto.replace(" ", "")
        return texto.strip()

    df['Factura'] = df['Factura'].apply(limpar)
    return df

def extrair_fecha(df):
    def extrair_data(row):
        provedor = row.get("Proveedor Iscala", "").strip().upper()
        linhas = row.get("conteudo_pdf", "").splitlines()

        # 🔹 Regra 1: SEI
        if provedor == "SEI":
            for i, linha in enumerate(linhas):
                linha_upper = linha.upper()
                if ("INVOICE DATE" in linha_upper or "CREDIT NOTE DATE" in linha_upper or "DOCUMENT DATE:" in linha_upper) and i + 1 < len(linhas):
                    data_bruta = linhas[i + 1].strip()
                    for fmt in ("%d.%m.%Y", "%Y.%m.%d"):
                        try:
                            return datetime.strptime(data_bruta, fmt).strftime("%d%m%y")
                        except ValueError:
                            continue
                    return data_bruta

        # 🔹 Regra 2: SNOWKY
        elif provedor == "SNOWKY":
            for i, linha in enumerate(linhas):
                if "DATE:" in linha.upper() and i + 1 < len(linhas):
                    data_bruta = linhas[i + 1].strip()
                    for fmt in ("%d.%m.%Y", "%Y.%m.%d"):
                        try:
                            return datetime.strptime(data_bruta, fmt).strftime("%d%m%y")
                        except ValueError:
                            continue
                    return data_bruta

        # 🔹 Regra ajustada: 5BE
        
        elif provedor == "5BE":
            # Traduções de meses PT/ES -> EN (abreviado)
            meses_traducao = {
                "jan": "Jan", "ene": "Jan",
                "fev": "Feb", "feb": "Feb",
                "mar": "Mar",
                "abr": "Apr",
                "mai": "May", "may": "May",
                "jun": "Jun",
                "jul": "Jul",
                "ago": "Aug",
                "set": "Sep", "sep": "Sep",
                "out": "Oct", "oct": "Oct",
                "nov": "Nov",
                "dez": "Dec", "dic": "Dec"
            }

            def tenta_parse(s: str):
                """Tenta converter uma string de data em 'ddmmyy', cobrindo mês textual e formatos numéricos."""
                s = s.strip()
                # Normaliza meses por extenso (se houver)
                for pt, en in meses_traducao.items():
                    s = re.sub(rf"(?i)\b{pt}\b", en, s)

                # Formatos candidatos (mês textual, numéricos e ISO)
                formatos = (
                    # mês textual abreviado (Jan/Feb…)
                    "%d-%b-%y", "%d/%b/%y",
                    "%d-%b-%Y", "%d/%b/%Y",

                    # totalmente numéricos (dia primeiro)
                    "%d/%m/%Y", "%d-%m-%Y",
                    "%d/%m/%y",  "%d-%m-%y",

                    # ISO (ano primeiro)
                    "%Y-%m-%d", "%Y/%m/%d",
                )
                # 1) tentativa direta
                for fmt in formatos:
                    try:
                        return datetime.strptime(s, fmt).strftime("%d%m%y")
                    except Exception:
                        pass

                # 2) fallback por regex: isola a "parte de data" e tenta novamente
                padroes = [
                    r"\b\d{4}-\d{2}-\d{2}\b",      # ISO YYYY-MM-DD
                    r"\b\d{4}/\d{2}/\d{2}\b",      # ISO YYYY/MM/DD
                    r"\b\d{1,2}/\d{1,2}/\d{2,4}\b",
                    r"\b\d{1,2}-\d{1,2}-\d{2,4}\b",
                    r"\b\d{1,2}/[A-Za-z]{3}/\d{2,4}\b",
                    r"\b\d{1,2}-[A-Za-z]{3}-\d{2,4}\b",
                ]
                for pat in padroes:
                    m = re.search(pat, s)
                    if m:
                        s2 = m.group(0).strip()
                        for fmt in formatos:
                            try:
                                return datetime.strptime(s2, fmt).strftime("%d%m%y")
                            except Exception:
                                pass
                return None  # não conseguiu interpretar

            # Procurar por "DATE" (com ou sem dois pontos) e aceitar valor na mesma linha ou na linha de baixo
            linhas_upper = [ln.upper() for ln in linhas]
            for i, ln_up in enumerate(linhas_upper):
                if "DATE" in ln_up:  # cobre "DATE", "DATE:", "** DATE**", etc.
                    # 1) Tenta mesma linha (após a palavra DATE opcionalmente seguida de ':')
                    original_line = linhas[i]
                    partes = re.split(r'(?i)DATE\s*:?', original_line, maxsplit=1)
                    candidato_mesma = partes[1].strip() if len(partes) > 1 else ""

                    # 2) Se vazio, tenta a linha imediatamente seguinte (como em vários 5BE)
                    candidato_proxima = linhas[i + 1].strip() if (not candidato_mesma and i + 1 < len(linhas)) else ""

                    # 3) Tenta parsear os candidatos
                    for cand in (candidato_mesma, candidato_proxima):
                        if cand:
                            parsed = tenta_parse(cand)
                            if parsed:
                                return parsed

            # Último recurso: varre todo o documento e devolve a primeira data válida encontrada
            for ln in linhas:
                parsed = tenta_parse(ln)
                if parsed:
                    return parsed

            return ""  # mantém compatibilidade com o restante do pipeline


        elif provedor == "US1239":
            encontrou_ref_claim = any("REF CLAIM" in linha.upper() for linha in linhas)

            def extrair_data_formatada(texto):
                match = re.search(r'\b(\d{1,2}/\d{1,2}/\d{2})\b', texto)
                if match:
                    try:
                        return datetime.strptime(match.group(1), "%m/%d/%y").strftime("%d%m%y")
                    except ValueError:
                        return None
                return None

            if encontrou_ref_claim:
                for linha in linhas:
                    data = extrair_data_formatada(linha)
                    if data:
                        return data
            else:
                for linha in linhas:
                    if "ELECTROLUX HOME PRODUCTS" in linha.upper():
                        data = extrair_data_formatada(linha)
                        if data:
                            return data

            # Se nenhuma das opções anteriores retornar uma data válida, busca a primeira data no documento
            for linha in linhas:
                data = extrair_data_formatada(linha)
                if data:
                    return data


        # 🔹 Regra 5: CLH
        elif provedor == "CLH":
            padrao_data = re.compile(r'\d{1,2} de [A-Za-zçÇñÑ]{3,15} de \d{4}')
            meses_traducao = {
                # Português abreviado e por extenso
                "jan": "Jan", "janeiro": "Jan",
                "fev": "Feb", "fevereiro": "Feb",
                "mar": "Mar", "março": "Mar",
                "abr": "Apr", "abril": "Apr",
                "mai": "May", "maio": "May",
                "jun": "Jun", "junho": "Jun",
                "jul": "Jul", "julho": "Jul",
                "ago": "Aug", "agosto": "Aug",
                "set": "Sep", "setembro": "Sep",
                "out": "Oct", "outubro": "Oct",
                "nov": "Nov", "novembro": "Nov",
                "dez": "Dec", "dezembro": "Dec",
                # Espanhol abreviado e por extenso
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
                "dic": "Dec", "diciembre": "Dec"
            }
            for linha in linhas:
                match = padrao_data.search(linha.lower())
                if match:
                    data_bruta = match.group(0)
                    for mes_local, mes_en in meses_traducao.items():
                        if f" de {mes_local} de " in data_bruta:
                            data_convertida = data_bruta.replace(f" de {mes_local} de ", f" de {mes_en} de ")
                            try:
                                return datetime.strptime(data_convertida, "%d de %b de %Y").strftime("%d%m%y")
                            except ValueError:
                                return data_bruta
                    return data_bruta


        # 🔹 Regra 6: 7DQ
        elif provedor == "7DQ":
            padrao_data = re.compile(r'\b\d{2}/\d{2}/\d{4}\b')
            for linha in linhas:
                match = padrao_data.search(linha)
                if match:
                    try:
                        return datetime.strptime(match.group(0), "%d/%m/%Y").strftime("%d%m%y")
                    except ValueError:
                        return match.group(0)
        # 🔹 Regra 7: 5JU
        elif provedor == "5JU":
            meses_traducao = {
                # Português abreviado e por extenso
                "jan": "Jan", "janeiro": "Jan",
                "fev": "Feb", "fevereiro": "Feb",
                "mar": "Mar", "março": "Mar",
                "abr": "Apr", "abril": "Apr",
                "mai": "May", "maio": "May",
                "jun": "Jun", "junho": "Jun",
                "jul": "Jul", "julho": "Jul",
                "ago": "Aug", "agosto": "Aug",
                "set": "Sep", "setembro": "Sep",
                "out": "Oct", "outubro": "Oct",
                "nov": "Nov", "novembro": "Nov",
                "dez": "Dec", "dezembro": "Dec",
                # Espanhol abreviado e por extenso
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
                "dic": "Dec", "diciembre": "Dec"
            }

            if any("CN NO" in linha.upper() for linha in linhas):
                for linha in linhas:
                    if "DATE:" in linha.upper():
                        partes = linha.upper().split("DATE:")
                        if len(partes) > 1:
                            data_bruta = partes[1].strip()
                            partes_data = re.split(r"[-/]", data_bruta)
                            if len(partes_data) == 3:
                                dia, mes, ano = partes_data
                                mes_lower = mes.lower()
                                mes_en = meses_traducao.get(mes_lower, mes.capitalize())
                                data_convertida = f"{dia}-{mes_en}-{ano}"
                                try:
                                    return datetime.strptime(data_convertida, "%d-%b-%Y").strftime("%d%m%y")
                                except ValueError:
                                    return data_bruta
            else:
                padrao_data = re.compile(r'\b\d{1,2}-[A-Za-zçÇñÑ]{3,9}-\d{2}\b')
                for i, linha in enumerate(linhas):
                    if "DESCRIPTION" in linha.upper() and i > 0:
                        linha_anterior = linhas[i - 1]
                        match = padrao_data.search(linha_anterior)
                        if match:
                            data_bruta = match.group(0)
                            partes = data_bruta.split("-")
                            if len(partes) == 3:
                                dia, mes, ano = partes
                                mes_lower = mes.lower()
                                if mes_lower in meses_traducao:
                                    mes_en = meses_traducao[mes_lower]
                                    data_convertida = f"{dia}-{mes_en}-{ano}"
                                    try:
                                        return datetime.strptime(data_convertida, "%d-%b-%y").strftime("%d%m%y")
                                    except ValueError:
                                        return data_bruta
                            return data_bruta


        # 🔹 Regra 8: 5WY
        elif provedor == "5WY":
            padrao_data = re.compile(r'\b\d{1,2}/[A-Za-z]{3}/\d{4}\b|\b\d{1,2}/\d{1,2}/\d{4}\b|\b\d{4}-\d{2}-\d{2}\b')
            meses_traducao = {
                "jan": "Jan", "ene": "Jan",
                "fev": "Feb", "feb": "Feb",
                "mar": "Mar",
                "abr": "Apr",
                "mai": "May", "may": "May",
                "jun": "Jun",
                "jul": "Jul",
                "ago": "Aug",
                "set": "Sep", "sep": "Sep",
                "out": "Oct", "oct": "Oct",
                "nov": "Nov",
                "dez": "Dec", "dic": "Dec"
            }
            for linha in linhas:
                match = padrao_data.search(linha)
                if match:
                    data_bruta = match.group(0)
                    if re.search(r'\d{4}-\d{2}-\d{2}', data_bruta):
                        try:
                            return datetime.strptime(data_bruta, "%Y-%m-%d").strftime("%d%m%y")
                        except ValueError:
                            return data_bruta
                    elif re.search(r'\d{1,2}/[A-Za-z]{3}/\d{4}', data_bruta):
                        partes = data_bruta.split("/")
                        dia, mes, ano = partes
                        mes_lower = mes.lower()
                        mes_en = meses_traducao.get(mes_lower, mes)
                        data_convertida = f"{dia}/{mes_en}/{ano}"
                        try:
                            return datetime.strptime(data_convertida, "%d/%b/%Y").strftime("%d%m%y")
                        except ValueError:
                            return data_bruta
                    else:
                        try:
                            return datetime.strptime(data_bruta, "%d/%m/%Y").strftime("%d%m%y")
                        except ValueError:
                            return data_bruta


        # 🔹 Regra 9: 7JR
        elif provedor == "7JR":
            padrao_data = re.compile(r'\b\d{1,2}/[A-Za-z]{3}/\d{4}\b|\b\d{1,2}/\d{1,2}/\d{4}\b')
            meses_traducao = {
                # Português e Espanhol abreviado
                "jan": "Jan", "ene": "Jan",
                "fev": "Feb", "feb": "Feb",
                "mar": "Mar",
                "abr": "Apr",
                "mai": "May", "may": "May",
                "jun": "Jun",
                "jul": "Jul",
                "ago": "Aug",
                "set": "Sep", "sep": "Sep",
                "out": "Oct", "oct": "Oct",
                "nov": "Nov",
                "dez": "Dec", "dic": "Dec"
            }
            for linha in linhas:
                match = padrao_data.search(linha)
                if match:
                    data_bruta = match.group(0)
                    # Verifica se o mês é texto (ex: Jan)
                    if re.search(r'\d{1,2}/[A-Za-z]{3}/\d{4}', data_bruta):
                        partes = data_bruta.split("/")
                        dia, mes, ano = partes
                        mes_lower = mes.lower()
                        mes_en = meses_traducao.get(mes_lower, mes)
                        data_convertida = f"{dia}/{mes_en}/{ano}"
                        try:
                            return datetime.strptime(data_convertida, "%d/%b/%Y").strftime("%d%m%y")
                        except ValueError:
                            return data_bruta
                    else:
                        try:
                            return datetime.strptime(data_bruta, "%d/%m/%Y").strftime("%d%m%y")
                        except ValueError:
                            return data_bruta

        # 🔹 Regra 10: 5DL
        elif provedor == "5DL":
            meses_traducao = {
                # Português abreviado e por extenso
                "jan": "Jan", "janeiro": "Jan",
                "fev": "Feb", "fevereiro": "Feb",
                "mar": "Mar", "março": "Mar",
                "abr": "Apr", "abril": "Apr",
                "mai": "May", "maio": "May",
                "jun": "Jun", "junho": "Jun",
                "jul": "Jul", "julho": "Jul",
                "ago": "Aug", "agosto": "Aug",
                "set": "Sep", "setembro": "Sep",
                "out": "Oct", "outubro": "Oct",
                "nov": "Nov", "novembro": "Nov",
                "dez": "Dec", "dezembro": "Dec",
                # Espanhol abreviado e por extenso
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
                "dic": "Dec", "diciembre": "Dec"
            }
            padrao_data = re.compile(r'\b\d{1,2}-[A-Za-zçÇñÑ]{3,9}-\d{2}\b')
            for linha in linhas:
                match = padrao_data.search(linha)
                if match:
                    data_bruta = match.group(0)
                    partes = data_bruta.split("-")
                    if len(partes) == 3:
                        dia, mes, ano = partes
                        mes_lower = mes.lower()
                        if mes_lower in meses_traducao:
                            mes_en = meses_traducao[mes_lower]
                            data_convertida = f"{dia}-{mes_en}-{ano}"
                            try:
                                return datetime.strptime(data_convertida, "%d-%b-%y").strftime("%d%m%y")
                            except ValueError:
                                return data_bruta
                    return data_bruta

        elif provedor == "BRR":
            meses_traducao = {
                # Português abreviado e por extenso
                "jan": "Jan", "janeiro": "Jan",
                "fev": "Feb", "fevereiro": "Feb",
                "mar": "Mar", "março": "Mar",
                "abr": "Apr", "abril": "Apr",
                "mai": "May", "maio": "May",
                "jun": "Jun", "junho": "Jun",
                "jul": "Jul", "julho": "Jul",
                "ago": "Aug", "agosto": "Aug",
                "set": "Sep", "setembro": "Sep",
                "out": "Oct", "outubro": "Oct",
                "nov": "Nov", "novembro": "Nov",
                "dez": "Dec", "dezembro": "Dec",
                # Espanhol abreviado e por extenso
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
                "dic": "Dec", "diciembre": "Dec"
            }

            if any("CREDIT NOTE" in linha.upper() for linha in linhas):
                padrao_ddmmyyyy = re.compile(r'\b\d{2}/\d{2}/\d{4}\b')
                for linha in linhas:
                    match = padrao_ddmmyyyy.search(linha)
                    if match:
                        try:
                            return datetime.strptime(match.group(0), "%d/%m/%Y").strftime("%d%m%y")
                        except ValueError:
                            return match.group(0)
            else:
                padrao_data = re.compile(r'\b\d{1,2}/[A-Za-zñÑ]{3,15}/\d{4}\b')
                for i, linha in enumerate(linhas):
                    if "FECHA" in linha.upper() and i + 1 < len(linhas):
                        linha_abaixo = linhas[i + 1]
                        match = padrao_data.search(linha_abaixo)
                        if match:
                            data_bruta = match.group(0)
                            partes = data_bruta.split("/")
                            if len(partes) == 3:
                                dia, mes, ano = partes
                                mes_lower = mes.lower()
                                if mes_lower in meses_traducao:
                                    mes_en = meses_traducao[mes_lower]
                                    data_convertida = f"{dia}/{mes_en}/{ano}"
                                    try:
                                        return datetime.strptime(data_convertida, "%d/%b/%Y").strftime("%d%m%y")
                                    except ValueError:
                                        return data_bruta
                            return data_bruta


        # 🔹 Regra 12: 5DU
        elif provedor == "5DU":
            meses_traducao = {
                # Traduções de abreviações com ponto
                "jan.": "Jan", "feb.": "Feb", "mar.": "Mar", "apr.": "Apr",
                "may.": "May", "jun.": "Jun", "jul.": "Jul", "aug.": "Aug",
                "sep.": "Sep", "oct.": "Oct", "nov.": "Nov", "dec.": "Dec"
            }
            padrao_data = re.compile(r'\b[A-Za-z]{3,4}\.\d{1,2},\d{4}\b')
            for linha in linhas:
                if "DATE:" in linha.upper():
                    match = padrao_data.search(linha)
                    if match:
                        data_bruta = match.group(0)
                        partes = re.split(r'[.,]', data_bruta)
                        if len(partes) == 3:
                            mes, dia, ano = partes
                            mes_lower = mes.lower() + "."
                            mes_en = meses_traducao.get(mes_lower, mes.capitalize())
                            data_convertida = f"{dia.zfill(2)}-{mes_en}-{ano}"
                            try:
                                return datetime.strptime(data_convertida, "%d-%b-%Y").strftime("%d%m%y")
                            except ValueError:
                                return data_bruta
                    return ""
                
        # 🔹 Regra 13: NINGBO HUA
        elif provedor == "NINGBO HUA":
            padrao_iso = re.compile(r'\b\d{4}-\d{2}-\d{2}\b')
            padrao_textual = re.compile(r'\b(\d{1,2})(?:st|nd|rd|th)?\s*,?\s*(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})\b', re.IGNORECASE)

            for linha in linhas:
                match_iso = padrao_iso.search(linha)
                if match_iso:
                    try:
                        return datetime.strptime(match_iso.group(0), "%Y-%m-%d").strftime("%d%m%y")
                    except ValueError:
                        return match_iso.group(0)

                match_textual = padrao_textual.search(linha)
                if match_textual:
                    dia, mes, ano = match_textual.groups()
                    try:
                        data = datetime.strptime(f"{dia} {mes} {ano}", "%d %B %Y")
                        return data.strftime("%d%m%y")
                    except ValueError:
                        return f"{dia} {mes} {ano}"

                    
        # 🔹 Regra 14: SGE
        elif provedor == "SGE":
            padrao_data = re.compile(r'\b\d{2}\.\d{2}\.\d{4}\b')
            for i, linha in enumerate(linhas):
                linha_upper = linha.upper()
                if ("INVOICE DATE" in linha_upper or "CREDIT NOTE DATE" in linha_upper) and i + 1 < len(linhas):
                    linha_abaixo = linhas[i + 1]
                    match = padrao_data.search(linha_abaixo)
                    if match:
                        try:
                            return datetime.strptime(match.group(0), "%d.%m.%Y").strftime("%d%m%y")
                        except ValueError:
                            return match.group(0)

                    
        return ""
    df["Fecha de Emisión"] = df.apply(extrair_data, axis=1)
    return df

def ajustar_coluna_fecha(df):

    def converter_data(data_str):
        if pd.isna(data_str) or len(data_str) != 6:
            return data_str
        try:
            dia = data_str[:2]
            mes = data_str[2:4]
            ano = data_str[4:]
            ano_completo = '20' + ano if int(ano) < 50 else '19' + ano
            return f"{dia}/{mes}/{ano_completo}"
        except Exception:
            return data_str

    df['Fecha de Emisión'] = df['Fecha de Emisión'].apply(converter_data)
    return df

def adicionar_tipo_doc(df):
    def classificar_tipo_doc(row):
        provedor = row.get("Proveedor Iscala", "").strip().upper()
        linhas = row.get("conteudo_pdf", "").splitlines()
        
        if provedor == "SEI":
            return linhas[0].strip().upper() if linhas else ""
        
        elif provedor == "SGE":
            return linhas[0].strip() if linhas else ""
        
        elif provedor == "5BE":
            linhas_upper = [linha.upper() for linha in linhas]
            if any("CREDIT NOTE" in linha for linha in linhas_upper):
                return "CREDIT NOTE"
            else:
                return "INVOICE"
        
        elif provedor == "5DL":
            linhas_upper = [linha.upper() for linha in linhas]
            if any("CREDIT NOTE" in linha for linha in linhas_upper):
                return "CREDIT NOTE"
            else:
                return "INVOICE"

        elif provedor == "5DU":
            linhas_upper = [linha.upper() for linha in linhas]
            if any("CREDIT NOTE" in linha for linha in linhas_upper):
                return "CREDIT NOTE"
            else:
                return "INVOICE"
            
        elif provedor == "5JU":
            linhas_upper = [linha.upper() for linha in linhas]
            if any("CREDIT NOTE" in linha for linha in linhas_upper):
                return "CREDIT NOTE"
            else:
                return "INVOICE"

        elif provedor == "5WY":
            linhas_upper = [linha.upper() for linha in linhas]
            if any("CREDIT NOTE" in linha for linha in linhas_upper):
                return "CREDIT NOTE"
            else:
                return "INVOICE"        

        elif provedor == "7DQ":
            linhas_upper = [linha.upper() for linha in linhas]
            if any("CREDIT NOTE" in linha for linha in linhas_upper):
                return "CREDIT NOTE"
            else:
                return "INVOICE"         

        elif provedor == "7JR":
            linhas_upper = [linha.upper() for linha in linhas]
            if any("CREDIT NOTE" in linha for linha in linhas_upper):
                return "CREDIT NOTE"
            else:
                return "INVOICE" 

        elif provedor == "BRR":
            linhas_upper = [linha.upper() for linha in linhas]
            if any("CREDIT NOTE" in linha for linha in linhas_upper):
                return "CREDIT NOTE"
            else:
                return "INVOICE" 

        elif provedor == "BRR":
            linhas_upper = [linha.upper() for linha in linhas]
            if any("CREDIT NOTE" in linha for linha in linhas_upper):
                return "CREDIT NOTE"
            else:
                return "INVOICE" 

        elif provedor == "CLH":
            linhas_upper = [linha.upper() for linha in linhas]
            if any("NOTA DE CRÉDITO" in linha for linha in linhas_upper):
                return "CREDIT NOTE"
            else:
                return "INVOICE" 

        elif provedor == "NINGBO HUA":
            linhas_upper = [linha.upper() for linha in linhas]
            if any("CREDIT NOTE" in linha for linha in linhas_upper):
                return "CREDIT NOTE"
            else:
                return "INVOICE" 

        elif provedor == "SNOWKY":
            linhas_upper = [linha.upper() for linha in linhas]
            if any("CREDIT NOTE" in linha for linha in linhas_upper):
                return "CREDIT NOTE"
            else:
                return "INVOICE" 

        elif provedor == "US1239":
            linhas_upper = [linha.upper() for linha in linhas]
            if any("REF CLAIM" in linha for linha in linhas_upper):
                return "CREDIT NOTE"
            else:
                return "INVOICE"                                                          
        return ""
    df["Tipo Doc"] = df.apply(classificar_tipo_doc, axis=1)
    return df

def adicionar_amount(df):
    def extrair_amount(row):
        provedor = row.get("Proveedor Iscala", "").strip().upper()
        linhas = row.get("conteudo_pdf", "").splitlines()

        if provedor == "SEI":
            for i, linha in enumerate(linhas):
                if "TOTAL AMOUNT(U.S DOLLAR)" in linha.upper():
                    return linhas[i + 1].strip() if i + 1 < len(linhas) else ""
           
            # Caso não encontre "TOTAL AMOUNT(U.S DOLLAR)", procurar "Total Net in Doc. Currency"
            for i, linha in enumerate(linhas):
                if "TOTAL NET IN DOC. CURRENCY" in linha.upper():
                    return linhas[i + 1].strip() if i + 1 < len(linhas) else ""

        if provedor == "SGE":
            for i, linha in enumerate(linhas):
                if "TOTAL AMOUNT(U.S DOLLAR)" in linha.upper():
                    return linhas[i + 1].strip() if i + 1 < len(linhas) else ""
                       

        if provedor == "SNOWKY":
            for i, linha in enumerate(linhas):
                if "QUANTITIES & DESCRIPTIONS" in linha.upper():
                    for offset in range(1, 11):
                        idx = i - offset
                        if idx >= 0:
                            linha_acima = linhas[idx].strip()
                            if "US$" in linha_acima.upper():
                                return linha_acima
                    return ""

        if provedor == "5BE":
            for i, linha in enumerate(linhas):
                if "SHIPPING MARKS:" in linha.upper():
                    return linhas[i - 1].strip() if i - 1 < len(linhas) else ""

            for i, linha in enumerate(linhas):
                if "HOMA APPLIANCES CO" in linha.upper():
                    return linhas[i - 2].strip() if i - 2 >= 0 else ""
    
        if provedor == "5JU":
            for i, linha in enumerate(linhas):
                if "REMARKS:" in linha.upper():
                    return linhas[i - 1].strip() if i > 0 else ""
            
            # Se não encontrar "REMARKS:", procurar "CREDIT NOTE"
            for i, linha in enumerate(linhas):
                if "CREDIT NOTE" in linha.upper():
                    return linhas[i - 4].strip() if i >= 4 else ""

        if provedor == "5DL":
            for i, linha in enumerate(linhas):
                if "COMMERCIAL INVOICE" in linha.upper():
                    return linhas[i - 2].strip() if i - 2 < len(linhas) else ""

        if provedor == "5DU":
            for i, linha in enumerate(linhas):
                if "TOTAL:" in linha.upper():
                    return linhas[i - 2].strip() if i - 2 < len(linhas) else ""

        if provedor == "5WY":
            for i, linha in enumerate(linhas):
                if "TOTAL" in linha.upper():
                    return linhas[i + 2].strip() if i + 2 < len(linhas) else ""

        if provedor == "7DQ":
            for i, linha in enumerate(linhas):
                if "PAYMENT CONDITIONS :" in linha.upper():
                    return linhas[i - 2].strip() if i - 2 < len(linhas) else ""

        if provedor == "7JR":
            for i, linha in enumerate(linhas):
                if "TOTAL AMOUNT:" in linha.upper():
                    return linhas[i + 2].strip() if i + 2 < len(linhas) else ""
            for i, linha in enumerate(linhas):
                if "REMARK:" in linha.upper() and i > 0:
                    return linhas[i - 1].strip()


        if provedor == "BRR":
            for i, linha in enumerate(linhas):
                if "COSTOS INTERNOS" in linha.upper():
                    valor = linhas[i + 4].strip() if i + 4 < len(linhas) else ""
                    if valor:
                        return valor
            for i, linha in enumerate(linhas):
                if "TOTAL IN FAVOUR" in linha.upper():
                    valor = linhas[i + 2].strip() if i + 2 < len(linhas) else ""
                    if valor:
                        return valor
            for i, linha in enumerate(linhas):
                if "TOTAL FOB" in linha.upper():
                    return linhas[i + 1].strip() if i + 1 < len(linhas) else ""

        if provedor == "CLH":
            for i, linha in enumerate(linhas):
                if "TOTAL FOB" in linha.upper():
                    return linhas[i + 1].strip() if i + 1 < len(linhas) else ""

        if provedor == "NINGBO HUA":
            for i, linha in enumerate(linhas):
                if "PAYMENT TERM" in linha.upper():
                    return linhas[i - 1].strip() if i > 0 else ""
            
            for i, linha in enumerate(linhas):
                if "AMOUNT IN WORDS" in linha.upper():
                    if i >= 2:
                        return linhas[i - 2].strip()
                    elif i >= 1:
                        return linhas[i - 1].strip()

        elif provedor == "US1239":
            for linha in linhas:
                linha_upper = linha.upper()
                if "GRAND TOTAL USA $" in linha_upper or "TOTALS" in linha_upper or "TOTAL VALUE" in linha_upper:
                    numeros = re.findall(r"\d[\d.,]*", linha)
                    if numeros:
                        return numeros[-1]  # Pega o último número da linha

                # Outras regras podem ser adicionadas aqui futuramente
        return ""
            
    df["Amount"] = df.apply(extrair_amount, axis=1)
    return df

def ajustar_amount(df):
    import re

    def limpar_amount(valor):
        if not isinstance(valor, str):
            valor = str(valor)
        # Remove letras, símbolo de dólar, hífens
        valor = re.sub(r'[A-Za-z$-]', '', valor).strip()

        # Regras de separadores
        if "," in valor and "." in valor:
            if valor.find(",") < valor.find("."):
                valor = valor.replace(",", "")  # Remove vírgula
            else:
                valor = valor.replace(".", "").replace(",", ".")  # Remove ponto, troca vírgula por ponto
        elif "," in valor:
            valor = valor.replace(",", ".")  # Troca vírgula por ponto

        # Conversão para float
        try:
            return float(valor)
        except ValueError:
            return None  # Ou np.nan, se preferir

    df["Amount"] = df["Amount"].apply(limpar_amount)
    return df

def adicionar_erro(df):
    def verificar_erro(Proveedor):
        return "Document can't be read" if not Proveedor.strip() else ""
    
    df["Error"] = df["Proveedor"].apply(verificar_erro)
    return df

def adicionar_colunas_fixas(df):

    df['Moneda'] = 'USD'
    df['Cod. Moneda'] = '01'
    df['Cuenta'] = '421202'
    return df

def adicionar_cod_autorizacion_ext(df):
    if 'Tipo Doc' in df.columns:
        df['Cód. de Autorización'] = df['Tipo Doc'].apply(
            lambda x: "91" if str(x).strip().upper() == "INVOICE"
            else "97" if str(x).strip().upper() == "CREDIT NOTE"
            else None
        )
    else:
        print("⚠️ Coluna 'Tipo Doc' não encontrada no DataFrame.")
    return df

def adicionar_tip_fac_ext(df):
    if 'Tipo Doc' in df.columns:
        df['Tipo de Factura'] = df['Tipo Doc'].apply(
            lambda x: "12" if str(x).strip().upper() == "INVOICE"
            else "12" if str(x).strip().upper() == "CREDIT NOTE"
            else None
        )
    else:
        print("⚠️ Coluna 'Tipo Doc' não encontrada no DataFrame.")
    return df

def remover_duplicatas_source_file(df):

    if 'source_file' in df.columns:
        return df.drop_duplicates(subset='source_file', keep='first')
    else:
        print("⚠️ Coluna 'source_file' não encontrada no DataFrame.")
        return df

def organizar_colunas_externos(df):

    colunas_desejadas = ['source_file','conteudo_pdf', 'Proveedor', 'Proveedor Iscala', 'Factura','Tipo Doc' ,'Cód. de Autorización','Tipo de Factura','Fecha de Emisión','Moneda', 
                         'Cod. Moneda', 'Amount', 'Tasa', 'Cuenta', 'Error']
    
    colunas_presentes = [col for col in colunas_desejadas if col in df.columns]

    df = df[colunas_presentes + [col for col in df.columns if col not in colunas_presentes]]
    
    return df

def op_gravada_negativo_CN_externos(df):
    if 'Tipo Doc' in df.columns and 'Amount' in df.columns:
        df['Amount'] = df.apply(
            lambda row: -abs(row['Amount']) if str(row['Tipo Doc']).strip().upper() == 'CREDIT NOTE' else row['Amount'],
            axis=1
        )
    return df


# ===========================================================
# MERGE PEC DO SHAREPOINT PELO NOME DO ARQUIVO
# ===========================================================

def merge_pec_fast(df_externos, df_sharepoint):

    df_ext = df_externos.copy()
    df_sp = df_sharepoint.copy()

    # Normaliza texto
    df_ext["key_ext"] = df_ext["source_file"].astype(str).str.lower()
    df_sp["key_sp"] = df_sp["name"].astype(str).str.lower()

    # Marca para combinar tudo
    df_ext["tmp"] = 1
    df_sp["tmp"] = 1

    # Faz todas as combinações possíveis (cartesiano)
    df_all = df_ext.merge(df_sp, on="tmp", suffixes=("_ext","_sp"))

    # Mantém somente onde há compatibilidade de nome
    df_all = df_all[
        df_all["key_ext"].str.contains(df_all["key_sp"], na=False)
        |
        df_all["key_sp"].str.contains(df_all["key_ext"], na=False)
    ]

    # Só precisamos de source_file + pec
    df_pec = df_all[["source_file", "pec"]].drop_duplicates()

    # Merge final inserindo a PEC no dataframe original
    df_final = df_externos.merge(df_pec, on="source_file", how="left")

    return df_final


def adicionar_pec_sharepoint(df_externos, df_sharepoint):
    """
    Função utilizada pelo service Externos para adicionar o número PEC.
    """
    if df_sharepoint is None or df_sharepoint.empty:
        return df_externos
    return merge_pec_fast(df_externos, df_sharepoint)
