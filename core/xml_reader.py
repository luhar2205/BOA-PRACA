import xml.etree.ElementTree as ET
import re

def calcular_peso_pacote_unitario(desc):
    """Identifica multiplicadores (2X180GR) e gramaturas simples (500G)"""
    if not desc: return None
    desc = desc.upper().replace(',', '.')
    
    match_mult = re.search(r'(\d+)\s*X\s*(\d+[.]?\d*)\s*(KG|G|GR)', desc)
    if match_mult:
        qtd_interna = float(match_mult.group(1))
        peso_cada = float(match_mult.group(2))
        unid = match_mult.group(3)
        total = (qtd_interna * peso_cada)
        return total if unid == 'KG' else total / 1000

    match_simples = re.search(r'(\d+[.]?\d*)\s*(KG|G|GR)', desc)
    if match_simples:
        valor = float(match_simples.group(1))
        unid = match_simples.group(2)
        return valor if unid == 'KG' else valor / 1000
    return None

def calcular_ua_final_blindada(desc_xml, desc_sankhya, q_com, q_trib):
    if q_com <= 0: return "0.000000"
    
    # 🧠 O PULO DO GATO: Prioriza extrair o peso da descrição MATRIZ (Sankhya)
    peso_pacote = calcular_peso_pacote_unitario(desc_sankhya)
    
    # Se não achou na matriz, tenta extrair da descrição do XML
    if not peso_pacote:
        peso_pacote = calcular_peso_pacote_unitario(desc_xml)
    
    ua_bruta = q_trib / q_com
    
    if peso_pacote and peso_pacote > 0:
        # A conversão real é o peso tributável / peso do pacote real
        ua_real = ua_bruta / peso_pacote
        peso_str = f"{peso_pacote*1000:.0f}g" if peso_pacote < 1 else f"{peso_pacote:.1f}kg"
        return f"{ua_real:.6f} (Ref: {round(ua_real)} un de {peso_str}/cx)"
    
    # Se não tem peso em lugar nenhum (Ex: Salame Mini 6KG onde 6KG é enganoso)
    # Confia na divisão bruta do XML (qTrib / qCom)
    return f"{ua_bruta:.6f}"

def extrair_dados_xml(caminho):
    try:
        tree = ET.parse(caminho)
        root = tree.getroot()
        for elem in root.iter():
            if "}" in elem.tag: elem.tag = elem.tag.split("}", 1)[1]
        
        num_nf = root.find(".//ide/nNF").text
        emitente = root.find(".//emit/xNome").text
        
        itens = []
        for det in root.findall(".//det"):
            prod = det.find("prod")
            ean_trib = prod.find("cEANTrib").text if prod.find("cEANTrib") is not None else ""
            ean_com = prod.find("cEAN").text if prod.find("cEAN") is not None else ""
            ean_final = ean_trib if ean_trib not in ["SEM GTIN", ""] else ean_com

            itens.append({
                "NF": num_nf, "FORNECEDOR": emitente,
                "Produto_XML": prod.find("xProd").text,
                "UN_NOTA": prod.find("uCom").text if prod.find("uCom") is not None else "",
                "EAN_XML": ean_final,
                "QTD_NOTA": float(prod.find("qCom").text or 0),
                "VLR_UNIT_NOTA": float(prod.find("vUnCom").text or 0),
                "qTrib": float(prod.find("qTrib").text or 0)
            })
        return itens
    except: return []