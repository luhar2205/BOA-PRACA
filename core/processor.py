import pandas as pd
from pathlib import Path
from thefuzz import fuzz
import re
from core.xml_reader import extrair_dados_xml, calcular_ua_final_blindada

def _expand_abbreviations(text):
    """🧠 Dicionário de Tradução: Transforma jargão de nota fiscal em português claro"""
    replaces = {
        r'\bF\.\b': 'FILE',
        r'\bF\b': 'FILE',
        r'\bSALMON\b': 'SALMAO',
        r'\bDEF\b': 'DEFUMADO',
        r'\bS/E\b': 'SEM ESPINHA',
        r'\bS/O\b': 'SEM OSSO',
        r'\bC/O\b': 'COM OSSO',
        r'\bCXA\b': 'CAIXA',
        r'\bCX\b': 'CAIXA',
        r'\bPCT\b': 'PACOTE',
        r'\bPT\b': 'PACOTE',
        r'\bRESF\b': 'RESFRIADO',
        r'\BCONG\b': 'CONGELADO',
        r'\bLING\b': 'LINGUICA',
        r'\bCALAB\b': 'CALABRESA',
        r'\bPREM\b': 'PREMIUM',
        r'\bESP\b': 'ESPECIAL'
    }
    for pattern, replacement in replaces.items():
        text = re.sub(pattern, replacement, text)
    return text

def _norm(v): 
    if pd.isna(v): return ""
    text = str(v).strip().upper()
    return _expand_abbreviations(text)

def _clean_ean(ean):
    """⚔️ Filtro de Limpeza: Arranca zeros e converte DUN-14 para EAN-13"""
    e = str(ean).strip()
    if e in ["", "SEM GTIN", "N/A", "NONE"]: return ""
    e = e.lstrip('0') 
    if len(e) == 14 and e.startswith('1'):
        e = e[1:] 
    return e

def rodar_automacao_total(pasta_xml, arquivo_base):
    df_raw = pd.read_excel(arquivo_base, header=2).fillna("")
    df_raw.columns = [c.strip() for c in df_raw.columns]
    
    ean_map = {}
    for i, row in df_raw.iterrows():
        ref = _clean_ean(row.get('Referência', ''))
        if ref: ean_map[ref] = i
    
    xmls = list(Path(pasta_xml).glob("*.xml"))
    todos_itens = []
    for p in xmls: todos_itens.extend(extrair_dados_xml(p))

    final = []
    for it in todos_itens:
        match_idx = None
        tipo_match = "❌ SEM MATCH"
        
        # Guarda o original para o Excel, mas usa o normalizado/traduzido para pensar
        desc_xml_raw = str(it["Produto_XML"]).strip().upper()
        desc_xml_norm = _norm(it["Produto_XML"]) 
        ean_xml = _clean_ean(it["EAN_XML"]) 

        if ean_xml and ean_xml in ean_map:
            match_idx = ean_map[ean_xml]
            tipo_match = "✅ EAN"
        
        if match_idx is None:
            scores = []
            for i, row in df_raw.iterrows():
                desc_s_norm = _norm(row.get('Descrição', ''))
                marca_s = _norm(row.get('Marca', ''))
                
                # A inteligência agora compara a versão traduzida das duas descrições!
                score_base = fuzz.token_set_ratio(desc_xml_norm, desc_s_norm)
                
                if marca_s and marca_s in desc_xml_norm:
                    score_base += 20
                
                if any(x in desc_xml_norm for x in ["500G", "1KG", "0.5KG", "360G", "180G"]) and any(x in desc_s_norm for x in ["500G", "1KG", "0.5KG", "360G", "180G"]):
                    score_base += 15

                if score_base > 75: 
                    scores.append((i, score_base))
            
            if scores:
                scores.sort(key=lambda x: x[1], reverse=True)
                match_idx = scores[0][0]
                tipo_match = f"🔶 INTELLIGENT ({scores[0][1]}%)"

        row_s = df_raw.iloc[match_idx] if match_idx is not None else {}
        desc_s_final = str(row_s.get('Descrição', "")).strip().upper()
        
        ua = calcular_ua_final_blindada(desc_xml_raw, desc_s_final, it["QTD_NOTA"], it["qTrib"])

        final.append({
            "NF": it["NF"],
            "FORNECEDOR": it["FORNECEDOR"],
            "EAN XML": it["EAN_XML"], 
            "PRODUTO XML": desc_xml_raw,
            "UN NOTA": it["UN_NOTA"],
            "QTD NOTA": it["QTD_NOTA"],
            "VLR UNIT NOTA": it["VLR_UNIT_NOTA"],
            "CÓD SANKHYA": row_s.get('Código', ""),
            "DESC SANKHYA": desc_s_final,
            "UA / CONVERSÃO": ua,
            "MATCH": tipo_match
        })

    return pd.DataFrame(final)