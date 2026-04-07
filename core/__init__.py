"""
Módulo de leitura e extração de dados das NF-e (XML).
Estrutura padrão: namespace urn:nfe NF-e 4.0
"""
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional
import logging

logger = logging.getLogger(__name__)

NS = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

def _text(element, path: str) -> Optional[str]:
    """Retorna o texto de um subelemento ou None se não existir."""
    node = element.find(path, NS)
    return node.text.strip() if node is not None and node.text else None

def extrair_itens_xml(caminho_xml: Path) -> list[dict]:
    """
    Lê um arquivo XML de NF-e e retorna lista de dicts com os campos relevantes.
    """
    try:
        tree = ET.parse(caminho_xml)
        root = tree.getroot()

        for elem in root.iter():
            if "}" in elem.tag:
                elem.tag = elem.tag.split("}", 1)[1]

        ide  = root.find(".//ide")
        emit = root.find(".//emit")
        num_nf    = _sem_ns(ide, "nNF")   if ide  else None
        fornecedor = _sem_ns(emit, "xNome") if emit else None

        itens = []
        for det in root.findall(".//det"):
            prod = det.find("prod")
            if prod is None:
                continue

            ean  = _sem_ns(prod, "cEAN")
            desc = _sem_ns(prod, "xProd")
            ua   = _sem_ns(prod, "uCom")
            n_item = det.attrib.get("nItem", "")

            if ean in ("SEM GTIN", ""):
                ean = None

            itens.append({
                "NF":            num_nf,
                "Fornecedor":    fornecedor,
                "Item":          n_item,
                "Produto_XML":   desc,
                "EAN_XML":       ean,
                "UA_XML":        ua,
                "_arquivo":      caminho_xml.name,
            })

        return itens

    except ET.ParseError as e:
        logger.warning(f"[XML CORROMPIDO] {caminho_xml.name}: {e}")
        return []
    except Exception as e:
        logger.warning(f"[ERRO] {caminho_xml.name}: {e}")
        return []

def _sem_ns(element, tag: str) -> Optional[str]:
    if element is None: return None
    node = element.find(tag)
    return node.text.strip() if node is not None and node.text else None