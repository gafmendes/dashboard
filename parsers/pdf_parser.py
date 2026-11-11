"""Parser do PDF pareado para extração dos metadados processuais.

Campos extraídos:
  - Classe judicial
  - Assunto principal (código / descrição CNJ)
  - Data de autuação (convertida para DD/MM/AAAA)
  - Valor da causa (formatado em moeda BRL: R$ X.XXX,XX)
  - Segredo de justiça (Sim / Não)
  - Data de conclusão (data mais recente de conclusão/conclusos anterior ao ato)
  - Remessa (normalizada, ex.: '32ª Remessa' ou 'PENDENTE')
"""

from __future__ import annotations

from dataclasses import dataclass
import datetime
import logging
from pathlib import Path
import re
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger("dashboard_atos.pdf_parser")

MESES_BR = {
    "janeiro": "01",
    "fevereiro": "02",
    "marco": "03",
    "março": "03",
    "abril": "04",
    "maio": "05",
    "junho": "06",
    "julho": "07",
    "agosto": "08",
    "setembro": "09",
    "outubro": "10",
    "novembro": "11",
    "dezembro": "12",
}

REGEX_DATA_BR = re.compile(r"\b(\d{2})/(\d{2})/(\d{4})\b")
REGEX_DATA_EXTENSO = re.compile(
    r"\b(\d{1,2})\s+de\s+([a-zA-ZçÇãÃõÕáÁéÉíÍóÓúÚ]+)\s+de\s+(\d{4})\b",
    re.IGNORECASE,
)
REGEX_VALOR_CAUSA = re.compile(r"R\$\s*([\d\.]+(?:,\d{2})?)")
REGEX_REMESSA = re.compile(
    r"\b(\d+)[ \t]*(?:ª|º|\.ª|\.º)?[ \t]*Remessa|Remessa[ \t]*(?:[:\-]|n[º°o]?[.:]?)*[ \t]*(\d+)",
    re.IGNORECASE,
)


@dataclass
class PdfMetadados:
    encontrado: bool = False
    caminho_pdf: Optional[Path] = None
    classe_judicial: str = "PENDENTE"
    assunto_principal: str = "PENDENTE"
    data_autuacao: str = "PENDENTE"
    valor_causa: str = "PENDENTE"
    segredo_justica: str = "Não"
    data_conclusao: str = ""
    remessa: str = "PENDENTE"
    data_assinatura: Optional[str] = None


def normalizar_data(texto_data: str) -> str:
    """Converte datas em formatos DD/MM/AAAA ou 'DD de mês de AAAA' para DD/MM/AAAA."""
    if not texto_data:
        return ""

    match_padrao = REGEX_DATA_BR.search(texto_data)
    if match_padrao:
        d, m, y = match_padrao.groups()
        return f"{int(d):02d}/{int(m):02d}/{y}"

    match_extenso = REGEX_DATA_EXTENSO.search(texto_data)
    if match_extenso:
        dia, mes_nome, ano = match_extenso.groups()
        mes_key = mes_nome.lower()
        mes_num = MESES_BR.get(mes_key)
        if mes_num:
            return f"{int(dia):02d}/{mes_num}/{ano}"

    return texto_data.strip()


def encontrar_pdf_pareado(caminho_docx: Path, todos_pdfs: List[Path], npu: str) -> Optional[Path]:
    """Pareia o arquivo .docx com seu respectivo .pdf.

    Critérios (nesta ordem):
      1. NPU idêntico; se houver mais de um, desempate pelo prefixo do nome até o 2º ' - '.
      2. Nome-base idêntico (sem extensão).
      3. Nenhum encontrado -> None
    """
    # Critério 2 direto: nome base idêntico
    nome_base_docx = caminho_docx.stem.lower()
    for p in todos_pdfs:
        if p.stem.lower() == nome_base_docx:
            return p

    # Critério 1: NPU idêntico
    pdfs_mesmo_npu = [p for p in todos_pdfs if npu in p.name]
    if len(pdfs_mesmo_npu) == 1:
        return pdfs_mesmo_npu[0]
    elif len(pdfs_mesmo_npu) > 1:
        # Desempate pelo prefixo até o 2º ' - '
        partes_docx = caminho_docx.name.split(" - ")
        prefixo_docx = " - ".join(partes_docx[:2]).lower() if len(partes_docx) >= 2 else partes_docx[0].lower()
        for p in pdfs_mesmo_npu:
            partes_pdf = p.name.split(" - ")
            prefixo_pdf = " - ".join(partes_pdf[:2]).lower() if len(partes_pdf) >= 2 else partes_pdf[0].lower()
            if prefixo_docx == prefixo_pdf:
                return p
        return pdfs_mesmo_npu[0]

    return None


def extrair_texto_pdf(caminho_pdf: Path, max_paginas: int = 3) -> str:
    """Extrai texto das primeiras páginas do PDF usando pdfplumber."""
    try:
        import pdfplumber  # type: ignore
    except ImportError:
        logger.debug("pdfplumber não instalado. Extração de PDF em modo fallback.")
        return ""

    try:
        texto_acumulado = []
        with pdfplumber.open(str(caminho_pdf)) as pdf:
            for pagina in pdf.pages[:max_paginas]:
                t = pagina.extract_text()
                if t:
                    texto_acumulado.append(t)
        return "\n".join(texto_acumulado)
    except Exception as exc:
        logger.warning("Falha ao ler PDF '%s': %s", caminho_pdf.name, exc)
        return ""


def parse_metadados_processuais(texto_pdf: str) -> PdfMetadados:
    """Analisa o texto do PDF com expressões regulares tolerantes a acentuação e caixa."""
    meta = PdfMetadados(encontrado=True)

    if not texto_pdf or not texto_pdf.strip():
        return meta

    # 1. Classe judicial
    m_classe = re.search(r"Classe(?:\s+judicial)?\s*[:\-]?\s*([^\n\r]+)", texto_pdf, re.IGNORECASE)
    if m_classe:
        meta.classe_judicial = m_classe.group(1).strip()

    # 2. Assunto principal (código / descrição CNJ)
    m_assunto = re.search(r"Assunto(?:\s+principal)?\s*[:\-]?\s*([^\n\r]+)", texto_pdf, re.IGNORECASE)
    if m_assunto:
        meta.assunto_principal = m_assunto.group(1).strip()

    # 3. Data de autuação / distribuição
    m_autuacao = re.search(
        r"(?:Autua[çc][ãa]o|Distribui[çc][ãa]o|Distribu[ií]do em)\s*[:\-]?\s*([^\n\r,]+)",
        texto_pdf,
        re.IGNORECASE,
    )
    if m_autuacao:
        meta.data_autuacao = normalizar_data(m_autuacao.group(1))

    # 4. Valor da causa
    m_valor = REGEX_VALOR_CAUSA.search(texto_pdf)
    if m_valor:
        meta.valor_causa = f"R$ {m_valor.group(1).strip()}"

    # 5. Segredo de justiça
    # Procura "segredo de justiça", "sigilo", "SJ" isolado
    if re.search(r"\b(segredo\s+de\s+justi[çc]a|sigilo(?:\s+absoluto)?|\bSJ\b)\b", texto_pdf, re.IGNORECASE):
        # Verifica se não é "Sem segredo de justiça" ou "Não"
        if not re.search(r"\b(?:sem|n[ãa]o)\s+segredo\s+de\s+justi[çc]a\b", texto_pdf, re.IGNORECASE):
            meta.segredo_justica = "Sim"

    # 6. Remessa
    m_remessa = REGEX_REMESSA.search(texto_pdf)
    if m_remessa:
        num = m_remessa.group(1) or m_remessa.group(2)
        meta.remessa = f"{int(num)}ª Remessa"

    # 7. Data de conclusão (mais recente anterior ao ato)
    # Busca trechos como "Conclusos para despacho/decisão/sentença em DD/MM/AAAA" ou "Conclusão: DD/MM/AAAA"
    matches_conclusao = re.findall(
        r"(?:conclus[ãa]o|conclusos?|vista)[^\n\r\d]*(\d{2}/\d{2}/\d{4})",
        texto_pdf,
        re.IGNORECASE,
    )
    if matches_conclusao:
        meta.data_conclusao = matches_conclusao[-1]

    # 8. Data de assinatura / publicação se expressa
    m_assinatura = re.search(
        r"(?:assinado\s+(?:eletronicamente|digitalmente)\s+em|data\s+de\s+juntada\s*[:\-]?)\s*(\d{2}/\d{2}/\d{4})",
        texto_pdf,
        re.IGNORECASE,
    )
    if m_assinatura:
        meta.data_assinatura = m_assinatura.group(1)

    return meta
