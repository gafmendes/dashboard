"""Parser do conteúdo de arquivos .docx com estrita observância à LGPD.

Extrai apenas:
  - NPU validado a partir das primeiras ~50 linhas do documento
  - Rótulos e validação de partes para verificação cruzada (NUNCA gravadas em planilha)
  - Cabeçalhos de atos (SENTENÇA, DECISÃO, DESPACHO) para desambiguação
"""

from __future__ import annotations

from dataclasses import dataclass, field
import logging
from pathlib import Path
import re
from typing import List, Optional

logger = logging.getLogger("dashboard_atos.docx_parser")

REGEX_PROCESSO_CORPO = re.compile(
    r"Processo\s*n[º°o]?[.:]?\s*(\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4})",
    re.IGNORECASE,
)

ROTULOS_PARTES = [
    "AUTOR",
    "AUTORA",
    "AUTORES",
    "REU",
    "RÉU",
    "RE",
    "RÉ",
    "REUS",
    "RÉUS",
    "EXEQUENTE",
    "EXECUTADO",
    "EXECUTADA",
    "REQUERENTE",
    "REQUERIDO",
    "REQUERIDA",
    "HERDEIRO",
    "HERDEIRA",
    "INVENTARIANTE",
    "INVENTARIADO",
    "IMPETRANTE",
    "IMPETRADO",
    "EMBARGANTE",
    "EMBARGADO",
    "RECORRENTE",
    "RECORRIDO",
    "AGRAVANTE",
    "AGRAVADO",
    "APELANTE",
    "APELADO",
]

REGEX_ROTULO_PARTES = re.compile(
    r"^\s*(" + "|".join(ROTULOS_PARTES) + r")\s*[:\-]\s*(.*)$",
    re.IGNORECASE,
)


@dataclass
class DocxExtracao:
    npu_corpo: Optional[str] = None
    cabecalho_ato: Optional[str] = None
    texto_primeiras_linhas: List[str] = field(default_factory=list)
    # Partes são mantidas APENAS em memória durante a validação (LGPD: jamais persistir)
    tem_partes: bool = False
    rotulos_detectados: List[str] = field(default_factory=list)


def extrair_linhas_docx(caminho_docx: Path, max_linhas: int = 50) -> List[str]:
    """Lê até `max_linhas` parágrafos/linhas de um arquivo .docx."""
    try:
        import docx  # type: ignore
    except ImportError:
        logger.debug("python-docx não instalado. Modo de leitura simplificada/mock ativo.")
        return []

    try:
        doc = docx.Document(str(caminho_docx))
        linhas = []
        for p in doc.paragraphs:
            texto = p.text.strip()
            if texto:
                linhas.append(texto)
            if len(linhas) >= max_linhas:
                break
        return linhas
    except Exception as exc:
        logger.warning("Falha ao abrir .docx '%s': %s", caminho_docx.name, exc)
        return []


def parse_conteudo_docx(linhas: List[str]) -> DocxExtracao:
    """Extrai NPU, cabeçalhos de ato e valida presença de partes a partir das primeiras linhas."""
    resultado = DocxExtracao(texto_primeiras_linhas=linhas[:50])

    for linha in linhas[:50]:
        # Busca NPU no corpo
        if not resultado.npu_corpo:
            m = REGEX_PROCESSO_CORPO.search(linha)
            if m:
                resultado.npu_corpo = m.group(1)

        # Busca cabeçalhos para desambiguação
        linha_upper = linha.strip().upper()
        if not resultado.cabecalho_ato:
            for cab in ["SENTENÇA", "SENTENCA", "DECISÃO", "DECISAO", "DESPACHO"]:
                if cab in linha_upper:
                    resultado.cabecalho_ato = cab
                    break

        # Detecção de partes para validação (sem armazenar os nomes de pessoas físicas/jurídicas)
        match_parte = REGEX_ROTULO_PARTES.match(linha)
        if match_parte:
            resultado.tem_partes = True
            rotulo = match_parte.group(1).upper()
            if rotulo not in resultado.rotulos_detectados:
                resultado.rotulos_detectados.append(rotulo)

    return resultado
