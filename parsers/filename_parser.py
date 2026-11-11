"""Parser do nome de arquivo .docx para atos judiciais.

Formatos suportados (podem coexistir):
  A) {NPU} - {Assessor} - {Palavras-chave} - {Descrição do ato}.docx
  B) {NPU} - {Assessor} - ED - {Palavras-chave} - {Descrição do ato}.docx
  C) {NPU} - ED - {Assessor} - {Palavras-chave} - {Descrição do ato}.docx
"""

from __future__ import annotations

from dataclasses import dataclass, field
import re
from typing import List, Optional
import unicodedata

REGEX_NPU = re.compile(r"(\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4})")


def normalizar_para_comparacao(texto: str) -> str:
    """Remove acentuação e converte para maiúsculas apenas para regras de comparação."""
    if not texto:
        return ""
    nfkd = unicodedata.normalize("NFKD", texto)
    sem_acento = "".join(c for c in nfkd if not unicodedata.combining(c))
    return sem_acento.strip().upper()


def limpar_sufixos_conflito_onedrive(texto: str) -> str:
    """Remove sufixos de cópia e conflito gerados pelo OneDrive ou Windows.

    Exemplos:
      - '... (1)', '... (2)'
      - '... - Copia', '... - Cópia', '... - Copy'
      - '...-DESKTOP-ABC1234', '...-LAPTOP-XYZ', '...-NOTEBOOK-ABC'
      - '... - Gabriel-PC'
    """
    if not texto:
        return ""
    t = texto.strip()
    # Remove sufixo numérico de cópia: " (1)", " (2)"
    t = re.sub(r"\s*\(\d+\)$", "", t)
    # Remove " - Copia", " - Cópia", " - Copy" (com ou sem numeração)
    t = re.sub(r"\s*-\s*C[oó]pia(\s*\(\d+\))?$", "", t, flags=re.IGNORECASE)
    t = re.sub(r"\s*-\s*Copy(\s*\(\d+\))?$", "", t, flags=re.IGNORECASE)
    # Remove sufixos clássicos de hostname Windows gerados em conflito de sincronização do OneDrive
    t = re.sub(
        r"(-DESKTOP-[A-Z0-9]+|-LAPTOP-[A-Z0-9]+|-NOTEBOOK-[A-Z0-9]+|\s*-\s*[A-Za-z0-9_\-]+-PC)$",
        "",
        t,
        flags=re.IGNORECASE,
    )
    return t.strip()


class FilenameParseError(Exception):
    """Exceção levantada quando o nome do arquivo foge dos padrões reconhecíveis."""
    pass


@dataclass
class ParsedFilename:
    npu: str
    assessor: str  # NÃO vai para a planilha (LGPD)
    tem_ed: bool
    palavras_chave_lista: List[str] = field(default_factory=list)
    palavras_chave_str: str = ""
    descricao_ato: str = ""
    tipo_ato: str = ""
    resultado_sentenca: str = ""


def classificar_tipo_ato(descricao: str, fallback_docx: Optional[str] = None) -> str:
    """Classifica o tipo de ato a partir da descrição ou cabeçalho do docx.

    Regras:
      - Começa com 'DESPACHO' -> Despacho
      - Começa com 'DECISÃO' / 'DECISAO' -> Decisão
      - Contém 'SENTENÇA', 'PARCIAL PROCEDENTE', 'PROCEDENTE', 'IMPROCEDENTE',
        'EXTINTO', 'EXTINÇÃO', 'HOMOLOGADO' -> Sentença
      - Ambíguo: cabeçalhos no .docx (primeiras 50 linhas)
      - Dúvida persistente: Indefinido
    """
    desc_norm = normalizar_para_comparacao(descricao)

    if desc_norm.startswith("DESPACHO"):
        return "Despacho"

    if desc_norm.startswith("DECISAO") or desc_norm.startswith("DECISAO INTERLOCUTORIA"):
        return "Decisão"

    palavras_chave_sentenca = [
        "SENTENCA",
        "PARCIAL PROCEDENTE",
        "PARCIALMENTE PROCEDENTE",
        "PROCEDENTE",
        "IMPROCEDENTE",
        "EXTINTO",
        "EXTINCAO",
        "HOMOLOGADO",
        "HOMOLOGACAO",
        "ACORDO",
        "PROVIMENTO",
        "ACOLHENDO",
        "ACOLHO",
        "NAO ACOLHENDO",
        "NAO ACOLHO",
        "NEGANDO PROVIMENTO",
        "NEGO PROVIMENTO",
        "DOU PROVIMENTO",
        "EMBARGOS",
    ]

    for termo in palavras_chave_sentenca:
        if termo in desc_norm:
            return "Sentença"

    if fallback_docx:
        fallback_norm = normalizar_para_comparacao(fallback_docx)
        for cabecalho in ["SENTENCA", "DECISAO", "DESPACHO"]:
            pos = fallback_norm.find(cabecalho)
            if pos != -1:
                if cabecalho == "SENTENCA":
                    return "Sentença"
                elif cabecalho == "DECISAO":
                    return "Decisão"
                elif cabecalho == "DESPACHO":
                    return "Despacho"

    return "Indefinido"


def determinar_resultado_sentenca(descricao: str, tipo_ato: str) -> str:
    """Determina o resultado da sentença (apenas quando tipo_ato == 'Sentença').

    Mapeamento case-insensitive e sem acentos:
      - PARCIAL PROCEDENTE / PARCIALMENTE PROCEDENTE -> Parcialmente procedente
      - PROCEDENTE (sem 'parcial'/'improcedente') -> Procedente
      - IMPROCEDENTE -> Improcedente
      - EXTINTO SEM RESOLUÇÃO / EXTINÇÃO SEM MÉRITO -> Extinto sem mérito
      - HOMOLOGADO / HOMOLOGO / ACORDO -> Homologação de acordo
      - NEGANDO PROVIMENTO / NEGO PROVIMENTO -> Negado provimento (ED)
      - ACOLHENDO / ACOLHO / DOU PROVIMENTO -> Provimento (ED)
      - NÃO ACOLHENDO / NÃO ACOLHO -> Não acolhimento (ED)
      - Outros -> Outro
    """
    if tipo_ato != "Sentença":
        return ""

    desc_norm = normalizar_para_comparacao(descricao)

    # 1. Parcialmente procedente
    if "PARCIAL PROCEDENTE" in desc_norm or "PARCIALMENTE PROCEDENTE" in desc_norm:
        return "Parcialmente procedente"

    # 2. Improcedente
    if "IMPROCEDENTE" in desc_norm:
        return "Improcedente"

    # 3. Procedente puro
    if "PROCEDENTE" in desc_norm and "PARCIAL" not in desc_norm and "IMPROCEDENTE" not in desc_norm:
        return "Procedente"

    # 4. Extinção sem mérito / sem resolução
    if (
        "EXTINTO SEM RESOLUCAO" in desc_norm
        or "EXTINCAO SEM MERITO" in desc_norm
        or "EXTINTO SEM MERITO" in desc_norm
        or "EXTINCAO SEM RESOLUCAO" in desc_norm
        or "EXTINTO" in desc_norm
        or "EXTINCAO" in desc_norm
    ):
        return "Extinto sem mérito"

    # 5. Homologação de acordo
    if "HOMOLOGADO" in desc_norm or "HOMOLOGO" in desc_norm or "HOMOLOGACAO" in desc_norm or "ACORDO" in desc_norm:
        return "Homologação de acordo"

    # 6. Embargos de Declaração (ED)
    if "NEGANDO PROVIMENTO" in desc_norm or "NEGO PROVIMENTO" in desc_norm:
        return "Negado provimento (ED)"

    if "NAO ACOLHENDO" in desc_norm or "NAO ACOLHO" in desc_norm:
        return "Não acolhimento (ED)"

    if "ACOLHENDO" in desc_norm or "ACOLHO" in desc_norm or "DOU PROVIMENTO" in desc_norm or "PROVIMENTO" in desc_norm:
        return "Provimento (ED)"

    return "Outro"


def parse_filename(nome_arquivo: str) -> ParsedFilename:
    """Extrai campos do nome do arquivo .docx conforme as regras A, B e C."""
    # Remove extensão
    base = nome_arquivo
    if base.lower().endswith(".docx"):
        base = base[:-5]
    elif base.lower().endswith(".doc"):
        base = base[:-4]

    # Limpa sufixos de cópia e conflito gerados pelo OneDrive ou Windows
    base = limpar_sufixos_conflito_onedrive(base)

    # 1. Extração do NPU
    match_npu = REGEX_NPU.search(base)
    if not match_npu:
        raise FilenameParseError(f"Nenhum NPU válido no formato CNJ encontrado em '{nome_arquivo}'.")

    npu = match_npu.group(1)

    # 2. Remover NPU e dividir por " - "
    resto = base.replace(npu, "", 1).strip()
    # Limpa hífens ou espaços iniciais
    resto = re.sub(r"^[\s\-]+", "", resto).strip()

    tokens = [t.strip() for t in re.split(r"\s+-\s+", resto) if t.strip()]
    if not tokens:
        raise FilenameParseError(f"Nome do arquivo '{nome_arquivo}' não possui dados após o NPU.")

    # 3. Detectar flag ED
    tem_ed = False
    tokens_filtrados = []
    for t in tokens:
        if normalizar_para_comparacao(t) == "ED":
            tem_ed = True
        else:
            tokens_filtrados.append(t)

    if not tokens_filtrados:
        raise FilenameParseError(f"Nome do arquivo '{nome_arquivo}' contém apenas marcador ED.")

    # 4. Classificar Assessor, Palavras-chave e Descrição do ato
    assessor = ""
    descricao_ato = ""
    palavras_chave_tokens: List[str] = []

    # A Descrição do ato é sempre o último token (limpo de sufixos de conflito)
    descricao_ato = limpar_sufixos_conflito_onedrive(tokens_filtrados[-1])
    intermediarios = tokens_filtrados[:-1]

    if intermediarios:
        primeiro = intermediarios[0]
        # Assessor: primeiro token restante que não contenha vírgulas e tenha <= 40 chars
        if "," not in primeiro and len(primeiro) <= 40:
            assessor = primeiro
            palavras_intermediarias = intermediarios[1:]
        else:
            # Não é assessor (pode ser palavras-chave diretamente)
            palavras_intermediarias = intermediarios

        for token in palavras_intermediarias:
            # Divide palavras-chave por vírgula
            partes = [p.strip() for p in token.split(",") if p.strip()]
            palavras_chave_tokens.extend(partes)
    else:
        # Se só temos 1 token após remover NPU e ED, esse token é a descrição do ato
        descricao_ato = tokens_filtrados[0]

    # Palavras-chave formatadas com "; " entre termos, preservando acentuação original
    palavras_chave_str = "; ".join(palavras_chave_tokens) if palavras_chave_tokens else ""

    tipo_ato = classificar_tipo_ato(descricao_ato)
    resultado_sentenca = determinar_resultado_sentenca(descricao_ato, tipo_ato)

    return ParsedFilename(
        npu=npu,
        assessor=assessor,
        tem_ed=tem_ed,
        palavras_chave_lista=palavras_chave_tokens,
        palavras_chave_str=palavras_chave_str,
        descricao_ato=descricao_ato,
        tipo_ato=tipo_ato,
        resultado_sentenca=resultado_sentenca,
    )
