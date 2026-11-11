"""Pacote de parsers de nomes de arquivo, arquivos DOCX e arquivos PDF."""

try:
    from parsers.filename_parser import (
        FilenameParseError,
        ParsedFilename,
        classificar_tipo_ato,
        determinar_resultado_sentenca,
        limpar_sufixos_conflito_onedrive,
        parse_filename,
    )
except ImportError:
    from dashboard_atos.parsers.filename_parser import (
        FilenameParseError,
        ParsedFilename,
        classificar_tipo_ato,
        determinar_resultado_sentenca,
        limpar_sufixos_conflito_onedrive,
        parse_filename,
    )

__all__ = [
    "FilenameParseError",
    "ParsedFilename",
    "classificar_tipo_ato",
    "determinar_resultado_sentenca",
    "limpar_sufixos_conflito_onedrive",
    "parse_filename",
]
