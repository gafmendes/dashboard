"""Ponto de entrada principal do consolidador de atos judiciais TJPE (8.17).

Suporta execução interativa, via agendador ou duplo clique (.bat), com flags:
  --dry-run: simulação sem alteração de arquivos
  --force-full: reprocessamento integral de todos os arquivos .docx
  --pasta: caminho alternativo para a pasta de arquivos
"""

from __future__ import annotations

import argparse
import datetime
import logging
from logging.handlers import RotatingFileHandler
import os
from pathlib import Path
import sys
import time
from typing import List

# Permite execução direta tanto como script (python main.py) quanto como pacote (python -m dashboard_atos.main)
_BASE_DIR = Path(__file__).resolve().parent
if str(_BASE_DIR) not in sys.path:
    sys.path.insert(0, str(_BASE_DIR))
if str(_BASE_DIR.parent) not in sys.path:
    sys.path.append(str(_BASE_DIR.parent))

try:
    import config
    from estado import EstadoRemoto, LockError
    from excel_writer import (
        ExcelDashboardWriter,
        RegistroAto,
        RegistroErro,
    )
    from parsers.docx_parser import (
        extrair_linhas_docx,
        parse_conteudo_docx,
    )
    from parsers.filename_parser import (
        FilenameParseError,
        classificar_tipo_ato,
        determinar_resultado_sentenca,
        limpar_sufixos_conflito_onedrive,
        parse_filename,
    )
    from parsers.pdf_parser import (
        encontrar_pdf_pareado,
        extrair_texto_pdf,
        parse_metadados_processuais,
    )
    from html_writer import gerar_dashboard_html
except ImportError:
    from dashboard_atos import config
    from dashboard_atos.estado import EstadoRemoto, LockError
    from dashboard_atos.excel_writer import (
        ExcelDashboardWriter,
        RegistroAto,
        RegistroErro,
    )
    from dashboard_atos.parsers.docx_parser import (
        extrair_linhas_docx,
        parse_conteudo_docx,
    )
    from dashboard_atos.parsers.filename_parser import (
        FilenameParseError,
        classificar_tipo_ato,
        determinar_resultado_sentenca,
        limpar_sufixos_conflito_onedrive,
        parse_filename,
    )
    from dashboard_atos.parsers.pdf_parser import (
        encontrar_pdf_pareado,
        extrair_texto_pdf,
        parse_metadados_processuais,
    )
    from dashboard_atos.html_writer import gerar_dashboard_html


def configurar_logger() -> logging.Logger:
    """Configura logger com saída no console e arquivo rotativo dashboard_atos.log."""
    logger = logging.getLogger("dashboard_atos")
    logger.setLevel(logging.INFO)

    # Evita duplicação de handlers se chamado múltiplas vezes
    if logger.handlers:
        return logger

    formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)s] [%(name)s]: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    try:
        config.ARQUIVO_LOG.parent.mkdir(parents=True, exist_ok=True)
        file_handler = RotatingFileHandler(
            str(config.ARQUIVO_LOG),
            maxBytes=config.LOG_MAX_BYTES,
            backupCount=config.LOG_BACKUP_COUNT,
            encoding="utf-8",
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except Exception as exc:
        logger.warning("Não foi possível criar log rotativo em disco: %s", exc)

    return logger


def formatar_data_mtime(mtime_timestamp: float) -> str:
    """Converte timestamp float para string DD/MM/AAAA."""
    dt = datetime.datetime.fromtimestamp(mtime_timestamp)
    return dt.strftime("%d/%m/%Y")


def executar(
    pasta_downloads: Path,
    dry_run: bool = False,
    force_full: bool = False,
) -> int:
    """Executa o ciclo completo de consolidação com filtro de 3 camadas."""
    logger = configurar_logger()
    inicio_iso = datetime.datetime.now().isoformat()
    logger.info("=" * 70)
    logger.info("INICIANDO CONSOLIDAÇÃO DE ATOS JUDICIAIS - TJPE (8.17)")
    logger.info("Perfil ativo: %s (%s)", config.PERFIL_SELECIONADO, config._cfg.get("descricao", ""))
    logger.info("Pasta monitorada (.docx/.pdf): %s", pasta_downloads)
    logger.info("Pasta de sincronização OneDrive: %s", config.PASTA_ONEDRIVE_DASHBOARD)
    logger.info("Modo Dry-Run: %s | Force-Full: %s", dry_run, force_full)
    logger.info("=" * 70)

    if not pasta_downloads.exists():
        logger.error("A pasta especificada não existe: %s", pasta_downloads)
        return 1

    # 1. Inicializa gerenciador de estado remoto
    estado = EstadoRemoto(
        remoto_sql=config.REMOTO_SQL,
        backup_dir=config.BACKUP_DIR,
        lock_ttl_minutos=config.LOCK_TTL_MINUTOS,
        retry_replace=config.RETRY_REPLACE,
        retry_sleep_seg=config.RETRY_SLEEP_SEG,
    )

    try:
        if not dry_run:
            estado.adquirir_lock()
        estado.carregar()
    except LockError as le:
        logger.error("Abortando execução devido ao lock: %s", le)
        return 2
    except Exception as exc:
        logger.critical("Falha ao inicializar estado remoto: %s", exc)
        return 3

    try:
        # 2. Localiza todos os arquivos .docx e .pdf na pasta
        # Agrupa arquivos .docx pelo nome canônico (removendo sufixos de cópia '(1)', '- Copia' do OneDrive)
        docx_por_canonica: Dict[str, List[Path]] = {}
        todos_pdf: List[Path] = []

        for p in pasta_downloads.iterdir():
            if p.is_file():
                # Ignora arquivos temporários do Word (~$*)
                if p.name.startswith("~$"):
                    continue
                ext = p.suffix.lower()
                if ext == ".docx":
                    stem_limpo = limpar_sufixos_conflito_onedrive(p.stem)
                    nome_canonico = stem_limpo.lower() + ext
                    docx_por_canonica.setdefault(nome_canonico, []).append(p)
                elif ext == ".pdf":
                    todos_pdf.append(p)

        todos_docx: List[Path] = []
        duplicatas_onedrive_disco = 0
        for nome_can, lista_arqs in docx_por_canonica.items():
            if len(lista_arqs) == 1:
                todos_docx.append(lista_arqs[0])
            else:
                # Múltiplas cópias do mesmo arquivo geradas por sincronização do OneDrive
                # Prioriza o arquivo sem sufixo de cópia, ou o arquivo mais recente
                originais = [
                    f for f in lista_arqs
                    if f.stem == limpar_sufixos_conflito_onedrive(f.stem)
                ]
                escolhido = originais[0] if originais else max(lista_arqs, key=lambda f: f.stat().st_mtime)
                todos_docx.append(escolhido)
                duplicatas_onedrive_disco += (len(lista_arqs) - 1)
                logger.info(
                    "Conflito OneDrive no disco: %d cópia(s) redundante(s) ignorada(s) para '%s'. Utilizando '%s'.",
                    len(lista_arqs) - 1,
                    nome_can,
                    escolhido.name,
                )

        logger.info(
            "Total de arquivos encontrados na pasta: %d .docx (%d cópias OneDrive desconsideradas) e %d .pdf",
            len(todos_docx),
            duplicatas_onedrive_disco,
            len(todos_pdf),
        )

        # 3. CAMADA 1: Filtro por mtime (MAX(finalizada_em))
        corte_mtime = 0.0 if force_full else estado.obter_corte_camada1()
        if corte_mtime > 0:
            logger.info("Camada 1 ativa: ponto de corte mtime = %s", datetime.datetime.fromtimestamp(corte_mtime))
        else:
            logger.info("Camada 1 inativa (primeira execução ou --force-full). Processando lote integral.")

        candidatos_camada1: List[Path] = []
        ignorados_camada1 = 0

        for docx_path in todos_docx:
            try:
                st = docx_path.stat()
                if force_full or corte_mtime == 0.0 or st.st_mtime >= corte_mtime:
                    candidatos_camada1.append(docx_path)
                else:
                    ignorados_camada1 += 1
            except OSError:
                continue

        logger.info(
            "Camada 1 concluída: %d arquivos pré-selecionados (%d ignorados por data de modificação)",
            len(candidatos_camada1),
            ignorados_camada1,
        )

        # 4. CAMADA 2: Filtro persistente por tabela arquivos_processados
        arquivos_para_processar: List[Tuple[Path, str, os.stat_result]] = []
        total_ignorados = ignorados_camada1
        contagem_novos = 0
        contagem_atualizados = 0

        for docx_path in candidatos_camada1:
            try:
                st = docx_path.stat()
            except OSError as exc:
                logger.warning("Falha ao acessar stat() de '%s': %s", docx_path.name, exc)
                continue

            if force_full:
                arquivos_para_processar.append((docx_path, "FORCADO", st))
            else:
                status, reg = estado.verificar_camada2(
                    str(docx_path.resolve()),
                    st.st_size,
                    st.st_mtime,
                    basename=docx_path.name,
                )
                if status == "IGNORAR":
                    total_ignorados += 1
                elif status == "NOVO":
                    contagem_novos += 1
                    arquivos_para_processar.append((docx_path, "NOVO", st))
                elif status == "ATUALIZAR":
                    contagem_atualizados += 1
                    arquivos_para_processar.append((docx_path, "ATUALIZAR", st))

        logger.info(
            "Camada 2 concluída: %d a processar (Novos: %d, Alterados: %d, Ignorados: %d)",
            len(arquivos_para_processar),
            contagem_novos,
            contagem_atualizados,
            total_ignorados,
        )

        if dry_run:
            logger.info("--- MODO DRY-RUN: RELAÇÃO DOS ARQUIVOS QUE SERIAM PROCESSADOS ---")
            for docx_p, st_tipo, _ in arquivos_para_processar:
                logger.info("  [%s] %s", st_tipo, docx_p.name)
            logger.info("Dry-run finalizado com sucesso. Nenhuma alteração foi realizada em disco.")
            return 0

        # 5. Processamento dos arquivos selecionados
        lista_atos: List[RegistroAto] = []
        lista_erros: List[RegistroErro] = []
        remessas_detectadas: List[str] = []
        palavras_detectadas: List[str] = []
        total_erros = 0

        for docx_path, tipo_evento, st in arquivos_para_processar:
            try:
                # 5.1 Parsing do nome do arquivo
                parsed_nome = parse_filename(docx_path.name)
            except FilenameParseError as fpe:
                logger.warning("Erro de parsing de nome em '%s': %s", docx_path.name, fpe)
                lista_erros.append(RegistroErro(arquivo=docx_path.name, npu="", motivo=str(fpe)))
                total_erros += 1
                continue
            except Exception as exc:
                logger.error("Erro inesperado no nome de '%s': %s", docx_path.name, exc)
                lista_erros.append(RegistroErro(arquivo=docx_path.name, npu="", motivo=f"Erro inesperado: {exc}"))
                total_erros += 1
                continue

            # 5.2 Leitura do conteúdo do .docx (primeiras 50 linhas)
            linhas_docx = extrair_linhas_docx(docx_path, max_linhas=50)
            extracao_docx = parse_conteudo_docx(linhas_docx)

            # Validação NPU nome vs corpo
            npu_final = parsed_nome.npu
            if extracao_docx.npu_corpo and extracao_docx.npu_corpo != parsed_nome.npu:
                logger.warning(
                    "Divergência de NPU no arquivo '%s'! Nome: '%s' vs Conteúdo: '%s'. Priorizando conteúdo.",
                    docx_path.name,
                    parsed_nome.npu,
                    extracao_docx.npu_corpo,
                )
                lista_erros.append(
                    RegistroErro(
                        arquivo=docx_path.name,
                        npu=extracao_docx.npu_corpo,
                        motivo=f"NPU divergente: nome '{parsed_nome.npu}' vs corpo '{extracao_docx.npu_corpo}'",
                    )
                )
                npu_final = extracao_docx.npu_corpo

            # Desambiguação de tipo de ato caso necessário
            tipo_ato_final = parsed_nome.tipo_ato
            if tipo_ato_final == "Indefinido" and extracao_docx.cabecalho_ato:
                tipo_ato_final = classificar_tipo_ato(
                    parsed_nome.descricao_ato,
                    fallback_docx=extracao_docx.cabecalho_ato,
                )

            # 5.3 Pareamento com o PDF
            pdf_path = encontrar_pdf_pareado(docx_path, todos_pdf, npu_final)
            if pdf_path:
                texto_pdf = extrair_texto_pdf(pdf_path, max_paginas=3)
                meta_pdf = parse_metadados_processuais(texto_pdf)
                meta_pdf.caminho_pdf = pdf_path
            else:
                meta_pdf = parse_metadados_processuais("")
                meta_pdf.encontrado = False
                lista_erros.append(
                    RegistroErro(
                        arquivo=docx_path.name,
                        npu=npu_final,
                        motivo="Nenhum arquivo PDF correspondente localizado na pasta.",
                    )
                )

            # 5.4 Data do ato
            data_ato = formatar_data_mtime(st.st_mtime)
            if meta_pdf.data_assinatura:
                data_ato = meta_pdf.data_assinatura

            # 5.5 Montagem do RegistroAto (100% LGPD COMPLIANT: sem nomes de partes, CPF, RG, etc.)
            ato = RegistroAto(
                npu=npu_final,
                classe_judicial=meta_pdf.classe_judicial,
                assunto_principal=meta_pdf.assunto_principal,
                data_autuacao=meta_pdf.data_autuacao,
                valor_causa=meta_pdf.valor_causa,
                remessa=meta_pdf.remessa,
                tipo_ato=tipo_ato_final,
                data_ato=data_ato,
                data_conclusao=meta_pdf.data_conclusao,
                segredo_justica=meta_pdf.segredo_justica,
                palavras_chave=parsed_nome.palavras_chave_str,
                resultado_sentenca=parsed_nome.resultado_sentenca,
                descricao_ato_normalizada=parsed_nome.descricao_ato,
                caminho_origem=str(docx_path.resolve()),
            )
            lista_atos.append(ato)

            if meta_pdf.remessa and meta_pdf.remessa != "PENDENTE":
                remessas_detectadas.append(meta_pdf.remessa)

            for p in parsed_nome.palavras_chave_lista:
                if p:
                    palavras_detectadas.append(p)

            # Atualiza banco em memória
            estado.registrar_arquivo(
                caminho_str=str(docx_path.resolve()),
                basename=docx_path.name,
                tamanho=st.st_size,
                mtime=st.st_mtime,
                npu=npu_final,
                tipo_ato=tipo_ato_final,
            )

        # 6. CAMADA 3: Sincronização com Excel (dashboard_atos.xlsx)
        excel_writer = ExcelDashboardWriter(config.SAIDA_XLSX)
        stats_excel = excel_writer.sincronizar(
            novos_atos=lista_atos,
            todas_remessas=remessas_detectadas,
            todas_palavras=palavras_detectadas,
            erros=lista_erros,
            dry_run=dry_run,
        )

        # 6.1 GERAÇÃO DO DASHBOARD VISUAL INTERATIVO (dashboard_atos.html)
        # 6.1 GERAÇÃO DO DASHBOARD VISUAL INTERATIVO (dashboard_atos.html)
        if not dry_run:
            try:
                caminho_html = config.SAIDA_HTML
                atos_dicts = [
                    {
                        "npu": a.npu,
                        "classe_judicial": a.classe_judicial,
                        "assunto_principal": a.assunto_principal,
                        "data_autuacao": a.data_autuacao,
                        "valor_causa": a.valor_causa,
                        "remessa": a.remessa,
                        "tipo_ato": a.tipo_ato,
                        "data_ato": a.data_ato,
                        "data_conclusao": a.data_conclusao,
                        "segredo_justica": a.segredo_justica,
                        "palavras_chave": a.palavras_chave,
                        "resultado_sentenca": a.resultado_sentenca,
                        "descricao_ato": a.descricao_ato_normalizada,
                    }
                    for a in lista_atos
                ]
                gerar_dashboard_html(
                    caminho_html=caminho_html,
                    caminho_xlsx=config.SAIDA_XLSX,
                    atos_dados=atos_dicts,
                    erros=lista_erros,
                )
                logger.info("Dashboard visual HTML gerado em: %s", caminho_html)

                # Se a planilha estiver no OneDrive, copia para Downloads para acesso rápido local
                if config.COPIAR_PARA_DOWNLOADS and config.PASTA_DOWNLOADS != config.SAIDA_XLSX.parent:
                    import shutil
                    if config.SAIDA_XLSX.exists():
                        shutil.copy2(config.SAIDA_XLSX, config.PASTA_DOWNLOADS / config.SAIDA_XLSX.name)
                    if caminho_html.exists():
                        shutil.copy2(caminho_html, config.PASTA_DOWNLOADS / caminho_html.name)
                    logger.info("Cópia rápida de acesso sincronizada na pasta Downloads: %s", config.PASTA_DOWNLOADS)
            except Exception as exc_html:
                logger.warning("Aviso ao gerar dashboard HTML: %s", exc_html)

        # 7. Atualização do estado com remessas e métricas da execução
        estado.registrar_remessas(remessas_detectadas)

        fim_iso = datetime.datetime.now().isoformat()
        estado.registrar_execucao(
            iniciada_em=inicio_iso,
            finalizada_em=fim_iso,
            novos=stats_excel.get("adicionados", 0),
            atualizados=stats_excel.get("atualizados_pendentes", 0),
            ignorados=total_ignorados + stats_excel.get("inalterados", 0),
            erros=total_erros,
        )

        # 8. Persistência atômica e geração de backup datado
        caminho_salvo = estado.salvar_e_persistir()
        logger.info("Estado persistido com sucesso em: %s", caminho_salvo)

        logger.info("=" * 70)
        logger.info("CONSOLIDAÇÃO CONCLUÍDA COM SUCESSO!")
        logger.info(
            "Resumo: Novos adicionados: %d | Atualizados (pendências): %d | "
            "Ignorados: %d | Erros registrados: %d",
            stats_excel.get("adicionados", 0),
            stats_excel.get("atualizados_pendentes", 0),
            total_ignorados + stats_excel.get("inalterados", 0),
            total_erros,
        )
        logger.info("=" * 70)
        logger.info("ONDE ACESSAR SEUS ARQUIVOS CONSOLIDADOS:")
        logger.info("1. NO ONEDRIVE (Sincronizado na Nuvem - Acesse de Casa ou Trabalho):")
        logger.info("   -> Planilha Excel: %s", config.SAIDA_XLSX)
        logger.info("   -> Painel Web:     %s", config.SAIDA_HTML)
        logger.info("   -> Banco .SQL:     %s", config.REMOTO_SQL)
        if config.COPIAR_PARA_DOWNLOADS and config.PASTA_DOWNLOADS != config.SAIDA_XLSX.parent:
            logger.info("2. NA PASTA DOWNLOADS (Cópia Local de Acesso Rápido):")
            logger.info("   -> Planilha: %s", config.PASTA_DOWNLOADS / config.SAIDA_XLSX.name)
            logger.info("   -> Painel:   %s", config.PASTA_DOWNLOADS / "dashboard_atos.html")
        logger.info("=" * 70)
        return 0

    finally:
        if not dry_run:
            estado.liberar_lock()
        estado.fechar()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Consolidador incremental de Atos Judiciais - TJPE (Segmento 8.17)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Executa toda a lógica sem gravar alterações no .sql ou .xlsx.",
    )
    parser.add_argument(
        "--force-full",
        action="store_true",
        help="Reprocessa integralmente todos os .docx ignorando Camadas 1 e 2.",
    )
    parser.add_argument(
        "--pasta",
        type=Path,
        default=config.PASTA_DOWNLOADS,
        help="Caminho da pasta a monitorar (padrão: config.PASTA_DOWNLOADS).",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    sys.exit(executar(pasta_downloads=args.pasta, dry_run=args.dry_run, force_full=args.force_full))
