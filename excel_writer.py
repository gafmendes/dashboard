"""Gerador e atualizador da planilha dashboard_atos.xlsx.

Cria e sincroniza o arquivo Excel com:
  - Tabelas Estruturadas (ListObjects) com nomes padrão:
    * TabelaAtos (Aba 'Lançamento')
    * TabelaRemessas (Aba 'Remessas')
    * TabelaPalavras (Aba 'Palavras-chave')
    * TabelaListas (Aba 'Listas')
  - Fórmulas automáticas de Excel:
    * Data de chegada à Central: =IFERROR(VLOOKUP([@Remessa],TabelaRemessas,2,FALSE),"PENDENTE")
    * Tempo de tramitação: =[@[Data do ato]]-[@[Data de chegada à Central]]
    * Tempo parado: =[@[Data de chegada à Central]]-[@[Data de autuação]]
  - Aba Dashboard com resumos, fórmulas e gráficos nativos openpyxl
  - Aba LOG_ERROS (oculta)
  - Idempotência e Camada 3 de deduplicação (nunca sobrescreve dados manuais)
"""

from __future__ import annotations

from dataclasses import dataclass
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

logger = logging.getLogger("dashboard_atos.excel_writer")


@dataclass
class RegistroAto:
    npu: str
    classe_judicial: str
    assunto_principal: str
    data_autuacao: str
    valor_causa: str
    remessa: str
    tipo_ato: str
    data_ato: str
    data_conclusao: str
    segredo_justica: str
    palavras_chave: str
    resultado_sentenca: str
    descricao_ato_normalizada: str
    caminho_origem: str


@dataclass
class RegistroErro:
    arquivo: str
    npu: str
    motivo: str


COLUNAS_LANCAMENTO = [
    "NPU",
    "Classe judicial",
    "Assunto principal",
    "Data de autuação",
    "Valor da causa",
    "Remessa",
    "Data de chegada à Central",
    "Tipo de ato",
    "Data do ato",
    "Data de conclusão",
    "Tempo de elaboração (h)",
    "Tempo de tramitação até o ato (dias)",
    "Tempo parado na unidade jurisdicional (dias)",
    "Segredo de justiça",
    "Palavras-chave",
    "Resultado da sentença",
]

VALORES_TIPO_ATO = ["Despacho", "Decisão", "Sentença", "Indefinido"]
VALORES_RESULTADO_SENTENCA = [
    "Procedente",
    "Improcedente",
    "Parcialmente procedente",
    "Extinto sem mérito",
    "Homologação de acordo",
    "Negado provimento (ED)",
    "Provimento (ED)",
    "Não acolhimento (ED)",
    "Outro",
]
VALORES_SEGREDO = ["Sim", "Não"]


import re
from collections import OrderedDict

try:
    from parsers.filename_parser import (
        limpar_sufixos_conflito_onedrive,
        normalizar_para_comparacao,
    )
except ImportError:
    try:
        from dashboard_atos.parsers.filename_parser import (
            limpar_sufixos_conflito_onedrive,
            normalizar_para_comparacao,
        )
    except ImportError:
        def limpar_sufixos_conflito_onedrive(texto: str) -> str:
            t = re.sub(r"\s*\(\d+\)$", "", (texto or "").strip())
            t = re.sub(r"\s*-\s*C[oó]pia(\s*\(\d+\))?$", "", t, flags=re.IGNORECASE)
            t = re.sub(r"\s*-\s*Copy(\s*\(\d+\))?$", "", t, flags=re.IGNORECASE)
            t = re.sub(
                r"(-DESKTOP-[A-Z0-9]+|-LAPTOP-[A-Z0-9]+|-NOTEBOOK-[A-Z0-9]+|\s*-\s*[A-Za-z0-9_\-]+-PC)$",
                "",
                t,
                flags=re.IGNORECASE,
            )
            return t.strip()

        def normalizar_para_comparacao(texto: str) -> str:
            import unicodedata

            if not texto:
                return ""
            nfkd = unicodedata.normalize("NFKD", texto)
            return "".join(c for c in nfkd if not unicodedata.combining(c)).strip().upper()


def normalizar_chave_identidade_ato(
    npu: str,
    tipo_ato: str,
    resultado: str = "",
    palavras: str = "",
    descricao: str = "",
) -> Tuple[str, str, bool, str]:
    """Gera chave canônica de unicidade de um ato judicial à prova de conflitos do OneDrive.

    Retorna: (npu_limpo, tipo_ato_padronizado, is_ed, descricao_canonica)
    - Sentenças de mérito de um mesmo processo compartilham a mesma unicidade.
    - Decisões/Despachos com o mesmo objeto e NPU compartilham a mesma unicidade.
    - Sufixos do OneDrive como ' (1)', ' - Copia' e hostnames são limpos automaticamente.
    """
    npu_limpo = re.sub(r"[^\d\.\-]", "", npu or "").strip()
    tipo = (tipo_ato or "").strip().capitalize() if tipo_ato else "Indefinido"

    res_upper = (resultado or "").strip().upper()
    desc_upper = (descricao or "").strip().upper()
    pal_upper = (palavras or "").strip().upper()

    is_ed = (
        "(ED)" in res_upper
        or "EMBARGOS" in desc_upper
        or "EMBARGOS" in pal_upper
        or " ED " in f" {desc_upper} "
    )

    desc_limpa = limpar_sufixos_conflito_onedrive(descricao or palavras)
    desc_norm = normalizar_para_comparacao(desc_limpa)

    if tipo == "Sentença" and not is_ed:
        desc_canonica = "SENTENCA_MERITO"
    elif tipo == "Sentença" and is_ed:
        desc_canonica = "SENTENCA_ED"
    else:
        desc_canonica = desc_norm

    return (npu_limpo, tipo, is_ed, desc_canonica)


def chave_unicidade_ato(reg: RegistroAto) -> Tuple[str, str, str, str, str]:
    """Chave da Camada 3 para compatibilidade e testes.

    Normaliza sufixos de cópia e conflitos do OneDrive na descrição.
    """
    desc_limpa = limpar_sufixos_conflito_onedrive(reg.descricao_ato_normalizada).strip().upper()
    return (
        reg.npu.strip(),
        reg.tipo_ato.strip(),
        reg.data_ato.strip(),
        desc_limpa,
        reg.resultado_sentenca.strip(),
    )


class ExcelDashboardWriter:
    """Gerencia a criação e atualização incremental do dashboard_atos.xlsx."""

    def __init__(self, caminho_xlsx: Path):
        self.caminho_xlsx = Path(caminho_xlsx)

    def sincronizar(
        self,
        novos_atos: List[RegistroAto],
        todas_remessas: List[str],
        todas_palavras: List[str],
        erros: List[RegistroErro],
        dry_run: bool = False,
    ) -> Dict[str, int]:
        """Aplica novos atos e remessas à planilha respeitando a Camada 3 de deduplicação."""
        try:
            import openpyxl  # type: ignore
            from openpyxl.chart import BarChart, Reference  # type: ignore
            from openpyxl.comments import Comment  # type: ignore
            from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # type: ignore
            from openpyxl.worksheet.table import Table, TableStyleInfo  # type: ignore
        except ImportError:
            logger.warning("openpyxl não instalado. Modo de simulação ativa.")
            return {
                "adicionados": len(novos_atos),
                "atualizados_pendentes": 0,
                "inalterados": 0,
                "total_erros": len(erros),
            }

        # Carrega pasta de trabalho existente ou cria nova
        if self.caminho_xlsx.exists():
            wb = openpyxl.load_workbook(str(self.caminho_xlsx))
            logger.info("Planilha existente carregada: %s", self.caminho_xlsx)
        else:
            wb = openpyxl.Workbook()
            # Remove aba padrão 'Sheet'
            wb.remove(wb.active)
            logger.info("Criando nova planilha estruturada: %s", self.caminho_xlsx)

        # 1. Garante existência das abas na ordem estrita solicitada
        ordem_abas = ["Lançamento", "Remessas", "Palavras-chave", "Listas", "Dashboard", "LOG_ERROS"]
        for nome_aba in ordem_abas:
            if nome_aba not in wb.sheetnames:
                wb.create_sheet(title=nome_aba)

        # Ordenar abas no workbook
        wb._sheets = [wb[nome] for nome in ordem_abas if nome in wb.sheetnames]

        ws_lancamento = wb["Lançamento"]
        ws_remessas = wb["Remessas"]
        ws_palavras = wb["Palavras-chave"]
        ws_listas = wb["Listas"]
        ws_dashboard = wb["Dashboard"]
        ws_erros = wb["LOG_ERROS"]

        # 2. Configura Aba "Listas"
        self._configurar_aba_listas(ws_listas)

        # 3. Configura e atualiza Aba "Remessas" (TabelaRemessas)
        self._sincronizar_aba_remessas(ws_remessas, todas_remessas)

        # 4. Sincroniza Aba "Lançamento" (TabelaAtos) com Camada 3
        stats_lancamento = self._sincronizar_aba_lancamento(ws_lancamento, novos_atos)

        # 5. Sincroniza Aba "Palavras-chave" (TabelaPalavras)
        self._sincronizar_aba_palavras(ws_palavras, todas_palavras)

        # 6. Sincroniza Aba "Dashboard" com tabelas consolidadas e gráficos nativos
        self._configurar_aba_dashboard(ws_dashboard, ws_lancamento)

        # 7. Sincroniza Aba "LOG_ERROS" (Oculta)
        self._sincronizar_aba_erros(ws_erros, erros)
        ws_erros.sheet_state = "hidden"

        # 8. Define a aba Dashboard como ATIVA para abrir direto nela ao dar duplo-clique no Excel
        wb.active = ws_dashboard
        for sheet in wb.worksheets:
            if hasattr(sheet, "sheet_view"):
                sheet.sheet_view.tabSelected = (sheet == ws_dashboard)

        if not dry_run:
            self.caminho_xlsx.parent.mkdir(parents=True, exist_ok=True)
            wb.save(str(self.caminho_xlsx))
            logger.info("Planilha salva com sucesso em: %s", self.caminho_xlsx)

        return stats_lancamento

    # -------------------------------------------------------------------------
    # Métodos Internos de Estruturação e Fórmulas
    # -------------------------------------------------------------------------
    def _configurar_aba_listas(self, ws: Any) -> None:
        """Preenche a aba de domínio com as listas padronizadas."""
        from openpyxl.worksheet.table import Table, TableStyleInfo  # type: ignore

        if ws.max_row <= 1:
            ws.cell(row=1, column=1, value="Tipo de ato")
            ws.cell(row=1, column=2, value="Resultado da sentença")
            ws.cell(row=1, column=3, value="Segredo de justiça")

            max_len = max(len(VALORES_TIPO_ATO), len(VALORES_RESULTADO_SENTENCA), len(VALORES_SEGREDO))
            for i in range(max_len):
                row = i + 2
                ws.cell(row=row, column=1, value=VALORES_TIPO_ATO[i] if i < len(VALORES_TIPO_ATO) else "")
                ws.cell(
                    row=row,
                    column=2,
                    value=VALORES_RESULTADO_SENTENCA[i] if i < len(VALORES_RESULTADO_SENTENCA) else "",
                )
                ws.cell(row=row, column=3, value=VALORES_SEGREDO[i] if i < len(VALORES_SEGREDO) else "")

            tab = Table(displayName="TabelaListas", ref=f"A1:C{max_len + 1}")
            style = TableStyleInfo(name="TableStyleLight1", showFirstColumn=False, showLastColumn=False, showRowStripes=True)
            tab.tableStyleInfo = style
            ws.add_table(tab)

    def _sincronizar_aba_remessas(self, ws: Any, novas_remessas: List[str]) -> None:
        """Mantém a TabelaRemessas sem sobrescrever dados já preenchidos pelo usuário."""
        from openpyxl.worksheet.table import Table, TableStyleInfo  # type: ignore

        remessas_existentes: Dict[str, Tuple[int, Any]] = {}
        if ws.max_row > 1:
            for r in range(2, ws.max_row + 1):
                val_remessa = ws.cell(row=r, column=1).value
                val_data = ws.cell(row=r, column=2).value
                if val_remessa:
                    remessas_existentes[str(val_remessa).strip()] = (r, val_data)

        # Cabeçalho se nova
        if ws.max_row <= 1 or ws.cell(row=1, column=1).value != "Remessa":
            ws.cell(row=1, column=1, value="Remessa")
            ws.cell(row=1, column=2, value="Data de chegada à Central")

        # Adiciona apenas as não existentes
        todas_unicas = sorted(
            list(set(remessas_existentes.keys()) | {r for r in novas_remessas if r and r != "PENDENTE"})
        )

        # Re-escreve de forma ordenada preservando datas manuais
        row_idx = 2
        for r_nome in todas_unicas:
            data_anterior = remessas_existentes.get(r_nome, (0, ""))[1]
            ws.cell(row=row_idx, column=1, value=r_nome)
            ws.cell(row=row_idx, column=2, value=data_anterior or "")
            row_idx += 1

        total_rows = max(row_idx - 1, 2)
        # Atualiza ou cria a tabela estruturada
        if "TabelaRemessas" in ws.tables:
            ws.tables["TabelaRemessas"].ref = f"A1:B{total_rows}"
        else:
            tab = Table(displayName="TabelaRemessas", ref=f"A1:B{total_rows}")
            tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
            ws.add_table(tab)

    def _sincronizar_aba_lancamento(self, ws: Any, novos_atos: List[RegistroAto]) -> Dict[str, int]:
        """Aplica a Camada 3 de deduplicação e insere/atualiza registros na TabelaAtos.

        Garante que:
          1. Registros duplicados existentes na planilha (gerados por conflitos do OneDrive) sejam unificados.
          2. Novos atos correspondentes a atos já existentes apenas atualizem campos 'PENDENTE', preservando edições manuais.
          3. Não ocorram duplicidades quando arquivos forem renomeados com sufixos '(1)', '- Copia' ou variações de data.
        """
        from openpyxl.worksheet.table import Table, TableStyleInfo  # type: ignore

        # Garante cabeçalho
        if ws.max_row <= 1 or ws.cell(row=1, column=1).value != COLUNAS_LANCAMENTO[0]:
            for col_idx, col_nome in enumerate(COLUNAS_LANCAMENTO, start=1):
                ws.cell(row=1, column=col_idx, value=col_nome)

        # 1. Carrega todos os registros atuais da planilha
        # Dicionário ordenado para preservar a ordem das linhas na planilha
        registros_consolidados: "OrderedDict[Tuple[str, str, bool, str], Dict[str, Any]]" = OrderedDict()
        duplicatas_existentes_purgadas = 0

        if ws.max_row > 1:
            for r in range(2, ws.max_row + 1):
                npu_val = str(ws.cell(row=r, column=1).value or "").strip()
                if not npu_val or npu_val == "None":
                    continue

                tipo_val = str(ws.cell(row=r, column=8).value or "").strip()
                data_ato_val = str(ws.cell(row=r, column=9).value or "").strip()
                palavras_val = str(ws.cell(row=r, column=15).value or "").strip()
                resultado_val = str(ws.cell(row=r, column=16).value or "").strip()

                key = normalizar_chave_identidade_ato(
                    npu=npu_val,
                    tipo_ato=tipo_val,
                    resultado=resultado_val,
                    palavras=palavras_val,
                )

                dados_linha = {
                    "npu": npu_val,
                    "classe_judicial": str(ws.cell(row=r, column=2).value or "").strip() or "PENDENTE",
                    "assunto_principal": str(ws.cell(row=r, column=3).value or "").strip() or "PENDENTE",
                    "data_autuacao": str(ws.cell(row=r, column=4).value or "").strip() or "PENDENTE",
                    "valor_causa": str(ws.cell(row=r, column=5).value or "").strip() or "PENDENTE",
                    "remessa": str(ws.cell(row=r, column=6).value or "").strip() or "PENDENTE",
                    "tipo_ato": tipo_val or "Indefinido",
                    "data_ato": data_ato_val,
                    "data_conclusao": str(ws.cell(row=r, column=10).value or "").strip(),
                    "tempo_elaboracao": ws.cell(row=r, column=11).value,
                    "segredo_justica": str(ws.cell(row=r, column=14).value or "").strip() or "Não",
                    "palavras_chave": palavras_val,
                    "resultado_sentenca": resultado_val or "Outro",
                }

                if key in registros_consolidados:
                    # Linha duplicada já existente na planilha: mescla dados preservando valores preenchidos
                    duplicatas_existentes_purgadas += 1
                    existente = registros_consolidados[key]
                    for campo in ["classe_judicial", "assunto_principal", "data_autuacao", "valor_causa", "remessa"]:
                        if existente[campo] in ("PENDENTE", "") and dados_linha[campo] not in ("PENDENTE", ""):
                            existente[campo] = dados_linha[campo]
                    if not existente.get("tempo_elaboracao") and dados_linha.get("tempo_elaboracao"):
                        existente["tempo_elaboracao"] = dados_linha["tempo_elaboracao"]
                    if not existente.get("data_conclusao") and dados_linha.get("data_conclusao"):
                        existente["data_conclusao"] = dados_linha["data_conclusao"]
                else:
                    registros_consolidados[key] = dados_linha

        if duplicatas_existentes_purgadas > 0:
            logger.info(
                "Deduplicação de planilha ativa: %d linha(s) duplicada(s) detectadas e unificadas com sucesso.",
                duplicatas_existentes_purgadas,
            )

        adicionados = 0
        atualizados_pendentes = 0
        inalterados = 0

        # 2. Processa os novos atos vindos da execução atual
        for ato in novos_atos:
            chave_ato = normalizar_chave_identidade_ato(
                npu=ato.npu,
                tipo_ato=ato.tipo_ato,
                resultado=ato.resultado_sentenca,
                palavras=ato.palavras_chave,
                descricao=ato.descricao_ato_normalizada,
            )

            if chave_ato in registros_consolidados:
                reg = registros_consolidados[chave_ato]
                modificado = False
                for campo, valor in [
                    ("classe_judicial", ato.classe_judicial),
                    ("assunto_principal", ato.assunto_principal),
                    ("data_autuacao", ato.data_autuacao),
                    ("valor_causa", ato.valor_causa),
                    ("remessa", ato.remessa),
                ]:
                    if reg[campo] in ("PENDENTE", "") and valor not in ("PENDENTE", ""):
                        reg[campo] = valor
                        modificado = True

                if not reg.get("data_conclusao") and ato.data_conclusao:
                    reg["data_conclusao"] = ato.data_conclusao
                    modificado = True

                if reg.get("tipo_ato") in ("Indefinido", "") and ato.tipo_ato not in ("Indefinido", ""):
                    reg["tipo_ato"] = ato.tipo_ato
                    modificado = True

                if modificado:
                    atualizados_pendentes += 1
                else:
                    inalterados += 1
            else:
                # Novo registro legítimo
                registros_consolidados[chave_ato] = {
                    "npu": ato.npu,
                    "classe_judicial": ato.classe_judicial,
                    "assunto_principal": ato.assunto_principal,
                    "data_autuacao": ato.data_autuacao,
                    "valor_causa": ato.valor_causa,
                    "remessa": ato.remessa,
                    "tipo_ato": ato.tipo_ato,
                    "data_ato": ato.data_ato,
                    "data_conclusao": ato.data_conclusao,
                    "tempo_elaboracao": "",
                    "segredo_justica": ato.segredo_justica,
                    "palavras_chave": ato.palavras_chave,
                    "resultado_sentenca": ato.resultado_sentenca,
                }
                adicionados += 1

        # 3. Reescreve as linhas consolidadas na planilha
        row_idx = 2
        for key, reg in registros_consolidados.items():
            ws.cell(row=row_idx, column=1, value=reg["npu"])
            ws.cell(row=row_idx, column=2, value=reg["classe_judicial"])
            ws.cell(row=row_idx, column=3, value=reg["assunto_principal"])
            ws.cell(row=row_idx, column=4, value=reg["data_autuacao"])
            ws.cell(row=row_idx, column=5, value=reg["valor_causa"])
            ws.cell(row=row_idx, column=6, value=reg["remessa"])

            # Coluna 7: Fórmula estruturada de chegada à Central
            ws.cell(
                row=row_idx,
                column=7,
                value='=IFERROR(VLOOKUP([@Remessa],TabelaRemessas,2,FALSE),"PENDENTE")',
            )

            ws.cell(row=row_idx, column=8, value=reg["tipo_ato"])
            ws.cell(row=row_idx, column=9, value=reg["data_ato"])
            ws.cell(row=row_idx, column=10, value=reg["data_conclusao"])
            ws.cell(row=row_idx, column=11, value=reg.get("tempo_elaboracao") or "")

            # Coluna 12: Tempo de tramitação até o ato (dias)
            ws.cell(
                row=row_idx,
                column=12,
                value="=[@[Data do ato]]-[@[Data de chegada à Central]]",
            )

            # Coluna 13: Tempo parado na unidade jurisdicional (dias)
            ws.cell(
                row=row_idx,
                column=13,
                value="=[@[Data de chegada à Central]]-[@[Data de autuação]]",
            )

            ws.cell(row=row_idx, column=14, value=reg["segredo_justica"])
            ws.cell(row=row_idx, column=15, value=reg["palavras_chave"])
            ws.cell(row=row_idx, column=16, value=reg["resultado_sentenca"])

            row_idx += 1

        # 4. Limpa quaisquer linhas excedentes anteriores (se haviam duplicatas apagadas)
        if ws.max_row >= row_idx:
            for r in range(row_idx, ws.max_row + 1):
                for c in range(1, len(COLUNAS_LANCAMENTO) + 1):
                    ws.cell(row=r, column=c, value=None)

        total_linhas = max(row_idx - 1, 2)
        ref_table = f"A1:P{total_linhas}"
        if "TabelaAtos" in ws.tables:
            ws.tables["TabelaAtos"].ref = ref_table
        else:
            tab = Table(displayName="TabelaAtos", ref=ref_table)
            tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
            ws.add_table(tab)

        return {
            "adicionados": adicionados,
            "atualizados_pendentes": atualizados_pendentes,
            "inalterados": inalterados,
            "duplicatas_purgadas": duplicatas_existentes_purgadas,
        }

    def _sincronizar_aba_palavras(self, ws: Any, todas_palavras: List[str]) -> None:
        """Mantém a lista ordenada de palavras-chave únicas (TabelaPalavras)."""
        from openpyxl.worksheet.table import Table, TableStyleInfo  # type: ignore

        if ws.max_row <= 1 or ws.cell(row=1, column=1).value != "Palavra":
            ws.cell(row=1, column=1, value="Palavra")

        existentes: Set[str] = set()
        if ws.max_row > 1:
            for r in range(2, ws.max_row + 1):
                val = ws.cell(row=r, column=1).value
                if val:
                    existentes.add(str(val).strip())

        combinadas = sorted(list(existentes | {p.strip() for p in todas_palavras if p and p.strip()}))

        row_idx = 2
        for pal in combinadas:
            ws.cell(row=row_idx, column=1, value=pal)
            row_idx += 1

        total_linhas = max(row_idx - 1, 2)
        ref_tab = f"A1:A{total_linhas}"
        if "TabelaPalavras" in ws.tables:
            ws.tables["TabelaPalavras"].ref = ref_tab
        else:
            tab = Table(displayName="TabelaPalavras", ref=ref_tab)
            tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight8", showRowStripes=True)
            ws.add_table(tab)

    def _configurar_aba_dashboard(self, ws: Any, ws_lancamento: Optional[Any] = None) -> None:
        """Configura a aba Dashboard com cartões KPI, tabelas de resumo e gráficos nativos do Excel."""
        from collections import Counter
        from openpyxl.chart import BarChart, Reference  # type: ignore
        from openpyxl.comments import Comment  # type: ignore
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # type: ignore

        # 1. Coleta métricas reais a partir da aba Lançamento (TabelaAtos)
        total_atos = 0
        contagem_tipo: Counter[str] = Counter()
        contagem_resultado: Counter[str] = Counter()
        contagem_remessas: Counter[str] = Counter()
        contagem_ed = 0

        if ws_lancamento is not None and ws_lancamento.max_row > 1:
            for r in range(2, ws_lancamento.max_row + 1):
                npu = ws_lancamento.cell(row=r, column=1).value
                if not npu:
                    continue
                total_atos += 1
                tipo = str(ws_lancamento.cell(row=r, column=8).value or "Indefinido").strip()
                contagem_tipo[tipo] += 1

                res = str(ws_lancamento.cell(row=r, column=16).value or "").strip()
                if res:
                    contagem_resultado[res] += 1

                rem = str(ws_lancamento.cell(row=r, column=6).value or "").strip()
                if rem and rem != "PENDENTE":
                    contagem_remessas[rem] += 1

                palavras = str(ws_lancamento.cell(row=r, column=15).value or "").upper()
                if "EMBARGOS" in palavras or "ED" in palavras or "(ED)" in res:
                    contagem_ed += 1

        # Garante visibilidade da grade
        if hasattr(ws, "views") and ws.views.sheetView:
            ws.views.sheetView[0].showGridLines = True

        # Estilos padrão
        font_titulo = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
        font_subtitulo = Font(name="Segoe UI", size=9, italic=True, color="E2E8F0")
        fill_banner = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
        fill_sub_banner = PatternFill(start_color="2C4C7E", end_color="2C4C7E", fill_type="solid")
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

        border_thin = Side(border_style="thin", color="CBD5E1")
        card_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

        # 2. Banner Superior
        ws.merge_cells("A1:L1")
        cell_a1 = ws.cell(row=1, column=1, value="PAINEL DE GESTÃO DE ATOS JUDICIAIS - TJPE (8.17)")
        cell_a1.font = font_titulo
        cell_a1.fill = fill_banner
        cell_a1.alignment = align_center
        ws.row_dimensions[1].height = 28

        ws.merge_cells("A2:L2")
        cell_a2 = ws.cell(row=2, column=1, value="Consolidação Automatizada • Assessoria de Magistrado • 100% LGPD Safe")
        cell_a2.font = font_subtitulo
        cell_a2.fill = fill_sub_banner
        cell_a2.alignment = align_center
        ws.row_dimensions[2].height = 18

        # Comentário de apoio em A1
        ws.cell(row=1, column=1).comment = Comment(
            "PAINEL CONSOLIDADO DO ASSESSOR:\n"
            "- Cartões de KPI e gráficos refletem automaticamente os dados da aba 'Lançamento'.\n"
            "- Novos atos adicionados são incorporados a cada execução do consolidador.",
            "Dashboard Atos TJPE",
        )

        # 3. Cartões KPI (Linhas 4 a 6)
        kpi_defs = [
            ("TOTAL DE ATOS", total_atos, '=COUNTA(TabelaAtos[NPU])', "1E293B", "F8FAFC", 1, 2),
            ("DESPACHOS", contagem_tipo.get("Despacho", 0), '=COUNTIF(TabelaAtos[Tipo de ato],"Despacho")', "1E40AF", "EFF6FF", 3, 4),
            ("DECISÕES", contagem_tipo.get("Decisão", 0), '=COUNTIF(TabelaAtos[Tipo de ato],"Decisão")', "B45309", "FFFBEB", 5, 6),
            ("SENTENÇAS", contagem_tipo.get("Sentença", 0), '=COUNTIF(TabelaAtos[Tipo de ato],"Sentença")', "047857", "ECFDF5", 7, 8),
            ("EMBARGOS (ED)", contagem_ed, contagem_ed, "6D28D9", "F5F3FF", 9, 10),
        ]

        ws.row_dimensions[4].height = 16
        ws.row_dimensions[5].height = 30

        for titulo, valor_estatico, formula_ou_val, cor_header, cor_fundo, col_inicio, col_fim in kpi_defs:
            # Cabeçalho do Card
            ws.merge_cells(start_row=4, start_column=col_inicio, end_row=4, end_column=col_fim)
            c_header = ws.cell(row=4, column=col_inicio, value=titulo)
            c_header.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
            c_header.fill = PatternFill(start_color=cor_header, end_color=cor_header, fill_type="solid")
            c_header.alignment = align_center

            # Valor do Card
            ws.merge_cells(start_row=5, start_column=col_inicio, end_row=6, end_column=col_fim)
            c_val = ws.cell(row=5, column=col_inicio)
            if isinstance(formula_ou_val, str) and formula_ou_val.startswith("="):
                c_val.value = formula_ou_val
            else:
                c_val.value = valor_estatico

            c_val.font = Font(name="Segoe UI", size=20, bold=True, color=cor_header)
            c_val.fill = PatternFill(start_color=cor_fundo, end_color=cor_fundo, fill_type="solid")
            c_val.alignment = align_center

            # Aplica bordas nos cantos do cartão
            for r in range(4, 7):
                for c in range(col_inicio, col_fim + 1):
                    ws.cell(row=r, column=c).border = card_border

        # 4. Tabela Resumo 1: Distribuição por Tipo de Ato (A8:C13)
        ws.cell(row=8, column=1, value="Tipo de Ato").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=8, column=1).fill = fill_banner
        ws.cell(row=8, column=2, value="Quantidade").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=8, column=2).fill = fill_banner
        ws.cell(row=8, column=3, value="% do Total").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=8, column=3).fill = fill_banner

        tipos = ["Despacho", "Decisão", "Sentença", "Indefinido"]
        for idx, t in enumerate(tipos, start=9):
            ws.cell(row=idx, column=1, value=t)
            ws.cell(row=idx, column=2, value=f'=COUNTIF(TabelaAtos[Tipo de ato], "{t}")')
            ws.cell(row=idx, column=3, value=f'=B{idx}/MAX($A$5, 1)')
            ws.cell(row=idx, column=3).number_format = "0.0%"

        # 5. Tabela Resumo 2: Resultados de Sentença (E8:G17)
        ws.cell(row=8, column=5, value="Resultado da Sentença").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=8, column=5).fill = fill_banner
        ws.cell(row=8, column=6, value="Quantidade").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=8, column=6).fill = fill_banner
        ws.cell(row=8, column=7, value="% Sentenças").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=8, column=7).fill = fill_banner

        resultados = [
            "Procedente",
            "Parcialmente procedente",
            "Improcedente",
            "Extinto sem mérito",
            "Homologação de acordo",
            "Provimento (ED)",
            "Negado provimento (ED)",
            "Não acolhimento (ED)",
            "Outro",
        ]
        for idx, res in enumerate(resultados, start=9):
            ws.cell(row=idx, column=5, value=res)
            ws.cell(row=idx, column=6, value=f'=COUNTIF(TabelaAtos[Resultado da sentença], "{res}")')
            ws.cell(row=idx, column=7, value=f'=F{idx}/MAX($G$5, 1)')
            ws.cell(row=idx, column=7).number_format = "0.0%"

        # 6. Tabela Resumo 3: Top Remessas (I8:K17)
        ws.cell(row=8, column=9, value="Top Remessas").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=8, column=9).fill = fill_banner
        ws.cell(row=8, column=10, value="Quantidade").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=8, column=10).fill = fill_banner
        ws.cell(row=8, column=11, value="% Total").font = Font(bold=True, color="FFFFFF")
        ws.cell(row=8, column=11).fill = fill_banner

        top_rem = contagem_remessas.most_common(9)
        for idx in range(9):
            row_num = 9 + idx
            if idx < len(top_rem):
                nome_rem, qtd_rem = top_rem[idx]
                ws.cell(row=row_num, column=9, value=nome_rem)
                ws.cell(row=row_num, column=10, value=f'=COUNTIF(TabelaAtos[Remessa], "{nome_rem}")')
                ws.cell(row=row_num, column=11, value=f'=J{row_num}/MAX($A$5, 1)')
            else:
                ws.cell(row=row_num, column=9, value="-")
                ws.cell(row=row_num, column=10, value=0)
                ws.cell(row=row_num, column=11, value=0)
            ws.cell(row=row_num, column=11).number_format = "0.0%"

        # 7. Inserção de Gráficos Nativos do Excel
        try:
            # Gráfico 1: Tipo de Ato (Colunas)
            chart_tipos = BarChart()
            chart_tipos.type = "col"
            chart_tipos.style = 10
            chart_tipos.title = "Distribuição de Atos por Tipo"
            chart_tipos.y_axis.title = "Quantidade"
            chart_tipos.x_axis.title = "Tipo de Ato"
            chart_tipos.width = 14
            chart_tipos.height = 10
            data_t = Reference(ws, min_col=2, min_row=8, max_row=12)
            cats_t = Reference(ws, min_col=1, min_row=9, max_row=12)
            chart_tipos.add_data(data_t, titles_from_data=True)
            chart_tipos.set_categories(cats_t)
            chart_tipos.legend = None
            ws.add_chart(chart_tipos, "A19")

            # Gráfico 2: Resultados de Sentença (Barras Horizontais)
            chart_res = BarChart()
            chart_res.type = "bar"
            chart_res.style = 11
            chart_res.title = "Resultados de Sentenças & Embargos"
            chart_res.x_axis.title = "Quantidade"
            chart_res.y_axis.title = "Resultado"
            chart_res.width = 17
            chart_res.height = 10
            data_r = Reference(ws, min_col=6, min_row=8, max_row=17)
            cats_r = Reference(ws, min_col=5, min_row=9, max_row=17)
            chart_res.add_data(data_r, titles_from_data=True)
            chart_res.set_categories(cats_r)
            chart_res.legend = None
            ws.add_chart(chart_res, "E19")

            # Gráfico 3: Top Remessas
            chart_rem = BarChart()
            chart_rem.type = "col"
            chart_rem.style = 13
            chart_rem.title = "Top Remessas Atendidas"
            chart_rem.y_axis.title = "Quantidade"
            chart_rem.x_axis.title = "Remessa"
            chart_rem.width = 16
            chart_rem.height = 10
            data_m = Reference(ws, min_col=10, min_row=8, max_row=17)
            cats_m = Reference(ws, min_col=9, min_row=9, max_row=17)
            chart_rem.add_data(data_m, titles_from_data=True)
            chart_rem.set_categories(cats_m)
            chart_rem.legend = None
            ws.add_chart(chart_rem, "I19")
        except Exception as exc:
            logger.warning("Aviso ao gerar gráficos nativos na aba Dashboard: %s", exc)

    def _sincronizar_aba_erros(self, ws: Any, erros: List[RegistroErro]) -> None:
        """Registra logs de divergências e arquivos problemáticos."""
        if ws.max_row <= 1:
            ws.cell(row=1, column=1, value="Arquivo")
            ws.cell(row=1, column=2, value="NPU")
            ws.cell(row=1, column=3, value="Motivo")

        start_row = ws.max_row + 1
        for e in erros:
            ws.cell(row=start_row, column=1, value=e.arquivo)
            ws.cell(row=start_row, column=2, value=e.npu)
            ws.cell(row=start_row, column=3, value=e.motivo)
            start_row += 1
