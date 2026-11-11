"""Módulo de persistência de estado remoto (Arquitetura B).

O estado é mantido em um único arquivo .sql remoto que é carregado em um
banco SQLite em memória a cada execução, manipulado de forma volátil e depois
serializado atomicamente de volta para o .sql com backup datado e lock com TTL.
"""

from __future__ import annotations

import datetime
import json
import logging
import os
from pathlib import Path
import platform
import sqlite3
import time
from typing import Optional, Tuple

logger = logging.getLogger("dashboard_atos.estado")

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS arquivos_processados (
    caminho TEXT PRIMARY KEY,
    basename TEXT,
    tamanho INTEGER,
    mtime REAL,
    processado_em TEXT,
    npu TEXT,
    tipo_ato TEXT
);

CREATE INDEX IF NOT EXISTS idx_arquivos_npu ON arquivos_processados(npu);

CREATE TABLE IF NOT EXISTS execucoes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    iniciada_em TEXT,
    finalizada_em TEXT,
    novos INTEGER,
    atualizados INTEGER,
    ignorados INTEGER,
    erros INTEGER,
    maquina TEXT
);

CREATE TABLE IF NOT EXISTS remessas (
    remessa TEXT PRIMARY KEY,
    data_chegada TEXT
);
"""


class LockError(Exception):
    """Exceção levantada quando outro processo detém o lock do arquivo de estado."""
    pass


class EstadoRemoto:
    """Gerenciador do estado SQLite em memória serializado em arquivo .sql remoto."""

    def __init__(
        self,
        remoto_sql: Path,
        backup_dir: Path,
        lock_ttl_minutos: int = 30,
        retry_replace: int = 3,
        retry_sleep_seg: float = 2.0,
    ):
        self.remoto_sql = Path(remoto_sql)
        self.backup_dir = Path(backup_dir)
        self.lock_ttl_seg = lock_ttl_minutos * 60
        self.retry_replace = retry_replace
        self.retry_sleep_seg = retry_sleep_seg

        self.lock_file = self.remoto_sql.parent / f"{self.remoto_sql.stem}.lock"
        self.conn: Optional[sqlite3.Connection] = None
        self._lock_adquirido = False

    def adquirir_lock(self) -> None:
        """Adquire o arquivo .lock com verificação de TTL (30 min)."""
        try:
            self.remoto_sql.parent.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            raise RuntimeError(
                f"Acesso negado ou caminho inválido ao tentar criar a pasta de estado: '{self.remoto_sql.parent}'. "
                f"Abra o arquivo 'config.py' e ajuste a variável REMOTO_SQL para um caminho existente no seu computador "
                f"(ex: C:\\Users\\Gabriel Mendes\\OneDrive - TJPE\\dashboard\\estado.sql). Detalhes do erro: {exc}"
            ) from exc

        agora = time.time()

        if self.lock_file.exists():
            try:
                mtime_lock = self.lock_file.stat().st_mtime
                idade_lock = agora - mtime_lock
                if idade_lock < self.lock_ttl_seg:
                    # Lock ainda é válido e ativo
                    try:
                        conteudo = self.lock_file.read_text(encoding="utf-8")
                    except Exception:
                        conteudo = "(indisponível)"
                    raise LockError(
                        f"Execução bloqueada: o arquivo de lock '{self.lock_file.name}' está ativo "
                        f"(idade: {int(idade_lock)}s / limite: {self.lock_ttl_seg}s). "
                        f"Detalhes do proprietário: {conteudo}"
                    )
                else:
                    logger.warning(
                        "Lock residual detectado (idade %d s > limite %d s). Sobrescrevendo...",
                        int(idade_lock),
                        self.lock_ttl_seg,
                    )
            except OSError as exc:
                logger.warning("Falha ao checar lock existente: %s. Tentando prosseguir.", exc)

        info_lock = {
            "pid": os.getpid(),
            "maquina": platform.node(),
            "iniciado_em": datetime.datetime.now().isoformat(),
            "timestamp": agora,
        }

        try:
            self.lock_file.write_text(json.dumps(info_lock, indent=2), encoding="utf-8")
            self._lock_adquirido = True
            logger.info("Lock adquirido com sucesso: %s", self.lock_file)
        except OSError as exc:
            raise LockError(f"Não foi possível gravar o arquivo de lock '{self.lock_file}': {exc}") from exc

    def liberar_lock(self) -> None:
        """Libera o arquivo .lock se adquirido por esta instância."""
        if not self._lock_adquirido:
            return
        try:
            if self.lock_file.exists():
                self.lock_file.unlink()
                logger.info("Lock liberado com sucesso: %s", self.lock_file)
        except OSError as exc:
            logger.warning("Aviso ao remover arquivo de lock '%s': %s", self.lock_file, exc)
        finally:
            self._lock_adquirido = False

    # -------------------------------------------------------------------------
    # Ciclo de Vida: Carregamento do .sql -> SQLite em memória
    # -------------------------------------------------------------------------
    def carregar(self) -> sqlite3.Connection:
        """Carrega ou inicializa o banco SQLite em memória a partir do .sql remoto."""
        self.conn = sqlite3.connect(":memory:")
        self.conn.row_factory = sqlite3.Row

        if self.remoto_sql.exists():
            logger.info("Lendo estado persistente de: %s", self.remoto_sql)
            try:
                sql_script = self.remoto_sql.read_text(encoding="utf-8")
                self.conn.executescript(sql_script)
                logger.info("Estado remoto restaurado com sucesso na memória.")
            except Exception as exc:
                raise RuntimeError(
                    f"Falha ao restaurar banco em memória do arquivo '{self.remoto_sql}': {exc}"
                ) from exc
        else:
            logger.info("Arquivo de estado '%s' não encontrado. Criando schema inicial.", self.remoto_sql)
            self.conn.executescript(SCHEMA_SQL)
            self.conn.commit()

        # Garante integridade do schema
        self.conn.executescript(SCHEMA_SQL)
        self.conn.commit()
        return self.conn

    # -------------------------------------------------------------------------
    # Consultas de Filtro Incremental (Camadas 1 e 2)
    # -------------------------------------------------------------------------
    def obter_corte_camada1(self) -> float:
        """Retorna o timestamp (float) da última execução concluída com sucesso.

        Se não houver execuções registradas, retorna 0.0 (sem filtro de corte).
        """
        assert self.conn is not None, "Banco não inicializado."
        cursor = self.conn.cursor()
        cursor.execute("SELECT MAX(finalizada_em) FROM execucoes WHERE finalizada_em IS NOT NULL")
        row = cursor.fetchone()
        if row and row[0]:
            try:
                # ISO datetime -> timestamp
                dt = datetime.datetime.fromisoformat(row[0])
                return dt.timestamp()
            except ValueError:
                pass
        return 0.0

    def verificar_camada2(
        self,
        caminho_str: str,
        tamanho: int,
        mtime: float,
        basename: Optional[str] = None,
    ) -> Tuple[str, Optional[sqlite3.Row]]:
        """Avalia um arquivo contra a tabela arquivos_processados.

        Busca por caminho absoluto OU por basename canônico para ser resiliente a
        trocas de computador (Casa vs Trabalho) e arquivos com sufixos de conflito do OneDrive.

        Retornos possíveis de status:
          - "NOVO": arquivo ainda não cadastrado.
          - "IGNORAR": arquivo cadastrado com tamanho e mtime idênticos.
          - "ATUALIZAR": arquivo cadastrado, porém com tamanho ou mtime alterado.
        """
        assert self.conn is not None, "Banco não inicializado."

        if not basename:
            basename = Path(caminho_str).name

        try:
            from parsers.filename_parser import limpar_sufixos_conflito_onedrive
        except ImportError:
            try:
                from dashboard_atos.parsers.filename_parser import limpar_sufixos_conflito_onedrive
            except ImportError:
                def limpar_sufixos_conflito_onedrive(x: str) -> str:
                    return x

        ext = Path(basename).suffix
        stem = Path(basename).stem
        basename_limpo = limpar_sufixos_conflito_onedrive(stem) + ext

        cursor = self.conn.cursor()
        cursor.execute(
            """
            SELECT caminho, basename, tamanho, mtime, processado_em, npu, tipo_ato
            FROM arquivos_processados
            WHERE caminho = ? OR basename = ? OR basename = ?
            LIMIT 1
            """,
            (caminho_str, basename, basename_limpo),
        )
        registro = cursor.fetchone()
        if registro is None:
            return "NOVO", None

        # Tolerância de 2.0 segundos para diferenças de sincronização do OneDrive / rede
        mesmo_tamanho = int(registro["tamanho"]) == int(tamanho)
        mesmo_mtime = abs(float(registro["mtime"]) - float(mtime)) < 2.0

        if mesmo_tamanho and mesmo_mtime:
            return "IGNORAR", registro
        return "ATUALIZAR", registro

    def registrar_arquivo(
        self,
        caminho_str: str,
        basename: str,
        tamanho: int,
        mtime: float,
        npu: str,
        tipo_ato: str,
    ) -> None:
        """Insere ou atualiza o arquivo processado na tabela de controle."""
        assert self.conn is not None, "Banco não inicializado."
        agora_iso = datetime.datetime.now().isoformat()
        cursor = self.conn.cursor()
        cursor.execute(
            """
            INSERT INTO arquivos_processados (caminho, basename, tamanho, mtime, processado_em, npu, tipo_ato)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(caminho) DO UPDATE SET
                basename = excluded.basename,
                tamanho = excluded.tamanho,
                mtime = excluded.mtime,
                processado_em = excluded.processado_em,
                npu = excluded.npu,
                tipo_ato = excluded.tipo_ato
            """,
            (caminho_str, basename, tamanho, mtime, agora_iso, npu, tipo_ato),
        )

    def obter_remessas_conhecidas(self) -> dict[str, str]:
        """Retorna dicionário {remessa: data_chegada}."""
        assert self.conn is not None, "Banco não inicializado."
        cursor = self.conn.cursor()
        cursor.execute("SELECT remessa, data_chegada FROM remessas")
        return {row["remessa"]: row["data_chegada"] or "" for row in cursor.fetchall()}

    def registrar_remessas(self, remessas_detectadas: list[str]) -> None:
        """Garante que remessas detectadas estejam no banco sem sobrescrever data existente."""
        assert self.conn is not None, "Banco não inicializado."
        cursor = self.conn.cursor()
        for remessa in remessas_detectadas:
            if not remessa or remessa == "PENDENTE":
                continue
            cursor.execute(
                """
                INSERT INTO remessas (remessa, data_chegada)
                VALUES (?, '')
                ON CONFLICT(remessa) DO NOTHING
                """,
                (remessa,),
            )

    def registrar_execucao(
        self,
        iniciada_em: str,
        finalizada_em: str,
        novos: int,
        atualizados: int,
        ignorados: int,
        erros: int,
    ) -> None:
        """Registra métricas da execução atual."""
        assert self.conn is not None, "Banco não inicializado."
        cursor = self.conn.cursor()
        cursor.execute(
            """
            INSERT INTO execucoes (iniciada_em, finalizada_em, novos, atualizados, ignorados, erros, maquina)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                iniciada_em,
                finalizada_em,
                novos,
                atualizados,
                ignorados,
                erros,
                platform.node(),
            ),
        )
        self.conn.commit()

    # -------------------------------------------------------------------------
    # Persistência Atômica e Backup
    # -------------------------------------------------------------------------
    def salvar_e_persistir(self) -> Path:
        """Serializa o banco SQLite em memória para o arquivo .sql remoto com troca atômica e backup.

        Passos:
          1. Copiar o .sql atual para backups/estado_AAAAMMDD_HHMMSS.sql.
          2. Serializar com iterdump() para .sql.tmp na mesma pasta remota.
          3. os.replace(tmp, estado.sql) com retry (3 tentativas, backoff 2s).
        """
        assert self.conn is not None, "Banco não inicializado."
        self.conn.commit()

        # 1. Backup datado do arquivo atual se ele existir
        if self.remoto_sql.exists():
            self.backup_dir.mkdir(parents=True, exist_ok=True)
            timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            arquivo_backup = self.backup_dir / f"estado_{timestamp_str}.sql"
            try:
                arquivo_backup.write_bytes(self.remoto_sql.read_bytes())
                logger.info("Backup datado gerado com sucesso: %s", arquivo_backup)
            except OSError as exc:
                logger.error("Falha ao criar backup em '%s': %s", arquivo_backup, exc)
                raise RuntimeError(f"Abortando: não foi possível criar backup datado antes de persistir: {exc}") from exc

        # 2. Serializar banco em memória para .sql.tmp
        tmp_file = self.remoto_sql.parent / f"{self.remoto_sql.name}.tmp"
        try:
            dump_lines = list(self.conn.iterdump())
            conteudo_dump = "\n".join(dump_lines) + "\n"
            tmp_file.write_text(conteudo_dump, encoding="utf-8")
        except OSError as exc:
            raise RuntimeError(f"Falha ao gerar dump temporário em '{tmp_file}': {exc}") from exc

        # 3. Troca atômica resiliente (os.replace) com retry para OneDrive/rede
        sucesso = False
        ultimo_erro = None
        for tentativa in range(1, self.retry_replace + 1):
            try:
                os.replace(tmp_file, self.remoto_sql)
                sucesso = True
                logger.info(
                    "Estado remoto persistido atomicamente em: %s (tentativa %d/%d)",
                    self.remoto_sql,
                    tentativa,
                    self.retry_replace,
                )
                break
            except OSError as exc:
                ultimo_erro = exc
                logger.warning(
                    "Tentativa %d/%d de troca atômica falhou para '%s': %s. Aguardando %.1fs...",
                    tentativa,
                    self.retry_replace,
                    self.remoto_sql,
                    exc,
                    self.retry_sleep_seg,
                )
                time.sleep(self.retry_sleep_seg)

        if not sucesso:
            raise RuntimeError(
                f"Falha crítica ao persistir estado em '{self.remoto_sql}' após {self.retry_replace} tentativas. "
                f"O arquivo temporário foi preservado em '{tmp_file}' para inspeção manual. Causa: {ultimo_erro}"
            )

        return self.remoto_sql

    def fechar(self) -> None:
        """Fecha a conexão em memória."""
        if self.conn:
            try:
                self.conn.close()
            except Exception:
                pass
            self.conn = None
