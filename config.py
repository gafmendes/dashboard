"""Configurações centrais do Dashboard de Atos Judiciais - TJPE (8.17).

Identifica se o script está rodando no computador de CASA ou do TRABALHO
através de verificação if/else direta dos caminhos específicos de cada ambiente.
"""

from __future__ import annotations

from pathlib import Path

# =============================================================================
# 1. DEFINIÇÃO EXPLÍCITA DOS CAMINHOS DE TRABALHO E CASA
# =============================================================================
CAMINHO_TRABALHO = Path(r"C:\Users\gafm\Downloads")
ONEDRIVE_TRABALHO = Path(r"C:\Users\gafm\OneDrive")

CAMINHO_CASA = Path(r"C:\Users\Gabriel Mendes\Downloads")
ONEDRIVE_CASA = Path(r"C:\Users\Gabriel Mendes\OneDrive")


# =============================================================================
# 2. VERIFICAÇÃO IF / ELSE: QUAL COMPUTADOR ESTÁ SENDO UTILIZADO
# =============================================================================
# Verifica se está no computador de CASA
if CAMINHO_CASA.exists() or Path(r"C:\Users\Gabriel Mendes").exists():
    PERFIL_SELECIONADO = "CASA"
    PASTA_DOWNLOADS = CAMINHO_CASA
    PASTA_ONEDRIVE = ONEDRIVE_CASA

# Se não, verifica se está no computador do TRABALHO
elif CAMINHO_TRABALHO.exists() or Path(r"C:\Users\gafm").exists():
    PERFIL_SELECIONADO = "TRABALHO"
    PASTA_DOWNLOADS = CAMINHO_TRABALHO
    PASTA_ONEDRIVE = ONEDRIVE_TRABALHO

# Fallback automático caso execute a partir de outro terminal/ambiente
else:
    user_home = Path.home()
    if "gafm" in user_home.name.lower():
        PERFIL_SELECIONADO = "TRABALHO"
        PASTA_DOWNLOADS = CAMINHO_TRABALHO
        PASTA_ONEDRIVE = ONEDRIVE_TRABALHO
    elif "gabriel" in user_home.name.lower() or "mendes" in user_home.name.lower():
        PERFIL_SELECIONADO = "CASA"
        PASTA_DOWNLOADS = CAMINHO_CASA
        PASTA_ONEDRIVE = ONEDRIVE_CASA
    else:
        PERFIL_SELECIONADO = "CASA"
        PASTA_DOWNLOADS = user_home / "Downloads"
        PASTA_ONEDRIVE = user_home / "OneDrive"

_cfg = {
    "descricao": f"Computador de {PERFIL_SELECIONADO.capitalize()}",
    "pasta_downloads": PASTA_DOWNLOADS,
    "pasta_onedrive": PASTA_ONEDRIVE,
}


# =============================================================================
# 3. DESTINOS ONEDRIVE (Nuvem) E DOWNLOADS (Local)
# =============================================================================
PASTA_ONEDRIVE_DASHBOARD = PASTA_ONEDRIVE / "dashboard"
try:
    PASTA_ONEDRIVE_DASHBOARD.mkdir(parents=True, exist_ok=True)
except Exception:
    pass

# Salva diretamente no OneDrive para sincronização automática na nuvem
SALVAR_NO_ONEDRIVE = True
COPIAR_PARA_DOWNLOADS = True

# Banco SQLite em memória serializado em .sql no OneDrive
REMOTO_SQL = PASTA_ONEDRIVE_DASHBOARD / "estado.sql"
BACKUP_DIR = PASTA_ONEDRIVE_DASHBOARD / "backups"

# Planilha Excel e Painel Web salvos no OneDrive
SAIDA_XLSX = PASTA_ONEDRIVE_DASHBOARD / "dashboard_atos.xlsx"
SAIDA_HTML = PASTA_ONEDRIVE_DASHBOARD / "dashboard_atos.html"


# =============================================================================
# 4. PARÂMETROS DE CONCORRÊNCIA, ESCRITA ATÔMICA E LOGGING
# =============================================================================
LOCK_TTL_MINUTOS = 30
RETRY_REPLACE = 3
RETRY_SLEEP_SEG = 2

ARQUIVO_LOG = PASTA_DOWNLOADS / "dashboard_atos.log"
LOG_MAX_BYTES = 5 * 1024 * 1024  # 5 MB
LOG_BACKUP_COUNT = 3
