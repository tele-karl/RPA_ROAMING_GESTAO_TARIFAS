import os
import re
from pathlib import Path
import pandas as pd

from datetime import datetime
import logging

from src.configs.configuration import (
    ORIGEM_DIR, DESTINO_DIR, SAIDA_DIR, CADASTRO_KEYWORD, MODELO_XML
)
from src.services.log_service import registrar_log
from src.utils.format_utils import (
    converter_unidade, formatar_valor, formatar_data_ddmmyyyy
)
from src.utils.xml_utils import (
    carregar_xml_para_dfs, atualizar_xml_multiplas_abas
)

# CRIAÇÃO DE LOGS PARA MONITORAR PREENCHIMENTO DO LOADER
# Pasta de logs
LOGS_DIR = "logs_loader"

# Timestamp da execução
timestamp_execucao = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

# Nome do arquivo de log (único por execução)
log_file_loader = os.path.join(
    LOGS_DIR,
    f"loader_{timestamp_execucao}.log"
)

# print(f" valor log file: {log_file_loader}")
# exit()

# Configuração do logging
logging.basicConfig(
    filename=log_file_loader,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    encoding="utf-8"
)

# Leitura do TAP consolidado
def carregar_dados_excel(caminho_arquivo: str) -> dict:
    """
    Lê o TAP consolidado (.xlsx) e retorna dict:
        {"Base Disc Rates": df, "Int. Call Disc.": df}
    (detecta cabeçalho pela 1ª coluna conter 'PLMN')
    """
    try:
        registrar_log("Iniciado", "Sucesso", f"Carregando dados do Excel consolidado: {caminho_arquivo}")
        dados = {}
        xls = pd.ExcelFile(caminho_arquivo, engine="openpyxl")
        for aba in xls.sheet_names:
            if "Base Disc Rates" in aba or "Int. Call Disc." in aba:
                df_temp = pd.read_excel(caminho_arquivo, sheet_name=aba, header=None)
                header_idx = df_temp.index[
                    df_temp.iloc[:, 0].astype(str).str.contains("PLMN", case=False, na=False)
                ][0]
                df_b = pd.read_excel(caminho_arquivo, sheet_name=aba, skiprows=header_idx+1)
                df_i = pd.read_excel(caminho_arquivo, sheet_name=aba, skiprows=header_idx)
                if "Base Disc Rates" in aba:
                    dados["Base Disc Rates"] = df_b
                else:
                    dados["Int. Call Disc."] = df_i
        registrar_log("Iniciado", "Sucesso", f"Dados do Excel consolidado carregados com sucesso.")
        return dados
    except Exception as e:
        registrar_log("Iniciado", "Sucesso", f"Erro ao carregar Excel consolidado: {e}")

# Helpers para arquivo de rejeitados (TXT)
def _encontrar_txt_rejeitados(saida_dir: Path) -> Path | None:
    """
    Procura o arquivo .txt que contenha 'plmns_rejeitados' no nome.
    Se houver mais de um, pega o mais recente (por data de modificação).
    """
    candidatos = list(saida_dir.glob(f"*{CADASTRO_KEYWORD}*.txt"))
    if not candidatos:
        raise FileNotFoundError(
            f"Nenhum arquivo '*{CADASTRO_KEYWORD}*.txt' encontrado em: {saida_dir.resolve()}"
        )
    candidatos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidatos[0]

def _ler_txt(path: Path) -> str:
    """
    Lê o conteúdo do TXT com fallback de encoding.
    """
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="latin-1", errors="replace")

# Injeção de PLMNs ausentes
def _filtrar_linhas_por_tarifa(df_xml_subset, dados_origem, plmn_doador):
    """
    Retorna apenas as linhas do subset (já filtrado pelo doador) que
    devem ser copiadas para a operadora ausente, de acordo com as regras:
      - Se houver dados na col. F (Base Disc Rates) -> IAIRTIME + Tier 000NNAT
      - Se houver dados na col. M (Base Disc Rates) -> IAIRTIME + Tier 000NMTC
      - Se houver dados na col. S (Base Disc Rates) -> IAIRTIME + Tier 000IHOME
      - Se houver dados na col. G (Int. Call Disc.) -> IAIRTIME + Tier NOT IN {000NNAT,000NMTC,000IHOME}
      - Se houver dados na col. Y (Base Disc Rates) -> ISMS_EVT (qualquer Tier)
      - Se houver dados na col. AA (Base Disc Rates) -> IGPRSVOL (qualquer Tier)
      - ITOLLDUR -> SEMPRE copiar
    """
    try:
        df_base = dados_origem.get("Base Disc Rates")
        df_int = dados_origem.get("Int. Call Disc.")

        # Normaliza PLMN doador para comparação
        plmn = str(plmn_doador).strip().upper()

        # Subconjuntos do TAP por PLMN (se existirem)
        if df_base is not None and not df_base.empty:
            df_base_plmn = df_base[df_base.iloc[:, 0].astype(str).str.strip().str.upper() == plmn]
        else:
            df_base_plmn = None
        if df_int is not None and not df_int.empty:
            df_int_plmn = df_int[df_int.iloc[:, 0].astype(str).str.strip().str.upper() == plmn]
        else:
            df_int_plmn = None

        # Sinais de presença de valores no TAP
        tem_nnat = bool(df_base_plmn is not None and not df_base_plmn.empty and df_base_plmn.iloc[:, 5].notna().any())   # F idx=5
        tem_nmtc = bool(df_base_plmn is not None and not df_base_plmn.empty and df_base_plmn.iloc[:, 12].notna().any())  # M idx=12
        tem_ihome = bool(df_base_plmn is not None and not df_base_plmn.empty and df_base_plmn.iloc[:, 18].notna().any()) # S idx=18
        tem_sms  = bool(df_base_plmn is not None and not df_base_plmn.empty and df_base_plmn.iloc[:, 24].notna().any())  # Y idx=24
        tem_gprs = bool(df_base_plmn is not None and not df_base_plmn.empty and df_base_plmn.iloc[:, 26].notna().any())  # AA idx=26
        tem_outros = bool(df_int_plmn is not None and not df_int_plmn.empty and df_int_plmn.iloc[:, 6].notna().any())    # G idx=6

        # Se subset vier vazio, retorna vazio
        if df_xml_subset is None or df_xml_subset.empty:
            return df_xml_subset

        # Garantir que as colunas de referência existam (Rating Component idx=5, Tier idx=8)
        ncols = df_xml_subset.shape[1]
        if ncols <= 8:
            # se faltar colunas, expande com colunas vazias
            df_xml_subset = df_xml_subset.reindex(columns=list(range(9)))

        # Séries normalizadas para comparação
        rc = df_xml_subset.iloc[:, 5].astype(str).str.strip().str.upper()
        tier = df_xml_subset.iloc[:, 8].astype(str).str.strip().str.upper()

        # Lista de filtros
        filtros = []
        if tem_nnat:
            filtros.append((rc == 'IAIRTIME') & (tier == '000NNAT'))
        if tem_nmtc:
            filtros.append((rc == 'IAIRTIME') & (tier == '000NMTC'))
        if tem_ihome:
            filtros.append((rc == 'IAIRTIME') & (tier == '000IHOME'))
        if tem_outros:
            filtros.append((rc == 'IAIRTIME') & (~tier.isin(['000NNAT','000NMTC','000IHOME'])))
        if tem_sms:
            filtros.append(rc == 'ISMS_EVT')
        if tem_gprs:
            filtros.append(rc == 'IGPRSVOL')
        # Sempre ITOLLDUR
        filtros.append(rc == 'ITOLLDUR')

        if not filtros:
            return df_xml_subset.iloc[0:0].copy()

        filtro_final = filtros[0]
        for f in filtros[1:]:
            filtro_final = filtro_final | f

        result = df_xml_subset[filtro_final].copy()

        # Log detalhado (contagens por tipo)
        try:
            counts = {
                'IAIRTIME_000NNAT': int(((rc=='IAIRTIME') & (tier=='000NNAT')).sum()) if tem_nnat else 0,
                'IAIRTIME_000NMTC': int(((rc=='IAIRTIME') & (tier=='000NMTC')).sum()) if tem_nmtc else 0,
                'IAIRTIME_000IHOME': int(((rc=='IAIRTIME') & (tier=='000IHOME')).sum()) if tem_ihome else 0,
                'IAIRTIME_OUTROS': int(((rc=='IAIRTIME') & (~tier.isin(['000NNAT','000NMTC','000IHOME']))).sum()) if tem_outros else 0,
                'ISMS_EVT': int((rc=='ISMS_EVT').sum()) if tem_sms else 0,
                'IGPRSVOL': int((rc=='IGPRSVOL').sum()) if tem_gprs else 0,
                'ITOLLDUR': int((rc=='ITOLLDUR').sum()),
            }
            logging.info(f"Filtro de tarifas aplicado ao doador {plmn}: {counts} -> mantendo {len(result)} linha(s)")
        except Exception:
            pass

        return result
    except Exception as e:
        try:
            logging.error(f"Erro em filtrar_linhas_por_tarifa: {e}")
        except Exception:
            pass
        return df_xml_subset

def _obter_linhas_do_modelo(plmn_faltante, dados_origem, caminho_modelo_xml):
    """
    Obtém linhas de serviços do RATE_MODELO.xml para um PLMN faltante.
    Copia exatamente como no modelo, substituindo apenas o PLMN (coluna índice 4).
    """

    try:     
        planilhas_modelo, _, _, _ = carregar_xml_para_dfs(caminho_modelo_xml)

        COL_PLMN = 4
        novas_linhas = []

        for nome_aba, info in planilhas_modelo.items():
            df_modelo = info.get("df")
            if df_modelo is None or df_modelo.empty:
                continue

            max_cols = max(df_modelo.shape[1], 5)
            if df_modelo.shape[1] < max_cols:
                df_modelo = df_modelo.reindex(columns=list(range(max_cols)))

            linhas = _filtrar_linhas_por_tarifa(df_modelo, dados_origem, plmn_faltante)

            if linhas.empty:
                continue

            linhas = linhas.copy(deep=True)
            linhas.iloc[:, COL_PLMN] = plmn_faltante  # substituição obrigatória

            novas_linhas.append(linhas)

            logging.info(
                f"[FALLBACK-MODELO] {len(linhas)} linha(s) copiadas do modelo para '{plmn_faltante}' na aba '{nome_aba}'.")
            
            registrar_log("Finalizado", "Sucesso", f"[FALLBACK-MODELO] {len(linhas)} linha(s) copiadas do modelo para '{plmn_faltante}' na aba '{nome_aba}'.")

        return novas_linhas

    except Exception as e:
        logging.error(f"[ERRO MODELO] Falha ao obter linhas do modelo: {e}")
        return []
    
def _copiar_do_doador(
    grupo,
    plmn_destino,
    dados_origem,
    mapa_grupo_xml,
    planilhas,
    faltantes,
    col_plmn=4,
    col_rc=5,
    col_tier=8
):
    """
    Seleciona o doador de forma inteligente:
      1) Filtra PLMNs do mesmo grupo.
      2) Seleciona o que começa com o mesmo prefixo (2 letras).
      3) Se não houver match -> fallback MODELO (retorna None).
      4) Copia apenas os serviços faltantes (ou todos se faltantes=None).
    """

    if not isinstance(plmn_destino, str):
        logging.error(f"PLMN destino inválido: {plmn_destino!r}")
        return []

    plmn_norm = plmn_destino.strip().upper()
    prefixo = plmn_norm[:3]

    doadores = sorted(list(mapa_grupo_xml.get(grupo, [])))
    if not doadores:
        return None  # sem doadores no grupo

    # Seleção inteligente por prefixo
    doadores_match = [
        d for d in doadores 
        if d.upper().startswith(prefixo)
        and d.upper() != plmn_destino.upper()
    ]

    if doadores_match:
        doador = doadores_match[0]   # Ordem alfabética
        registrar_log("Iniciado", "Sucesso", f"[DOADOR] {doador} (prefixo '{prefixo}') doa para {plmn_destino}")
        logging.warning(f"[DOADOR] {doador} (prefixo '{prefixo}') doa para {plmn_destino}")
    else:
        registrar_log("Finalizado", "Sucesso", f"[DOADOR] Nenhum match de prefixo '{prefixo}': usar MODELO")
        logging.warning(f"[DOADOR] Nenhum match prefixo '{prefixo}': usar MODELO")
        return None

    novos = []

    for info_aba in planilhas.values():
        df_aba = info_aba.get("df")
        if df_aba is None or df_aba.empty or df_aba.shape[1] <= col_plmn:
            continue

        # Linhas do doador
        mask = df_aba.iloc[:, col_plmn].astype(str).str.upper() == doador
        linhas_doador = df_aba[mask]

        filtradas = _filtrar_linhas_por_tarifa(linhas_doador, dados_origem, doador)
        if filtradas.empty:
            continue

        filtradas = filtradas.copy()
        filtradas.iloc[:, col_plmn] = plmn_destino

        # Caso "faltantes=None" -> copiar tudo que o filtro TAP determinou
        if faltantes is None:
            novos.append(filtradas)
            continue

        # Gerar chaves RC+TIER
        filtradas_chaves = list(zip(
            filtradas.iloc[:, col_rc].astype(str).str.upper(),
            filtradas.iloc[:, col_tier].astype(str).str.upper()
        ))

        # Filtrar apenas (RC, TIER) faltantes
        filtradas = filtradas[[key in faltantes for key in filtradas_chaves]]

        if not filtradas.empty:
            novos.append(filtradas)

    return novos if novos else None

# Injetar PLMNs ausentes no XML 
def _injetar_plmns_ausentes(planilhas: dict, dados_origem: dict, saida_dir: str) -> None:
    """
    Injeta linhas APENAS para PLMNs que:
      - estão no XLSX (col A),
      - NÃO estão no TXT de rejeitados,
      - e NÃO estão no XML (col E, idx 4).
    O doador é um PLMN do MESMO GRUPO (col B do XLSX) já presente no XML,
    considerando uma verificação do prefixo.
    As linhas clonadas (de todas as abas) são anexadas ao final da ÚLTIMA aba.
    - fallback com arquivo modelo quando não há doador no XML
    - substituição correta do PLMN na coluna índice 4
    - serviços copiados exatamente conforme regras do XLSX
    """
    try:
        logging.info("Iniciado verificação de PLMNs faltantes no XML de origem.")
        registrar_log("Iniciado", "Sucesso", f"Iniciado verificação de PLMNs faltantes no XML de origem.")

        # Caminho para o arquivo modelo RATE_MODELO.xml
        caminho_modelo_xml = MODELO_XML

        # Prepara conjunto de PLMNs do XLSX e mapa PLMN->grupo (coluna B / idx=1)
        df_base = dados_origem.get("Base Disc Rates")
        df_int = dados_origem.get("Int. Call Disc.")

        # Helpers
        def plmn_col(df):
            if df is None:
                return pd.Series([], dtype=str)
            col = df.iloc[:, 0].astype(str).str.strip().str.upper()
            return col[~col.isin(["", "NAN", "NONE", "NULL", "0", "N/A"])]

        def grupo_col(df):
            if df is None:
                return pd.Series([], dtype=str)
            col = df.iloc[:, 1].astype(str).str.strip().str.upper()
            return col[~col.isin(["", "NAN", "NONE", "NULL", "0", "N/A"])]

        # PLMNs do XLSX
        plmns_xlsx = set(plmn_col(df_base)) | set(plmn_col(df_int))

        # Mapa PLMN -> Grupo
        mapa_grupo = {}
        for df in (df_base, df_int):
            if df is None:
                continue
            for p, g in zip(plmn_col(df), grupo_col(df)):
                if p and p not in mapa_grupo:
                    mapa_grupo[p] = g

        # Ler TXT de rejeitados (se não houver, consideramos conjunto vazio)
        try:
            rejeitados_txt = _encontrar_txt_rejeitados(Path(saida_dir))
            conteudo = _ler_txt(rejeitados_txt)
            logging.info(f"Arquivo de rejeitados utilizado: {rejeitados_txt}")
        except FileNotFoundError:
            logging.warning("Nenhum arquivo de rejeitados encontrado.")
            conteudo = ""

        # Extrair tokens candidatos a PLMN (regex) e obter conjunto de rejeitados válidos (presentes no XLSX)
        rejeitados = {
            tok
            for tok in re.findall(r"[A-Za-z0-9]{3,10}", conteudo.upper())
            if tok in plmns_xlsx
        }

        # PLMNs presentes no XML
        COL_PLMN = 4
        plmns_xml = set()
        max_cols_xml = 0

        for _, info in planilhas.items():
            df = info.get("df")
            if df is None or df.empty:
                continue

            max_cols_xml = max(max_cols_xml, df.shape[1])
            
            if df.shape[1] > COL_PLMN:
                plmns_xml |= set(
                    df.iloc[:, COL_PLMN].astype(str).str.strip().str.upper()
                )

        # REGRA: PLMNs a injetar = (PLMNs XLSX - rejeitados) - PLMNs já no XML
        candidatos_validos = plmns_xlsx - rejeitados
        plmns_faltantes = [p for p in candidatos_validos if p not in plmns_xml]

        if not plmns_faltantes:
            registrar_log("Finalizado", "Sucesso", f"[PLMN-FALTANTE] Nenhum PLMN faltante encontrado.")
            logging.info("[PLMN-FALTANTE] Nenhum PLMN faltante encontrado.")
            return

        # Determinar ÚLTIMA ABA do XML (para anexar as linhas lá)
        ultima_aba = list(planilhas.keys())[-1]
        df_dest = planilhas[ultima_aba].get("df")

        if df_dest is None or df_dest.empty:
            df_dest = pd.DataFrame(columns=list(range(max_cols_xml)))
        elif df_dest.shape[1] < max_cols_xml:
            df_dest = df_dest.reindex(columns=list(range(max_cols_xml)))

        #  Mapa grupo -> PLMNs que JÁ estão no XML (para achar doador por grupo)
        mapa_grupo_xml = {}
        for p in plmns_xml:
            g = mapa_grupo.get(p)
            if g:
                mapa_grupo_xml.setdefault(g, set()).add(p)

        novas_linhas = []

        # PROCESSAMENTO
        for plmn in plmns_faltantes:

            logging.info(f"[PLMN-FALTANTE] Detectado PLMN ausente no XML: {plmn}")
            registrar_log("Iniciado", "Sucesso", f"[PLMN-FALTANTE] Detectado PLMN ausente no XML: {plmn}")

            grupo = mapa_grupo.get(plmn, "")

            # Seleção inteligente do doador
            novos = _copiar_do_doador(
                grupo=grupo,
                plmn_destino=plmn,
                dados_origem=dados_origem,
                mapa_grupo_xml=mapa_grupo_xml,
                planilhas=planilhas,
                faltantes=None,   # copiar TUDO que filtrar_linhas_por_tarifa permitir
            )

            if novos is None:
                registrar_log("Iniciado", "Sucesso", f"[SEM-DOADOR] PLMN '{plmn}' inserido usando MODELO.")
                logging.warning(f"[SEM-DOADOR] PLMN '{plmn}' inserido usando MODELO.")
                linhas_modelo = _obter_linhas_do_modelo(plmn, dados_origem, caminho_modelo_xml)
                novas_linhas.extend(linhas_modelo)
                continue

            # Agregar linhas das abas
            novas_linhas.extend(novos)

            logging.info(f"[PLMN-INSERIDO] PLMN '{plmn}' inserido com sucesso no XML. {len(novos)} linhas adicionada(s).")
            registrar_log("Finalizado", "Sucesso", f"[PLMN-INSERIDO] PLMN '{plmn}' inserido com sucesso no XML. {len(novos)} linhas adicionada(s).")

        # Inserir na última aba
        if novas_linhas:
            df_dest = pd.concat([df_dest] + novas_linhas, ignore_index=True)
            planilhas[ultima_aba]["df"] = df_dest

    except Exception as e:
        logging.error(f"[ERRO] na função injetar_plmns_ausentes: {e}")
        raise

# Limpeza de expiradas
def _limpar_tarifas_expiradas(df_destino: pd.DataFrame) -> pd.DataFrame:
    """
    Remove linhas com tarifas expiradas. Coluna AD (index 29) = End Date.
    Mantém linhas sem data (ativas).
    """
    try:
        registrar_log("Iniciado", "Sucesso", "Remover tarifas expiradas.")
        col_end_date = 29  # Coluna AD
        if df_destino.empty:
            return df_destino
        resultado = df_destino[
            df_destino[col_end_date].isna() | (df_destino[col_end_date].astype(str).str.strip() == "")
        ]
        logging.info(f"Linhas restantes após limpeza: {len(resultado)}")
        return resultado
    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Erro ao limpar tarifas expiradas: {e}")

# Completar serviços para PLMNs já existentes no XML
def _completar_servicos(planilhas, dados_origem):
    """
    Garante que TODOS os serviços necessários para cada PLMN presente no XML
    estejam presentes, conforme regras de filtragem.

    Para cada PLMN já existente no XML:
        - Gera as linhas esperadas via RATE_MODELO.xml
        - Filtra essas linhas usando filtrar_linhas_por_tarifa()
        - Compara com linhas existentes no XML
        - O que estiver faltando é copiado de um PLMN doador do mesmo grupo e se não houver, 
          de RATE_MODELO.xml
    """

    logging.info("Iniciando verificação de serviços faltantes para PLMNs existentes.")
    registrar_log("Iniciado", "Sucesso", f"Verificação de serviços faltantes para PLMNs existentes.")
    

    # Carrega modelo
    caminho_modelo_xml = MODELO_XML
    planilhas_modelo, _, _, _ = carregar_xml_para_dfs(caminho_modelo_xml)

    df_base = dados_origem.get("Base Disc Rates")
    df_int = dados_origem.get("Int. Call Disc.")

    # Helpers
    def plmn_col(df):
        if df is None:
            return pd.Series([], dtype=str)
        col = df.iloc[:, 0].astype(str).str.strip().str.upper()
        return col[~col.isin(["", "NAN", "NONE", "NULL", "0", "N/A"])]

    def grupo_col(df):
        if df is None:
            return pd.Series([], dtype=str)
        col = df.iloc[:, 1].astype(str).str.strip().str.upper()
        return col[~col.isin(["", "NAN", "NONE", "NULL", "0", "N/A"])]

    # Mapeamento PLMN -> Grupo
    mapa_grupo = {}
    for df in (df_base, df_int):
        if df is None:
            continue
        for p, g in zip(plmn_col(df), grupo_col(df)):
            if p and g:
                mapa_grupo[p] = g

    # PLMNs presentes no XML
    COL_PLMN = 4
    plmns_xml = set()

    for info in planilhas.values():
        df = info.get("df")
        if df is None or df.empty:
            continue
        if df.shape[1] > COL_PLMN:
            plmns_xml |= set(df.iloc[:, COL_PLMN].astype(str).str.strip().str.upper())

    # Mapeamento grupo -> doadores existentes no XML
    mapa_grupo_xml = {}
    for p in plmns_xml:
        g = mapa_grupo.get(p)
        if g:
            mapa_grupo_xml.setdefault(g, set()).add(p)

    novas_linhas_gerais = []

    # PROCESSAR CADA PLMN JÁ EXISTENTE
    for plmn in plmns_xml:
        grupo = mapa_grupo.get(plmn, "")

        # 1) Gerar linhas esperadas do MODELO
        linhas_esperadas = []
        for _, info_mod in planilhas_modelo.items():
            df_mod = info_mod.get("df")
            if df_mod is None or df_mod.empty:
                continue

            max_cols = max(df_mod.shape[1], 5)
            df_mod = df_mod.reindex(columns=list(range(max_cols)))

            filtradas = _filtrar_linhas_por_tarifa(df_mod, dados_origem, plmn)
            if not filtradas.empty:
                filtradas = filtradas.copy()
                filtradas.iloc[:, COL_PLMN] = plmn
                linhas_esperadas.append(filtradas)

        if not linhas_esperadas:
            logging.warning(f"[SERVIÇOS] Nenhuma linha esperada no modelo para {plmn}.")
            continue

        df_esperado = pd.concat(linhas_esperadas, ignore_index=True)

        # 2) Linhas existentes para este PLMN no XML
        linhas_existentes = []
        for info in planilhas.values():
            df_aba = info.get("df")
            if df_aba is None or df_aba.empty:
                continue
            if df_aba.shape[1] <= COL_PLMN:
                continue

            mask = df_aba.iloc[:, COL_PLMN].astype(str).str.upper() == plmn
            linhas_existentes.append(df_aba[mask])

        df_existentes = (
            pd.concat(linhas_existentes, ignore_index=True)
            if linhas_existentes else
            pd.DataFrame(columns=df_esperado.columns)
        )

        # Não considerar tarifas expiradas como existentes
        # df_existentes = _limpar_tarifas_expiradas(df_existentes)

        # 3) Identificar serviços faltantes (RC+TIER)
        col_rc = 5
        col_tier = 8

        chave_esperado = set(zip(
            df_esperado.iloc[:, col_rc].astype(str).str.upper(),
            df_esperado.iloc[:, col_tier].astype(str).str.upper()
        ))

        chave_existentes = set(zip(
            df_existentes.iloc[:, col_rc].astype(str).str.upper(),
            df_existentes.iloc[:, col_tier].astype(str).str.upper()
        ))

        faltantes = chave_esperado - chave_existentes

        if not faltantes:
            logging.info(f"[SERVIÇOS] {plmn}: nenhum serviço faltante.")
            continue

        registrar_log("Iniciado", "Sucesso", f"[SERVIÇOS] {plmn}: faltando {len(faltantes)} serviços: buscar.")
        logging.info(f"[SERVIÇOS] {plmn}: faltando {len(faltantes)} serviços: buscar.")

        # 4) Copiar do doador inteligente (prefixo)
        novos = _copiar_do_doador(
            grupo=grupo,
            plmn_destino=plmn,
            dados_origem=dados_origem,
            mapa_grupo_xml=mapa_grupo_xml,
            planilhas=planilhas,
            faltantes=faltantes,
            col_plmn=COL_PLMN,
            col_rc=col_rc,
            col_tier=col_tier
        )

        # LOG — serviços copiados via DOADOR
        if novos:
            for df_novo in novos:
                if df_novo is None or df_novo.empty:
                    continue
                for _, row in df_novo.iterrows():
                    rc = str(row.iloc[col_rc]).strip()
                    tier = str(row.iloc[col_tier]).strip()
                registrar_log("Iniciado", "Sucesso", f"[SERVIÇOS] Serviço {rc} (Tier {tier}) copiado de PLMN doador do mesmo grupo para o receptor {plmn}.")
                logging.info(
                    f"[SERVIÇOS] Serviço {rc} (Tier {tier}) copiado de PLMN doador do mesmo grupo para o receptor {plmn}.")

        # 5) FALLBACK MODELO
        if novos is None:
            logging.warning(f"[SERVIÇOS] Sem doador para {plmn}, usando modelo.")
            registrar_log("Iniciado", "Sucesso", f"[SERVIÇOS] Sem doador para {plmn}, usando modelo.")
            
            novos = []
            df_keys = list(zip(
                df_esperado.iloc[:, col_rc].str.upper(),
                df_esperado.iloc[:, col_tier].str.upper()
            ))

            for (rc, tier), row in zip(df_keys, df_esperado.values.tolist()):
                if (rc, tier) in faltantes:
                    novos.append(pd.DataFrame([row], columns=df_esperado.columns))
            registrar_log("Iniciado", "Sucesso", f"[SERVIÇOS] Serviço {rc} (Tier {tier}) copiado do MODELO para o receptor {plmn}.")
            logging.info(
                f"[SERVIÇOS] Serviço {rc} (Tier {tier}) copiado do MODELO para o receptor {plmn}.")
            
        # 6) Agregar
        for n in novos:
            if n is not None and not n.empty:
                novas_linhas_gerais.append(n)

    # INSERIR NA ÚLTIMA ABA
    if novas_linhas_gerais:
        ultima_aba = list(planilhas.keys())[-1]
        df_dest = planilhas[ultima_aba]["df"]
        df_dest = pd.concat([df_dest] + novas_linhas_gerais, ignore_index=True)
        planilhas[ultima_aba]["df"] = df_dest
        logging.info(f"[SERVIÇOS] Inseridas {sum(len(n) for n in novas_linhas_gerais)} linhas faltantes.")
    else:
        logging.info("[SERVIÇOS] Nenhum serviço faltante encontrado.")
        registrar_log("Finalizado", "Sucesso", f"[SERVIÇOS] Nenhum serviço faltante encontrado.")

# Aplicação de regras
def _aplicar_regras(df_destino: pd.DataFrame, dados_origem: dict, cabecalho=None) -> pd.DataFrame:
    """
    Aplica regras de negócio por linha (ISMS_EVT, IGPRSVOL, IAIRTIME, ITOLLDUR)
    na planilha (DataFrame) de uma aba.
    Ajusta datas, unidades e custos conforme TAP.
    """

    logs_emitidos = set()

    def log_unico(chave, mensagem):
        """Função que permite o log de atualização de tarifas especificando
        operadoras, serviços e valores uma única vez para cada serviço"""
        if chave not in logs_emitidos:
            registrar_log("Iniciado", "Sucesso", mensagem)
            logging.info(mensagem)
            logs_emitidos.add(chave)

    try:
        if df_destino is None or df_destino.empty:
            return df_destino
        registrar_log("Iniciado", "Sucesso", "Aplicando regras no loader.")
        df_base = dados_origem.get("Base Disc Rates")
        df_int  = dados_origem.get("Int. Call Disc.")

        # Índices das colunas no destino (ajustáveis via cabeçalho)
        col_network_operator  = 4
        col_rating_component  = 5
        col_comp_direction    = 7
        col_tier              = 8
        col_start_date        = 28
        col_rate_unit         = 30
        col_additional_charge = 33
        col_rounding_unit     = 43
        col_unit_cost         = 51

        # Excluir linhas ISMS_EVT com Comp Direction = "I" (exclusivo SMS)
        df_destino = df_destino[
            ~((df_destino[col_rating_component] == "ISMS_EVT") &
              (df_destino[col_comp_direction] == "I"))
        ]

        for i, row in df_destino.iterrows():
            plmn = str(row[col_network_operator]).strip().upper()
            rating_component = str(row[col_rating_component]).strip()
            tier = str(row[col_tier]).strip()

            # Valor atual de rounding para preservar se necessário
            old_round = str(row[col_rounding_unit]).strip()

            # match por PLMN em Base Disc Rates
            dados_plmn = df_base[df_base.iloc[:, 0].astype(str).str.strip().str.upper() == plmn]
            if dados_plmn.empty:
                registrar_log("Iniciado", "Erro", f"PLMN {plmn} não encontrado na planilha consolidada — linha {i} mantida sem atualização.")
                continue 

            # Buscar coluna Effective Date pelo nome ou fallback
            if "Effective Date" in dados_plmn.columns:
                effective_date = dados_plmn.iloc[0]["Effective Date"]
            else:
                effective_date = dados_plmn.iloc[0].iloc[3]
            if isinstance(effective_date, pd.Timestamp):
                effective_date = effective_date.date()
            try:
                if isinstance(effective_date, str) and effective_date.isdigit():
                    effective_date = int(effective_date)
            except Exception as e:
                logging.warning(f"Erro ao buscar coluna Effective Date: {e}")
                pass
            if pd.notna(effective_date):
                df_destino.at[i, col_start_date] = formatar_data_ddmmyyyy(effective_date)

            # Helper para Rounding_Unit
            def atualizar_rounding(incremento_val):
                token = converter_unidade(incremento_val)
                if token is not None and str(token).strip() != "":
                    df_destino.at[i, col_rounding_unit] = token
                else:
                    # preserva o valor anterior
                    df_destino.at[i, col_rounding_unit] = old_round

            # Regras por componente
            if rating_component == "ISMS_EVT":
                valor = dados_plmn.iloc[0].iloc[24]
                # Garantir unidades de Rounding Unit e Rounding Unit
                df_destino.at[i, col_rate_unit] = "MINUTE"
                df_destino.at[i, col_rounding_unit] = "SECO"
                if pd.notna(valor):
                    df_destino.at[i, col_additional_charge] = formatar_valor(valor)
                    chave_log = f"{plmn}|ISMS_EVT"
                    log_unico(
                        chave_log,
                        f"Tarifa ISMS_EVT atualizada para {plmn}: {formatar_valor(valor)}")

            elif rating_component == "ITOLLDUR":
                # Garantir unidades de Rate Unit e Rounding Unit
                df_destino.at[i, col_rate_unit] = "MINUTE"
                df_destino.at[i, col_rounding_unit] = "SECO"

            elif rating_component == "IGPRSVOL":
                valor   = dados_plmn.iloc[0].iloc[26]
                unidade = dados_plmn.iloc[0].iloc[27]
                # Garantir que Rounding Unit seja sempre KBYTE
                df_destino.at[i, col_rounding_unit] = "KBYTE"
                if pd.notna(valor):
                    df_destino.at[i, col_unit_cost] = formatar_valor(valor)
                    df_destino.at[i, col_rate_unit] = converter_unidade(unidade)
                    chave_log = f"{plmn}|IGPRSVOL"
                    log_unico(
                        chave_log,
                        f"Tarifa IGPRSVOL atualizada para {plmn}: {formatar_valor(valor)}")

            elif rating_component == "IAIRTIME":
                # Rate Unit sempre MINUTE
                df_destino.at[i, col_rate_unit] = "MINUTE"

                if tier == "000NMTC":
                    valor = dados_plmn.iloc[0].iloc[12]
                    incremento = dados_plmn.iloc[0].iloc[13]
                    # Atualiza mesmo se unit_cost = 0
                    atualizar_rounding(incremento)  
                    if pd.notna(valor) and row[col_unit_cost] != "0":
                        df_destino.at[i, col_unit_cost] = formatar_valor(valor)
                        chave_log = f"{plmn}|IAIRTIME|000NMTC"
                        log_unico(
                            chave_log,
                            f"Tarifa IAIRTIME 000NMTC atualizada para {plmn}: {formatar_valor(valor)}")

                elif tier == "000NNAT":
                    valor = dados_plmn.iloc[0].iloc[5]
                    incremento = dados_plmn.iloc[0].iloc[6]
                    # Atualiza mesmo se unit_cost = 0
                    atualizar_rounding(incremento)  
                    if pd.notna(valor) and row[col_unit_cost] != "0":
                        df_destino.at[i, col_unit_cost] = formatar_valor(valor)
                        chave_log = f"{plmn}|IAIRTIME|000NNAT"
                        log_unico(
                            chave_log,
                            f"Tarifa IAIRTIME 000NNAT atualizada para {plmn}: {formatar_valor(valor)}")

                elif tier == "000IHOME":
                    valor = dados_plmn.iloc[0].iloc[18]
                    incremento = dados_plmn.iloc[0].iloc[19]
                    # Atualiza mesmo se unit_cost = 0
                    atualizar_rounding(incremento)  
                    if pd.notna(valor) and row[col_unit_cost] != "0":
                        df_destino.at[i, col_unit_cost] = formatar_valor(valor)
                        chave_log = f"{plmn}|IAIRTIME|000IHOME"
                        log_unico(
                            chave_log,
                            f"Tarifa IAIRTIME 000IHOME atualizada para {plmn}: {formatar_valor(valor)}")  
                # Para todos os demais Tiers (consulta "Int. Call Disc.")
                else:
                    # Converter unidade mesmo se valor for 0
                    if df_int is not None:
                        dados_int = df_int[df_int.iloc[:, 0].astype(str).str.strip().str.upper() == plmn]
                        if not dados_int.empty:
                            valor = dados_int.iloc[0].iloc[6]
                            incremento = dados_int.iloc[0].iloc[7]
                            # Atualiza mesmo se unit_cost = 0
                            atualizar_rounding(incremento)
                            if pd.notna(valor) and row[col_unit_cost] != "0":
                                df_destino.at[i, col_unit_cost] = formatar_valor(valor)
                                chave_log = f"{plmn}|IAIRTIME|OUTROS_TIERS"
                                log_unico(
                                    chave_log,
                                    f"Tarifa IAIRTIME (demais tiers) atualizada para {plmn}: {formatar_valor(valor)}")
                        else:
                            # preserva rounding
                            df_destino.at[i, col_rounding_unit] = old_round
                    else:
                        df_destino.at[i, col_rounding_unit] = old_round
                        
        registrar_log("Finalizado", "Sucesso", "Regras aplicadas ao loader.")
        return df_destino
    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Erro ao aplicar regras ao loader: {e}")

def _montar_planilhas_finais(planilhas_dict: dict) -> dict:
    """
    Recebe:
      planilhas_dict: { sheet: {"df": df, "estrutura": list, "cabecalho": list} }
    Retorna:
      planilhas_finais: { sheet: DataFrame_final (estrutura + cabeçalho + dados) }
    """
    finais = {}
    for nome, info in planilhas_dict.items():
        df = info["df"]
        linha_estrutura = info["estrutura"]
        linha_cabecalho = info["cabecalho"]

        df_final = pd.DataFrame([linha_estrutura, linha_cabecalho])
        if df is not None and not df.empty:
            df_final = pd.concat([df_final, df], ignore_index=True)

        # Limpeza pontual: célula B1 vazia se existir
        if len(df_final.columns) > 1 and len(df_final) > 0:
            if pd.notna(df_final.iloc[0, 1]) and str(df_final.iloc[0, 1]).strip() != "":
                df_final.iloc[0, 1] = ""

        # Substituir vírgulas por pontos em strings
        for col in df_final.columns:
            df_final[col] = df_final[col].apply(lambda x: str(x).replace(',', '.') if isinstance(x, str) else x)
        finais[nome] = df_final
    return finais

def salvar_xlsx_multiplas_abas(caminho_xlsx, planilhas_finais):
    """
    Grava um XLSX com múltiplas abas, preservando os nomes.
    """
    with pd.ExcelWriter(caminho_xlsx, engine="openpyxl") as writer:
        for nome, df in planilhas_finais.items():
            # header=False para manter estilo do XML: duas primeiras linhas são "dados"
            df.to_excel(writer, sheet_name=nome[:31] or "Sheet1", index=False, header=False)

# API principal do service
def preencher_loader_multiplas_abas():
    """
    Fluxo:
      1) localizar TAP .xlsx e RATE_*.xml em ORIGEM_DIR
      2) carregar Excel e XML
      3) injetar PLMNs ausentes (exceto rejeitados)
      4) limpar expiradas + aplicar regras
      5) montar finais e salvar XLSX+XML em DESTINO_DIR
    """
    try:
        registrar_log("Iniciado", "Sucesso", "Iniciar atualização do loader.")
        arquivo_origem_xlsx = next((f for f in os.listdir(ORIGEM_DIR) if "tap" in f.lower() and f.lower().endswith(".xlsx")), None)
        arquivo_origem_xml  = next((f for f in os.listdir(ORIGEM_DIR) if "rate" in f.lower() and f.lower().endswith(".xml")), None)

        if not arquivo_origem_xlsx or not arquivo_origem_xml:
            registrar_log("Iniciado", "Erro", "Arquivo de origem (xlsx ou xml) não encontrado.")
            print("Erro: Arquivo de origem ou destino não encontrado.")
        else:
            caminho_origem_xlsx = os.path.join(ORIGEM_DIR, arquivo_origem_xlsx) # arquivo xlsx
            caminho_origem_xml  = os.path.join(ORIGEM_DIR, arquivo_origem_xml)  # arquivo xml

            registrar_log("Iniciado", "Sucesso", f"Arquivo origem XLSX: {arquivo_origem_xlsx}")
            print(f"Arquivo origem XLSX: {arquivo_origem_xlsx}")
            registrar_log("Iniciado", "Sucesso", f"Arquivo origem XML: {arquivo_origem_xml}")
            print(f"Arquivo origem XML: {arquivo_origem_xml}")

            # 1) Ler Excel consolidado (fonte de tarifas)
            dados_origem = carregar_dados_excel(caminho_origem_xlsx)

            # 2) Ler XML origem (multi-aba) -> dict de planilhas
            planilhas, tree, root, ns = carregar_xml_para_dfs(caminho_origem_xml)

            # 2.0) Limpeza antecipada de tarifas expiradas
            for nome, info in planilhas.items():

                df = info.get("df")

                if df is None or df.empty:
                    continue

                planilhas[nome]["df"] = _limpar_tarifas_expiradas(df)

            # 2.1) Injetar PLMNs ausentes 
            try:
                _injetar_plmns_ausentes(planilhas, dados_origem, SAIDA_DIR)
            except Exception as e:
                registrar_log("Finalizado", "Erro", f"Falha ao injetar PLMNs ausentes: {e}")

            # 2.2) Completar serviços faltantes
            try:
                _completar_servicos(planilhas, dados_origem)
            except Exception as e:
                registrar_log("Finalizado", "Erro", f"Falha ao completar serviços faltantes: {e}")

            # 3) Limpeza e Regras por aba
            planilhas_transformadas = {}
            for nome, info in planilhas.items():
                df_destino = info["df"]
                if df_destino is None or df_destino.empty:
                    # Mantém estrutura/cabecalho mesmo sem dados
                    planilhas_transformadas[nome] = {
                        "df": df_destino,
                        "estrutura": info["estrutura"],
                        "cabecalho": info["cabecalho"]
                    }
                    continue

                # df_destino = _limpar_tarifas_expiradas(df_destino)
                df_destino = _aplicar_regras(df_destino, dados_origem, cabecalho=info["cabecalho"])

                planilhas_transformadas[nome] = {
                    "df": df_destino,
                    "estrutura": info["estrutura"],
                    "cabecalho": info["cabecalho"]
                }

            # 4) Montar DataFrames finais por aba (estrutura + cabeçalho + dados)
            planilhas_finais = _montar_planilhas_finais(planilhas_transformadas)

            # 5) Salvar XLSX atualizado (multi-aba)
            # Nome base derivado do arquivo XML de origem (ex.: RATE_JAN26.XML -> RATE_JAN26)
            nome_base_xml = os.path.splitext(arquivo_origem_xml)[0]   # remove extensão
            nome_xlsx_atualizado = f"{nome_base_xml}_atualizado.xlsx"
            destino_xlsx = os.path.join(DESTINO_DIR, nome_xlsx_atualizado)
            salvar_xlsx_multiplas_abas(destino_xlsx, planilhas_finais)
            registrar_log("Finalizado", "Sucesso", f"XLSX multi-aba atualizado salvo em: {destino_xlsx}")
            print(f"Processo concluído. Arquivo XLSX atualizado salvo em {destino_xlsx}")

            # 6) Atualizar o XML final (multi-aba) preservando nomes das sheets
            base_name, ext = os.path.splitext(os.path.basename(caminho_origem_xml))
            caminho_xml_destino = os.path.join(DESTINO_DIR, base_name + ext)
            atualizar_xml_multiplas_abas(tree, root, ns, caminho_xml_destino, planilhas_finais, sobrescrever_formulas=False)

            return {"xlsx": destino_xlsx, "xml": caminho_xml_destino}

    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Erro no processo principal de atualização do loader: {e}")