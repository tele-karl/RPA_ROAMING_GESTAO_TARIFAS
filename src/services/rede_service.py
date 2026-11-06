import os
import sys
import shutil
from datetime import datetime, timedelta
import openpyxl
from src.services.log_service import registrar_log
from src.utils.excel_utils import parse_date_generic
from src.configs.configuration import (
    ORIGEM_DIR, REDE_NOVOS_DIR, REDE_PROCESSADOS_DIR, REDE_EXCEPCIONAIS_DIR,
)

ABAS_VALIDAS = [
    "brasil base disc rates",
    "vivo base disc rates",
    "brasil int. call disc",
    "vivo int. call disc",
]

def listar_arquivos_rede() -> list[str]:
    """
    Lista arquivos .xlsx no diretório de rede (origem). 
    """
    print("[INFO]", "Listando arquivos no diretório de rede (origem)...")
    if not os.path.exists(REDE_NOVOS_DIR):
        print("[ERROR]", f"Diretório de rede não encontrado: {REDE_NOVOS_DIR}")
        registrar_log("Iniciado", "Erro", f"Diretório de rede não encontrado: {REDE_NOVOS_DIR}")
        return []
    arquivos = [f for f in os.listdir(REDE_NOVOS_DIR) if f.lower().endswith('.xlsx')]
    print("Finalizado", "Sucesso", f"{len(arquivos)} arquivos encontrados.")
    return arquivos

def verificar_effective_date(filepath: str):
    """
    Verifica a 'Effective Date' nas abas Base Disc Rates e Int. Call Disc.
    1) Coluna D (principal) (prioridade)
    2) Coluna fallback:
       - Base Disc Rates: coluna AG (33)
       - Int. Call Disc: coluna R (18)
    A busca respeita o início real de dados:
       - Base Disc Rates: linha 16
       - Int. Call Disc: linha 12
    Retorna um date (ou None se nada encontrado).
    """
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        # Primeira tentativa: coluna D
        for aba in wb.sheetnames:
            aba_lc = aba.lower()
            if "base disc rates" in aba_lc:
                start_row, col_idx = 16, 4  # D
            elif "int. call disc" in aba_lc:
                start_row, col_idx = 12, 4  # D
            else:
                continue
            ws = wb[aba]
            for r in range(start_row, ws.max_row + 1):
                d = parse_date_generic(ws.cell(row=r, column=col_idx).value)
                if d:
                    wb.close()
                    return d
        # Segunda tentativa (fallback): AG para Base Disc/ R para Int Call
        for aba in wb.sheetnames:
            aba_lc = aba.lower()
            if "base disc rates" in aba_lc:
                start_row, col_idx = 16, 33  # AG
                col_name = "AG"
            elif "int. call disc" in aba_lc:
                start_row, col_idx = 12, 18  # R
                col_name = "R"
            else:
                continue
            ws = wb[aba]
            for r in range(start_row, ws.max_row + 1):
                d = parse_date_generic(ws.cell(row=r, column=col_idx).value)
                if d:
                    print("[WARN]", f"Effective Date não encontrada na coluna D; usando fallback {col_name} na aba '{aba}': {d}")
                    wb.close()
                    return d
        # Nenhuma data válida em D, nem nos fallbacks
        wb.close()
        return None
    except Exception as e:
        registrar_log("Iniciado", "Erro", f"Falha ao verificar Effective Date em {filepath}: {e}")
        return None

def arquivo_tem_abas_obrigatorias(filepath: str) -> bool:
    """
    Verifica se o arquivo contém pelo menos uma das abas obrigatórias.
    Caso não contenha nenhuma, retorna False.
    """
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        abas_lower = [a.lower().strip() for a in wb.sheetnames]
        resultado = any(aba in abas_lower for aba in ABAS_VALIDAS)
        wb.close()
        return resultado
    except Exception as e:
        registrar_log("Iniciado", "Erro", f"Falha ao verificar abas obrigatórias em {filepath}: {e}")
        return False

def copiar_arquivos_validos() -> None:
    """
    Copia arquivos válidos do REDE_DIR para ORIGEM_DIR e REDE_PROCESSADOS_DIR.
    Critério: Effective Date <= hoje + 30 dias. 
    Abas obrigatórias são levadas em consideração. Arquivos mais antigos do que 6 meses
    são tratados exepcionalmente. 
    """
    try:
        print("[INFO]", "Iniciando verificação e cópia de arquivos válidos...")
        hoje_date = datetime.today().date()
        limite_date = hoje_date + timedelta(days=30)  # D+30
        limite_excepcional = hoje_date - timedelta(days=180)  # 6 meses atrás

        arquivos = listar_arquivos_rede()

        if not arquivos:
            mensagem = "Nenhum arquivo encontrado em 'Arquivos_tarifas_novos_TGR' da rede. Processo encerrado."
            print("[INFO]", mensagem)
            registrar_log("Finalizado", "Erro", mensagem)
            sys.exit(0)  # encerra o RPA de forma controlada

        os.makedirs(ORIGEM_DIR, exist_ok=True)
        os.makedirs(REDE_PROCESSADOS_DIR, exist_ok=True)
        os.makedirs(REDE_EXCEPCIONAIS_DIR, exist_ok=True)

        arquivos_copiados = 0

        for arquivo in arquivos:
            caminho_origem = os.path.join(REDE_NOVOS_DIR, arquivo)

            # Verificar abas obrigatória
            if not arquivo_tem_abas_obrigatorias(caminho_origem):
                print("[WARN]", f"Arquivo '{arquivo}' não possui abas obrigatórias. Será IGNORADO (não será processado).")
                registrar_log("Iniciado", "Aviso", f"Arquivo '{arquivo}' não possui abas obrigatórias. Ignorado.")
                continue
            
            # Verificação de Effective Date
            data_efetiva = verificar_effective_date(caminho_origem)
            if not data_efetiva:
                print("[WARN]", f"Arquivo '{arquivo}' sem Effective Date válida (mantido na rede).")
                registrar_log("Iniciado", "Erro", f"Arquivo '{arquivo}' sem Effective Date válida (mantido na rede).")
                # mantém o arquivo na pasta origem da rede (não move)
                continue
            
            # Arquivo com mais de 6 meses (EXCEPCIONAIS)
            if data_efetiva < limite_excepcional:
                print("[WARN]", f"Arquivo '{arquivo}' possui Effective Date abaixo da tolerancia de 6 meses ({data_efetiva})."
                    "Movendo para diretório de EXCEPCIONAIS.")
                try:
                    destino_ex = os.path.join(REDE_EXCEPCIONAIS_DIR, arquivo)
                    shutil.move(caminho_origem, destino_ex)
                    registrar_log("Finalizado", "Sucesso", f"Arquivo excepcional '{arquivo}' movido com sucesso.")
                except Exception as e:
                    registrar_log("Iniciado", "Erro", f"Falha ao mover arquivo excepcional '{arquivo}': {e}")
                continue
            
            # Regra D+30
            if data_efetiva <= limite_date:
                destino_local = os.path.join(ORIGEM_DIR, arquivo)
                destino_rede = os.path.join(REDE_PROCESSADOS_DIR, arquivo)
                try:
                    # Copiar para o diretório local de origem do processamento
                    shutil.copy2(caminho_origem, destino_local)
                    arquivos_copiados += 1
                    print("[INFO]", f"Cópia do arquivo {arquivo} realizada para {destino_local}")
                except Exception as e:
                    print("[ERROR]", f"Falha na cópia para ORIGEM_DIR: {arquivo}: {e}")
                    continue    # sem cópia local, não seguimos para o move
                try:
                    # Move da pasta REDE_NOVOS_DIR e coloca em REDE_PROCESSADOS_DI
                    shutil.move(caminho_origem, destino_rede)
                    print("[INFO]", f"{arquivo} aprovado pelo critério D+30. Movido para {destino_rede}")
                except Exception as e:
                    print("[ERROR]", f"Falha ao mover '{arquivo}': {e}")
            else:
                print("[WARN]", f"Arquivo '{arquivo}' fora do critério D+30. Mantido no diretório de origem da rede." 
                    f"(Effective Date: {data_efetiva}).")

        if arquivos_copiados == 0:
            mensagem = f"Nenhum arquivo válido (critério D+30) encontrado para construção da planilha consolidada. Processo encerrado."
            print("[WARN]", mensagem)
            registrar_log("Finalizado", "Erro", mensagem)
            sys.exit(0)
        
        registrar_log("Finalizado", "Sucesso", f"Arquivos copiados para {ORIGEM_DIR}")

    except Exception as e:
        registrar_log("Iniciado", "Erro", f"Falha na copia dos arquivos para {ORIGEM_DIR}: {e}")