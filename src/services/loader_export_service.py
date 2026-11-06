import os
import re
import time
from pathlib import Path
from typing import Iterable

from pywinauto import Application
from pyjab.jabdriver import JABDriver

import pandas as pd
import pyautogui
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from src.services.log_service import registrar_log

from src.configs.configuration import (
    ORIGEM_DIR, SAIDA_DIR, USUARIO_ICT, SENHA_ICT,
)
# Padrões e constantes das planilhas TAP
PADRAO_ARQ = re.compile(r"^TAP\s+([A-Za-zÀ-ÿ]{3})-(\d{2})\.xlsx$", re.IGNORECASE)
PAT_TAP = re.compile(r"^TAP\s+.+\.xlsx$", re.IGNORECASE)
SHEETS_BASE = ["VIVO BASE DISC RATES", "BRASIL BASE DISC RATES"]
SHEETS_INT = ["VIVO INT. CALL DISC", "BRASIL INT. CALL DISC"]

# Configs para detectar região de alerta visual em VERMELHO (pixels)
# Região do alerta (x, y, width, height) — AJUSTAR se necessário
REGION_ALERTA_PLMN_REF = (429, 742, 12, 5)  
RED_MIN, GREEN_MAX, BLUE_MAX = 180, 90, 90
PIXEL_COUNT_MIN = 20
# Resolução de referência (1920x1080)
REF_W, REF_H = 1920, 1080

def _base_dir() -> Path:
    """ Determina diretório base """
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd()

def construir_nome_loader() -> str:
    """
    Constrói o nome do loader para colocar no portal a partir do arquivo TAP
    (planilha consolidada)
    """
    origem = Path(ORIGEM_DIR)
    candidatos = sorted(
        [p for p in origem.glob("*.xlsx") if PADRAO_ARQ.match(p.name)],
        key=lambda x: x.stat().st_mtime,
        reverse=True
    )
    if not candidatos:
        registrar_log("Iniciado", "Erro", f"Nenhum arquivo no padrão 'TAP Mmm-YY.xlsx' em data/arquivos_origem")
        raise
    
    nome = candidatos[0].name
    mes_raw, ano = PADRAO_ARQ.match(nome).groups()
    mes3 = mes_raw[:3].upper()     # pega as 3 primeiras letras e põe em maiúsculas
    return f"RATE_{mes3}{ano}"

def _sheet_por_contem(sheet_names, candidatos_upper: Iterable[str]):
    """Retorna o primeiro nome de aba em writer_sheets cujo nome contém
    algum termo de 'candidatos' (case-insensitive)."""
    for s in sheet_names:
        s_up = str(s).upper()
        for alvo in candidatos_upper:
            if alvo in s_up:
                return s
    return None

# FUNÇÃO PARA LEITURA DE PLMNs
def obter_plmns_da_planilha() -> list[str]:
    """Lê PLMNs únicos das abas de interesse (coluna A),
    conforme linhas iniciais distintas para cada aba.
    - Base Disc Rates: a partir da linha 16 (1-based)
    - Int. Call Disc : a partir da linha 12 (1-based)
    Retorna: lista ordenada de PLMNs (strings).
    """
    try:
        origem = Path(ORIGEM_DIR)
        candidatos = [p for p in origem.glob("*.xlsx") if PAT_TAP.match(p.name)]
        if not candidatos:
            registrar_log("Iniciado", "Erro", f"Nenhum arquivo TAP *.xlsx encontrado em data/arquivos_origem")
            raise
        candidatos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        caminho = candidatos[0]

        # Descobrir folhas disponíveis sem carregar tudo
        xl = pd.ExcelFile(caminho, engine="openpyxl")
        # Encontrar abas alvo
        aba_base = _sheet_por_contem(xl.sheet_names, [s.upper() for s in SHEETS_BASE])
        aba_int = _sheet_por_contem(xl.sheet_names, [s.upper() for s in SHEETS_INT])

        plmns = set()

        def _coletar(aba, linha_inicio_1based):
            if not aba:
                return
            df = pd.read_excel(
                caminho, 
                sheet_name=aba, 
                usecols="A", 
                header=None, 
                engine="openpyxl"
            )
            # Converter para série, descartar linhas acima do início
            s = df.iloc[max(0, linha_inicio_1based - 1):, 0]
            # Normalizar: str, strip, maiúscula
            s = s.dropna().astype(str).str.strip()
            s = s[s != ""]
            for v in s:
                plmns.add(v)
    except Exception as e:
        registrar_log("Iniciado", "Erro", f"Falha ao tentar ler PLMNs no arquivo TAP consolidado: {e}")

    # Coleta com os recortes definidos
    _coletar(aba_base, 16)
    _coletar(aba_int, 12)
    # Retorna lista ordenada (alfabética) apenas para previsibilidade
    return sorted(plmns)

# VALIDAÇÃO VISUAL DE PLMN NÃO CADASTRADO
def normalizar_regiao_por_tela_atual(region_ref, w_ref=REF_W, h_ref=REF_H):
    # Tamanho reportado pelo pyautogui
    w_atual, h_atual = pyautogui.size()
    rx, ry, rw, rh = region_ref
    x = int(rx * w_atual / w_ref)
    y = int(ry * h_atual / h_ref)
    w = max(1, int(rw * w_atual / w_ref))
    h = max(1, int(rh * h_atual / h_ref))
    return (x, y, w, h)

def alerta_vermelho_presente(
    region_ref=REGION_ALERTA_PLMN_REF,
    red_min=RED_MIN, green_max=GREEN_MAX, blue_max=BLUE_MAX,
    pixel_count_min=PIXEL_COUNT_MIN
) -> tuple[bool, tuple[int,int,int,int], int]:
    """
    Retorna True se houver um número significativo de pixels vermelhos na região informada.
    Use para detectar o alerta visual (em vermelho) que o sistema mostra quando o PLMN não existe.
    Os thresholds podem ser ajustados de acordo com o tema/monitor.
    """
    try:
        region = normalizar_regiao_por_tela_atual(region_ref)
        img = pyautogui.screenshot(region=region).convert("RGB")   # PIL.Image
        cnt = 0
        for r, g, b, *rest in img.getdata():
            if r >= red_min and g <= green_max and b <= blue_max:
                cnt += 1
            if cnt >= pixel_count_min:
                return True
        return False
    except Exception as e:
        registrar_log("Iniciado", "Erro", f"Erro na detecção de sinal de Alerta Vermelho de PLMNs rejeitado: {e}")

# FUNÇÃO PARA CRIAR ARQUIVO COM PLMN(s) NÃO CADASTRADOS
def resumir_plmns_rejeitados(plmns_rejeitados, nome_loader, base_dir: Path | None = None):
    """
    Gera e salva o resumo de PLMNs rejeitados em data/saida/plmns_rejeitados_<nome_loader>.txt.
    
    Parâmetros:
    plmns_rejeitados : Iterable[str]
        Lista (ou set) de PLMNs rejeitados detectados no portal ICT.
    nome_loader : str
        Nome lógico do loader (ex.: RATE_MAR26), usado no nome do arquivo de saída.
    base_dir : Path | None
        Diretório base do projeto. Se None, usa o diretório do script atual.

    Retorna:
    dict com chaves:
        - "path": Path do arquivo gerado
        - "qtd": int quantidade de PLMNs rejeitados
        - "vazio": bool True se lista estiver vazia
    """
    # Determina base_dir
    base_dir = base_dir or _base_dir()
    saida_dir = Path(SAIDA_DIR)
    saida_dir.mkdir(parents=True, exist_ok=True)

    # Normaliza lista (strings sem espaços extras)
    rejeitados_norm = [str(x).strip() for x in plmns_rejeitados if str(x).strip() != ""]
    qtd = len(rejeitados_norm)
    vazio = (qtd == 0)

    # Monta caminho do arquivo
    nome_txt = f"plmns_rejeitados_{nome_loader}.txt"
    rejeitados_path = saida_dir / nome_txt

    # Salva apenas se houver algo, mas cria arquivo mesmo vazio
    try:
        with rejeitados_path.open("w", encoding="utf-8") as f:
            for plmn in rejeitados_norm:
                f.write(plmn + "\n")
        registrar_log("Iniciado", "Sucesso", f"PLMNs Rejeitados: {qtd} (salvos em {rejeitados_path})")
        print(f"[INFO] Rejeitados: {qtd} (salvos em {rejeitados_path})")
    except Exception as e:
        registrar_log("Iniciado", "Sucesso", f"Falha ao salvar lista de rejeitados: {e}")
        print(f"[WARN] Falha ao salvar lista de rejeitados: {e}")

    return {"path": rejeitados_path, "qtd": qtd, "vazio": vazio}

def _configurar_chrome(download_dir: Path) -> webdriver.Chrome:
    """Configuração do Chrome"""
    try:
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        # Ignorar erros de certificado SSL (comum em sites internos/homologação)
        chrome_options.add_argument("--ignore-certificate-errors")
        prefs = {
            "download.default_directory": str(download_dir.resolve()),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "safebrowsing.disable_download_protection": True,
        }
        chrome_options.add_experimental_option("prefs", prefs)
        # Argumentos adicionais para tentar inibir o bloqueio
        chrome_options.add_argument("--safebrowsing-disable-download-protection")
        chrome_options.add_argument("--safebrowsing-disable-extension-blacklist")
        # Inicializa o driver
        driver = webdriver.Chrome(options=chrome_options)
        return driver

    except Exception as e:
        registrar_log("Iniciado", "Falha", f"Falha na configuração do Chrome: {e}")

def executar_export_loader() -> dict:
    """
    Executa o fluxo completo do export do Loader no portal ICT.
    Retorna o 'resumo' {'path': Path, 'qtd': int, 'vazio': bool} de PLMNs rejeitados.
    """
    # Preparar pasta de download
    script_dir = _base_dir()
    download_dir = script_dir / "download"
    download_dir.mkdir(parents=True, exist_ok=True)
    # Limpa possíveis resíduos
    for f in download_dir.glob("*"):
        try:
            f.unlink()
        except Exception:
            pass

    driver = _configurar_chrome(download_dir)
    try:
        url = "https://roamingict.redecorp.br:14100/portal/home"
        print(f"Inciando abertura do portal do ICT. Acessando: {url}")
        driver.get(url)

        # Aguarda até 20 segundos pelos elementos aparecerem
        wait = WebDriverWait(driver, 20)
        # Login (ID/name/xpath fallback)
        try:
            # Tentativa 1: Procurar por ID 'username'
            campo_usuario = wait.until(EC.presence_of_element_located((By.ID, "username")))
        except:
            try:
                # Tentativa 2: Procurar por name 'j_username' (comum em portais Java/JBoss)
                campo_usuario = driver.find_element(By.NAME, "j_username")
            except:
                # Tentativa 3: Procurar por ID 'user'
                campo_usuario = driver.find_element(By.ID, "user")
        campo_usuario.clear()
        campo_usuario.send_keys(USUARIO_ICT)

        try:
            # Tentativa 1: Procurar por ID 'password'
            campo_senha = driver.find_element(By.ID, "password")
        except:
            try:
                # Tentativa 2: Procurar por name 'j_password'   
                campo_senha = driver.find_element(By.NAME, "j_password")
            except:
                # Tentativa 3: Procurar qualquer input type password
                campo_senha = driver.find_element(By.XPATH, "//input[@type='password']")
        campo_senha.clear()
        campo_senha.send_keys(SENHA_ICT)

        # Submeter
        try:
            # Procura botão de submit
            botao_login = driver.find_element(
                By.XPATH,
                "//button[@type='submit'] | //input[@type='submit'] | "
                "//button[contains(text(), 'Login')] | //button[contains(text(), 'Entrar')]"
            )
            botao_login.click()
        except:
            campo_senha.send_keys(Keys.RETURN)

        # Espera e clica em "CSG INTERCONNECT"
        # Tenta encontrar pelo texto exato ou parcial
        # XPath: procura qualquer elemento que contenha o texto "CSG INTERCONNECT"
        elemento_csg = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'CSG INTERCONNECT')]"))
        )
        elemento_csg.click()

        # (Opcional) chrome://downloads para aprovar download perigoso (shadow DOM) — mantido
        time.sleep(3)
        try:
            driver.execute_script("window.open()")
            driver.switch_to.window(driver.window_handles[-1])
            driver.get("chrome://downloads")
            time.sleep(3)
            script_aprovar = """
            var manager = document.querySelector('downloads-manager');
            if (manager) {
                var item = manager.shadowRoot.querySelector('downloads-item');
                if (item) {
                    var shadow = item.shadowRoot;
                    var btnKeep = shadow.querySelector('cr-button[focus-type="save"]');
                    var btnDangerous = shadow.querySelector('cr-button#dangerous');
                    if (btnDangerous) { btnDangerous.click(); return "Dangerous"; }
                    if (btnKeep) { btnKeep.click(); return "Keep"; }
                }
            }
            return "None";
            """
            resultado = driver.execute_script(script_aprovar)
            # Se clicou, tentar confirmar
            if "Dangerous" in resultado:
                time.sleep(1)
                # Tenta confirmar novamente se necessário
                driver.execute_script("""
                var manager = document.querySelector('downloads-manager');
                var item = manager.shadowRoot.querySelector('downloads-item');
                var shadow = item.shadowRoot;
                var btnProceed = shadow.querySelector('cr-button.action-button');
                if (btnProceed) btnProceed.click();
                """)
            # Fecha a aba de downloads e volta para a principal
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

        except Exception as e_down:
            print(f"[WARN] chrome://downloads não acessível ou não necessário: {e_down}")
            # Garante retorno para janela principal
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

        # Espera arquivo baixar e executar .jnlp
        print("Aguardando arquivo ser baixado...")
        arquivo_baixado = None
        for _ in range(60):
            arquivos = [f for f in os.listdir(download_dir) if not f.endswith('.crdownload') and not f.endswith('.tmp')]
            if arquivos:
                arquivo_baixado = download_dir / arquivos[0]
                break
            time.sleep(1)

        if not arquivo_baixado:
            registrar_log("Iniciado", "Erro", f"Tempo limite esgotado. Nenhum arquivo de download encontrado.")
            print("Tempo limite esgotado. Nenhum arquivo de download encontrado.")
            return {"path": Path(SAIDA_DIR) / "plmns_rejeitados_.txt", "qtd": 0, "vazio": True}

        # Executa arquivo (JNLP) — fluxo original
        try:
            os.startfile(arquivo_baixado)
            time.sleep(7)
            # Clique (ou comando) para EVITAR ATUALIZAÇÃO DO JAVA (!)
            pyautogui.click(961, 587)
            
            ## PARA CONEXÃO VIA VPN
            time.sleep(10) 
            pyautogui.press('enter')
            # pyautogui.click(1092, 610)

            # ## PARA CONEXÃO VIA VIVO_CORP
            # # Aguarda o Java carregar e exibir a janela de segurança
            # time.sleep(135)
            # # Selecionar posição "Eu aceito o risco e desejo..."
            # pyautogui.press('tab')
            # pyautogui.press('tab')
            # pyautogui.press('space')
            # time.sleep(0.1)
            # # Na imagem, o botão "Executar" já aparece focado após aceitar
            # # Então pressionar Enter deve acioná-lo.
            # pyautogui.press('enter')

            # Para fazer com clique
            # pyautogui.click(611, 729)
            # time.sleep(1)
            # # Clique posição "Executar"
            # pyautogui.click(1141, 727)
            # time.sleep(1)

            # Tempo para o Java baixar os recursos e abrir a janela principal
            time.sleep(10)  # Ajustar conforme modo de entrada: 10s para VIVO_CORP, 20s para VPN 

            # Adaptação da Interação com a Aplicação do CSG via PYJAB
            app = JABDriver(
                title="CSG Interconnect",
                bridge_dll=r"C:\Program Files\Java\jre1.8.0_481\bin\WindowsAccessBridge-64.dll"
            )

            app_win = Application().connect(
                title_re=".*CSG Interconnect.*"
            )

            # Acessar Bulk Export
            tools = next(e for e in app.find_elements() if e.name == "Tools")
            tools.click()

            bulk_load = next(e for e in app.find_elements() if e.name == "Bulk Load")
            bulk_load.click()

            export = next(e for e in app.find_elements() if e.name == "Export")
            export.click()

            # Selecionar Filter (para aparecer a Lista)
            filter_btn = next(
                e for e in app.find_elements() if e.name == "Filter" and e.role == "push button"
            )
            filter_btn.click()

            # Selecionar Add New
            add_new = next(e for e in app.find_elements() if e.name == "Add New")
            add_new.click()
            
            # Inserir o nome do loader baseado no nome da planilha consolidada "TAP Mmm-AA.xlsx"    
            nome_loader = construir_nome_loader()  # baseado em data/arquivos_origem/TAP Mmm-YY.xlsx
            # Digita o conteúdo da variável nome_loader
            campos = [e for e in app.find_elements() if e.role == "text"]
            campo_nome = campos[15]
            campo_nome.send_text(nome_loader)

            # Seleciona aba Job Parameters
            aba_job_param = next(
                e for e in app.find_elements()
                if e.role == "page tab list"
                and e.name == "Detail"
            )

            aba_job_param.select("Job Parameters")

            # Selecionar tipo de loader (Rate - Enhanced)
            loader = app.find_element_by_name("Loader")

            elementos = loader.parent.find_elements()

            idx_loader = next(
                i for i, e in enumerate(elementos)
                if e.role == "label" and e.name == "Loader"
            )

            loader_combo = elementos[idx_loader - 1]

            loader_combo.select("Rate - Enhanced")

            # Clique no campo File name
            label_file = app.find_element_by_name("File name")

            campo_file_name = next(
                filho for filho in label_file.parent.find_elements()
                if filho.role == "text"
            )

            campo_file_name.send_text(
                nome_loader + ".xml"
            )
            
            # Clique no botão com (...)
            botao_data_selection = next(
                filho
                for filho in app.find_elements()
                if (
                    filho.role == "push button"
                    and filho.name == "..."
                    and filho.parent.name == "Data Selection"
                )
            )

            botao_data_selection.click()

            # Selecionar TARIFA_1 (da aba Generic Selection) e clicar em Edit
            def selecionar_tarifa(app, app_win, id_tarifa):
                tabelas = [
                    e for e in app.find_elements()
                    if e.role == "table"
                ]

                for tabela in tabelas:
                    try:
                        total_linhas = (
                            tabela.get_element_information()
                            ["table"]["row_count"]
                        )

                        for row in range(total_linhas):
                            cell = (tabela._get_accessible_table_cell_info(row, 0))
                            info = (
                                tabela._get_accessible_context_info(
                                    cell.accessibleContext
                                )
                            )

                            if info.name == id_tarifa:
                                tabela._request_focus()
                                tabela._add_accessible_selection_from_context(
                                    cell.index
                                )

                                janela = app_win.top_window()
                                
                                janela.type_keys("{DOWN}")
                                janela.type_keys("{UP}")

                                return True

                    except Exception:
                        pass

                return False

            edit = next(
                e for e in app.find_elements()
                if e.role == "push button"
                and e.name == "Edit"
                and e.bounds["x"] > 0
            )

            if selecionar_tarifa(app, app_win, "TARIFA_1"):
                edit.click()

            # Remover query 
            while True:
                itens = [
                    e for e in app.find_elements()
                    if (
                        e.role == "label"
                        and "Billing Operator contains" in e.name
                    )
                ]

                if not itens:
                    break

                itens[0].click()

                remove = next(
                    e for e in app.find_elements()
                    if (
                        e.role == "push button"
                        and e.name == "Remove"
                        and e.get_element_information()["bounds"]["x"] > 0
                    )
                )
            
                remove.click()

            time.sleep(0.5)

            # Inserir PLMNs
            # Selecionar Campo Criteria -> Billing Oerator  
            for _ in range(1): 
                pyautogui.press('tab')
                time.sleep(0.3)
            for _ in range(2): 
                pyautogui.press('down')
                time.sleep(0.3)

            # Selecionar Contains para filtrar Operadoras 
            # pyautogui.click(201, 731)

            for _ in range(1): 
                pyautogui.press('tab')
                time.sleep(0.3)

            for _ in range(8): 
                pyautogui.press('down')
                time.sleep(0.3)

            # Clique de segurança
            pyautogui.click(414, 742)

            plmns = obter_plmns_da_planilha()
            plmns_rejeitados = []

            for plmn in plmns:
                # Clique para selecionar o local onde é inserido os PLMNs 
                pyautogui.click(414, 742)
                pyautogui.typewrite(str(plmn))  # Escreve PLMN
                # Inserir apertando o botão ADD
                for _ in range(2): 
                    pyautogui.press('tab')
                    time.sleep(0.5)

                # PONTO DE VERIFICAÇÃO DO ALERTA (SCREENSHOT DO SINAL VERMELHO)
                if alerta_vermelho_presente():
                    print(f"PLMN rejeitado pelo sistema: {plmn}")
                    plmns_rejeitados.append(str(plmn))
                    time.sleep(0.5)
                    # Segue para o próximo PLMN
                    continue
                # Se não há alerta confirma ADD
                pyautogui.press('enter'); time.sleep(1)

            # Voltar ao campo de entrada (Shift+Tab 2x) e deletar conteúdo
            pyautogui.keyDown('shift')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.keyUp('shift')
            pyautogui.press('delete')

            # SALVAR RESUMO DE PLMNs REJEITADOS 
            try:
                # base_dir: diretório do script atual (mesma lógica já usada no arquivo)
                resumo = resumir_plmns_rejeitados(plmns_rejeitados, nome_loader, base_dir=script_dir)
                # resumo = {"path": Path(...), "qtd": int, "vazio": bool}
            except Exception as _e:
                registrar_log("Iniciado", "Erro", f"Falha ao gerar resumo de PLMNs rejeitados: {_e}")
                print(f"[WARN] Falha ao gerar resumo de PLMNs rejeitados: {_e}")
                resumo = {"path": Path(SAIDA_DIR) / f"plmns_rejeitados_{nome_loader}.txt", "qtd": len(plmns_rejeitados), "vazio": len(plmns_rejeitados) == 0}

            # Inserir query e confirmar
            for _ in range(8): 
                pyautogui.press('tab')
                time.sleep(0.3)
            pyautogui.press('enter'); time.sleep(0.3) # aciona botão Insert
            pyautogui.press('tab'); time.sleep(0.3)   # acessar botão OK

            for _ in range(2): 
                pyautogui.press('enter')
                time.sleep(1.5)

            # VOLTAR PARA A ABA BULK EXPORT
            pyautogui.keyDown('ctrl'); time.sleep(0.3)
            pyautogui.press('tab'); time.sleep(0.3)
            pyautogui.press('tab'); time.sleep(0.3)
            pyautogui.keyUp('ctrl'); time.sleep(0.5)

            # Selecionar TARIFA_1 (da aba Bulk Export)
            botao = app.find_element_by_name("...")

            painel = botao.parent

            combo = next(
                e for e in painel.find_elements()
                if e.role == "combo box"
            )

            combo.select("TARIFA_1-tarifas")

            pyautogui.click(1590, 996); time.sleep(1)   # Clique OK
            pyautogui.click(1321, 763); time.sleep(5)   # Clique Status Monitor

            # Procedimento dentro da aba Job Status
            pyautogui.click(43, 266); time.sleep(12)   # Clica no + da pasta Bulk Export
            pyautogui.click(1712, 176); time.sleep(0.5) # Aciona Filter (apenas para atualizar status)
            pyautogui.click(64, 282); time.sleep(0.5)   # Clica no + da pasta Rate (para acompanhar status)

            status = ""
            while status != "Successful":
                time.sleep(30)
                pyautogui.click(1712, 176) # Aciona Filter (apenas para atualizar status)
                time.sleep(1)

                pyautogui.click(384, 280)  # Clica em Status
                time.sleep(0.5)

                pyautogui.keyDown('ctrl'); time.sleep(0.3)
                pyautogui.press('c'); time.sleep(0.3)
                pyautogui.keyUp('ctrl')
                time.sleep(0.5)

                status = pyperclip.paste() # Lê texto capturado

                if status == "Failed":
                    registrar_log("Finalizado", "Erro", "Falha ao exportar loader. Status = Failed (!)")
                    break                    
            
            time.sleep(2)  
            # Pressiona ALT para ativar a barra de menu
            pyautogui.press('alt'); time.sleep(1)
            # Navega para a direita até Tools (8 vezes)
            # File -> Edit -> View -> Agreement -> Network -> Billing -> Fin.Mgmt -> Admin -> Tools
            for _ in range(8): 
                pyautogui.press('right')
                time.sleep(0.3)
            # Abre o menu Tools pressionando para BAIXO
            pyautogui.press('down'); time.sleep(0.3)
            # Deve estar em "Bulk Load". Seta para DIREITA abre o submenu
            pyautogui.press('right'); time.sleep(0.3)
            # Desce para "View Load Results"
            for _ in range(7): 
                pyautogui.press('down')
                time.sleep(0.3)
            # Seta para DIREITA abre o submenu.
            pyautogui.press('right'); time.sleep(0.3)
            for _ in range(2): 
                pyautogui.press('down')
                time.sleep(0.3)
            pyautogui.press('enter'); time.sleep(2)  # Abre Bulk Export Results
            # Selecionar Filter 
            for _ in range(6): 
                pyautogui.press('tab')
                time.sleep(0.3)
            pyautogui.press('enter'); time.sleep(2)
            # Clique em Transer File
            pyautogui.click(1829, 264); time.sleep(3)

            # Selecionar local de salvamento (arquivos_origem)
            app_salvar = JABDriver(
                title="Save",
                bridge_dll=r"C:\Program Files\Java\jre1.8.0_481\bin\WindowsAccessBridge-64.dll"
            )

            # Campo "Nome do arquivo:"
            campo_arquivo = next(
                e for e in app_salvar.find_elements()
                if (
                    e.role == "text"
                    and e.name == "File name:"
                )
            )

            caminho_origem = (
                rf"{ORIGEM_DIR}\{nome_loader}.xml"
            )

            campo_arquivo.send_text(
                caminho_origem
            )

            # Botão Salvar
            botao_salvar = next(
                e for e in app_salvar.find_elements()
                if (
                    e.role == "push button"
                    and e.name == "Save"
                )
            )

            botao_salvar.click()

            # Fechar abas abertas do ICT
            for _ in range(4):
                pyautogui.click(235, 60)
                time.sleep(0.7)

            return resumo

        except Exception as e_exec:
            registrar_log("Iniciado", "Erro", f"Erro ao executar arquivo: {e_exec}")
            print(f"Erro ao executar arquivo: {e_exec}")
            return {"path": Path(SAIDA_DIR) / "plmns_rejeitados_.txt", "qtd": 0, "vazio": True}

    finally:
        try:
            driver.quit()
        except Exception:
            pass