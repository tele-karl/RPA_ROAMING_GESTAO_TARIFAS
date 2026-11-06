from pathlib import Path
import time
import pyautogui
import pyperclip

from pyjab.jabdriver import JABDriver

from src.services.log_service import registrar_log
from src.services.loader_export_service import construir_nome_loader

from src.configs.configuration import (
    DESTINO_DIR
)

def _focar_portal():
    """Traz a janela do portal para frente e garante o foco."""
    registrar_log("Iniciado", "Sucesso", "Trazendo o Portal do ICT para o primeiro plano.")
    try:
        w, h = pyautogui.size()
        pyautogui.click(w / 2, h / 2)  # foco no centro da tela
        time.sleep(2)
    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Falha ao focar portal: {e}")

def _encontrar_xml_atualizado() -> Path:
    """Retorna o .xml mais recente em DESTINO_DIR ou lança erro se não houver."""
    destino = Path(DESTINO_DIR)
    if not destino.exists():
        registrar_log("Iniciado", "Erro", f"Pasta não encontrada: {destino}")
        raise
    arquivos_xml = sorted(destino.glob("*.xml"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not arquivos_xml:
        registrar_log("Finalizado", "Erro", f"Nenhum .xml encontrado em {destino}")
        raise
    return arquivos_xml[0]

def start_load_loader() -> Path:
    """
    Fluxo UI:
      1) Focar portal
      2) Tools → Bulk Load → Start Load
      3) Add New → nome do loader (RATE_MMMYY)
      4) Job Parameters → Insert and Update → (...) → selecionar XML em data/arquivos_destino
      5) Confirmar/fechar
    Retorna o Path do XML carregado (para log/tracking).
    """
    xml_mais_recente = _encontrar_xml_atualizado()
    registrar_log("Iniciado", "Sucesso", f"Loader (XML) a enviar: {xml_mais_recente.name}")

    _focar_portal()

    try:
        registrar_log("Iniciado", "Sucesso", "Navegar em abas no portal.")
        # Clique em File para selecionar barra de Menu
        pyautogui.click(13, 35)
        time.sleep(1)

        # Navega para a direita até Tools (8 vezes)
        # File -> Edit -> View -> Agreement -> Network -> Billing -> Fin.Mgmt -> Admin -> Tools
        for _ in range(8):
            pyautogui.press("right")
            time.sleep(0.7)

        pyautogui.press("down"); time.sleep(0.7)   # abre Tools
        pyautogui.press("right"); time.sleep(0.7)  # deve estar em Bulk Load
        pyautogui.press("enter"); time.sleep(3)    # Start Load

        # Selecionar Filter (para aparecer a Lista)
        for _ in range(10):
            pyautogui.press("tab")
            time.sleep(0.7)
        pyautogui.press("enter"); time.sleep(3)

        # Clique no botão Add New
        pyautogui.click(1589, 761); time.sleep(2)

        # Nome do job (usa a mesma regra do export)
        nome_loader_atualizado = construir_nome_loader()
        pyautogui.typewrite(nome_loader_atualizado); time.sleep(0.7)

        # Pressiona Shift+Tab para seleção das abas
        pyautogui.keyDown("shift"); time.sleep(0.5)
        pyautogui.press("tab"); time.sleep(0.5)
        pyautogui.keyUp("shift"); time.sleep(0.5)
        # Ir para Job Parameters
        pyautogui.press("right"); time.sleep(0.7)

        # Seleciona "Insert and Update" (Importante!)
        pyautogui.click(294, 942); time.sleep(0.7)

        # Botão (...) – selecionar arquivo
        registrar_log("Iniciado", "Sucesso", "Abrindo diálogo para selecionar o XML.")
        pyautogui.click(906, 843); time.sleep(4)

        # Selecionar local onde está o loader atualizado (arquivos_destino)
        app_select_file = JABDriver(
            title="Select File",
            bridge_dll=r"C:\Program Files\Java\jre1.8.0_481\bin\WindowsAccessBridge-64.dll"
        )

        # Campo "Nome do arquivo:"
        campo_arquivo = next(
            e for e in app_select_file.find_elements()
            if (
                e.role == "text"
                and e.name == "File name:"
            )
        )

        caminho_origem = (
            rf"{DESTINO_DIR}\{nome_loader_atualizado}.xml"
        )

        campo_arquivo.send_text(
            caminho_origem
        )

        # Botão Salvar
        botao_select_file = next(
            e for e in app_select_file.find_elements()
            if (
                e.role == "push button"
                and e.name == "Select File"
            )
        )

        botao_select_file.click()

        # Selecionar OK para confirmar
        pyautogui.press("tab"); time.sleep(0.5)   
        pyautogui.press("enter"); time.sleep(15) # aumentar se necessário 

        # pyautogui.click(1844, 997) # clique cancel
        # time.sleep(2)

        # Clique Status Monitor
        pyautogui.click(1316, 762)
        time.sleep(5)
        
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
                registrar_log("Finalizado", "Erro", f"Falha no carregamento do loader. Status = Failed (!)")
                break

        time.sleep(3)
        # Fechar abas abertas
        for _ in range(3):
            pyautogui.click(235, 60)
            time.sleep(0.7)

        #Fechar Portal
        pyautogui.keyDown('alt'); time.sleep(0.3)
        pyautogui.press('f4'); time.sleep(0.3)
        pyautogui.keyUp('alt')
        time.sleep(1)
        pyautogui.press('enter')  

        registrar_log("Finalizado", "Sucesso", f"Start Load iniciado para: {xml_mais_recente.name}")
        return xml_mais_recente

    except Exception as e:
        registrar_log("Finalizado", "Erro", f"Erro ao iniciar Start Load: {e}")
        raise