class ParametrosDados:
    def __init__(self, nomePR="COTACAO", nomeArquivo="13 - Cotação Dollar"):
        self.nomePR = nomePR
        self.nomeArquivo = nomeArquivo

### Cabeçalho padrão de todos os scripts de RPA SIG ###
import pyautogui
import time
import pygetwindow as gw
import sys
from pathlib import Path
import argparse

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import base64



# Passa o argumento de data e caminho via linha de comando
parser = argparse.ArgumentParser()
parser.add_argument('--initial_date', type=str, required=False)
parser.add_argument('--final_date', type=str, required=False)
parser.add_argument('--path', type=str, required=False)
args = parser.parse_args()

if args.initial_date and args.final_date and args.path:
    class DateFilter:
        def __init__(self, initial, final, path):
            self.initial_date = initial
            self.final_date = final
            self.path = path
    
    datesFilter = DateFilter(args.initial_date, args.final_date, args.path)
else:
    print("Error: initial_date, final_date, and path arguments are required.")
    sys.exit()

# Get absolute path and go up to project root (01 - RPA SIG)
current_file = Path(__file__).resolve()  # Make it absolute!
project_root = current_file.parent.parent.parent  # Go up 3 levels: 02-MP.py -> SIG_Suprimentos -> relatorios -> 01 - RPA SIG

sys.path.insert(0, str(project_root))

# Now your imports
from modules import (
    locate_image_on_screen,
    AbrePR,
    LimpaPR,
    SelecionaLayout,
    ClickOnExcel,
    CarregandoDados,
    WaitOnWindow,
    MouseBusy
)

total_steps = 5
current_step = 1

print(f"PROGRESS:{current_step}/{total_steps}")

#### Início do processo do relatório de Estoque de MP ####

time.sleep(1)  # time to switch to the correct screen

p = ParametrosDados()

# Inicializando o driver
#service = Service(ChromeDriverManager().install())
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("prefs", {
    "download.prompt_for_download": True,  # 🔥 force save dialog
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
#driver = webdriver.Chrome(service=service)
driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

try:
    
    driver.maximize_window()

    # Acessando o Google
    driver.get("https://www.bcb.gov.br/estabilidadefinanceira/historicocotacoes")

    wait = WebDriverWait(driver, 20)

    current_step += 1
    print(f"PROGRESS:{current_step}/{total_steps}")

    print("Página carregada.")

    btn_cookie = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Aceitar')]")))

    btn_cookie.click()

    # Tenta encontrar e mudar para o iframe
    # O iframe geralmente tem um ID ou name específico
    iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
    driver.switch_to.frame(iframe)

    # Agora tenta encontrar o elemento dentro do iframe
    elemINICIO = wait.until(EC.presence_of_element_located((By.ID, "DATAINI")))

    elemINICIO.clear()

    elemINICIO.send_keys(datesFilter.initial_date)

    # Agora tenta encontrar o elemento dentro do iframe
    elemFIM = wait.until(EC.presence_of_element_located((By.ID, "DATAFIM")))
    
    elemFIM.clear()
    
    elemFIM.send_keys(datesFilter.final_date)

    time.sleep(1)

    submit_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "botao")))
    driver.execute_script("arguments[0].click();", submit_button)  
    # submit_button.click()

    wait = WebDriverWait(driver, 20)

    nomePDF = "13 - Cotação - " + datesFilter.initial_date + " - " + datesFilter.final_date

    print(f"Nome do PDF: {nomePDF}")

    pdf_options = {
            "landscape": False,
            "displayHeaderFooter": False,
            "printBackground": True,
            "preferCSSPageSize": True,
        }

    result = driver.execute_cdp_cmd("Page.printToPDF", pdf_options)

    pdf_data = base64.b64decode(result['data'])
    with open(f"{datesFilter.path}\\{nomePDF}.pdf", "wb") as fall:
        fall.write(pdf_data)

    print(f"PDF salvo em: {datesFilter.path}\\{nomePDF}.pdf")

    current_step += 1
    print(f"PROGRESS:{current_step}/{total_steps}")

    csv_link = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//a[contains(@href, 'gerarCSVFechamentoMoedaNoPeriodo')]")
    ))

    print(f"Link encontrado: {csv_link.get_attribute('href')}")
    
    # Para baixar o arquivo CSV:
    # OPÇÃO A: Clicar no link (faz download automático)
    csv_link.click()
    print("Download iniciado!")

    janelaSalvarComo = WaitOnWindow("Salvar como", timeout=30)
    janelaSalvarComo.activate()
    time.sleep(2)

    location = locate_image_on_screen("./base/COTACAO-SalvarComo.png", waitFind=2)
    
    pyautogui.moveTo(x=location.left+location.width+10,y=location.top+location.height/2,duration=0.3)
    pyautogui.click()

    pyautogui.write(datesFilter.path)
    pyautogui.press('enter')
    time.sleep(0.5)

    location = locate_image_on_screen("./base/COTACAO-NomeArquivo.png", waitFind=2)
    
    pyautogui.moveTo(x=location.left+location.width+10,y=location.top+location.height/2,duration=0.3)
    pyautogui.click()

    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.5)
    pyautogui.press('delete')

    pyautogui.write(p.nomeArquivo, interval=0.05)


    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(3)
    
    janela = gw.getWindowsWithTitle("Salvar Como")

    if janela:
        pyautogui.press("left")
        time.sleep(0.2)
        pyautogui.press("enter")

    time.sleep(2)

    current_step += 1
    print(f"PROGRESS:{current_step}/{total_steps}")

finally:
    # Fechando o navegador
    driver.quit()