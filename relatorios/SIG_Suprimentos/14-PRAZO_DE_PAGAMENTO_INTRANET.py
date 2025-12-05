class ParametrosDados:
    def __init__(self, nomePR="PRAZO", nomeArquivo="14 - Prazo de Pagamento (Intranet)"):
        self.nomePR = nomePR
        self.nomeArquivo = nomeArquivo

from dotenv import load_dotenv
import os

# Load .env
load_dotenv()

# Read variables
INTRANET_USER = os.getenv("INTRANET_USER")
INTRANET_PASS = os.getenv("INTRANET_PASS")


### Cabeçalho padrão de todos os scripts de RPA SIG ###
import pyautogui
import time
import pygetwindow as gw
import sys
from pathlib import Path
import argparse
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import base64

class ErroPersonalizado(Exception):
    pass

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
    MouseBusy,
    ClipToExcel
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
    driver.get("https://intranet.cristalmaster.com.br/")

    wait = WebDriverWait(driver, 20)

    current_step += 1
    print(f"PROGRESS:{current_step}/{total_steps}")

    print("Página carregada.")

    userField = wait.until(EC.presence_of_element_located((By.NAME, "inputUsuario")))
    userPassword = wait.until(EC.presence_of_element_located((By.NAME, "inputSenha")))


    userField.send_keys(INTRANET_USER)
    userPassword.send_keys(INTRANET_PASS)

    btn_login = driver.find_element(By.ID, "btnForm")
    btn_login.click()


    logout = wait.until(EC.presence_of_element_located((By.XPATH, '//a[@href="/logout"]')))

    if logout:
        print("Usuário logado")

    driver.get("https://intranet.cristalmaster.com.br/compras/follow/importacao/relatorio")


    filtro = Select(driver.find_element(By.ID, "selectFiltro"))
    filtro.select_by_value("CC")

    checkbox = driver.find_element(By.ID, "checkboxResumida")

    # If it's checked, click to uncheck
    if checkbox.is_selected():
        checkbox.click()

    botao_pesquisa = driver.find_element(By.NAME, "submit")
    botao_pesquisa.click()

    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

    
    button = driver.find_element(By.CSS_SELECTOR,"button svg[data-icon='filter']")
    button.click()

    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".dtsb-title")))

    # 1) Select column index 36
    col_select = Select(wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select.dtsb-data"))))
    col_select.select_by_value("36")

    time.sleep(3)

    # 2) Select condition "between"
    cond_select = Select(wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select.dtsb-condition"))))
    cond_select.select_by_value("between")

    time.sleep(3)

    # 3) Wait for both date inputs
    date_inputs = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input.dtsb-value.dt-datetime")))

    window_size = driver.get_window_size()
    width = window_size['width']
    height = window_size['height']

    time.sleep(3)
    dataInicio = (datetime.strptime(datesFilter.initial_date,"%d%m%Y")).strftime("%d/%m/%Y")
    dataFinal = (datetime.strptime(datesFilter.final_date,"%d%m%Y")).strftime("%d/%m/%Y")
    # 4) Fill both dates
    date_inputs[0].clear()
    date_inputs[0].send_keys(dataInicio)

    date_inputs[1].clear()
    date_inputs[1].send_keys(dataFinal)

    pyautogui.press('enter')

    # 1. Find the breadcrumb element
    elem = driver.find_element(By.CSS_SELECTOR, "li.breadcrumb-item.active")

    # 2. Get Selenium coordinates (relative to viewport)
    loc = elem.location_once_scrolled_into_view
    size = elem.size

    # 3. Center of the element
    center_x = loc['x'] + 50
    center_y = loc['y'] + 120

    # 4. Get browser window position relative to screen (important!)
    window_pos = driver.get_window_position()  # returns {'x': ..., 'y': ...}

    # 5. Adjust Selenium viewport → screen absolute coordinates
    absolute_x = window_pos['x'] + center_x
    absolute_y = window_pos['y'] + center_y

    print("Clicking on screen at:", absolute_x, absolute_y)

    # 6. Perform PyAutoGUI click
    pyautogui.moveTo(absolute_x, absolute_y, duration=0.2)
    pyautogui.click()

    MouseBusy()

    time.sleep(5)

    copy_button = driver.find_element(By.CSS_SELECTOR, "button.buttons-copy")
    copy_button.click()

    wait.until(EC.visibility_of_element_located((By.ID, "datatables_buttons_info")))
    print("Apareceu")

    pyautogui.moveTo(x=width/2,y=height/2,duration=1)

    MouseBusy()

    time.sleep(5)
    #datatables_buttons_info

    ClipToExcel(datesFilter.path,p.nomeArquivo)

    # btn_cookie = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Aceitar')]")))

    # btn_cookie.click()

    # # Tenta encontrar e mudar para o iframe
    # # O iframe geralmente tem um ID ou name específico
    # iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
    # driver.switch_to.frame(iframe)

    # # Agora tenta encontrar o elemento dentro do iframe
    # elemINICIO = wait.until(EC.presence_of_element_located((By.ID, "DATAINI")))

    # elemINICIO.clear()

    # elemINICIO.send_keys(datesFilter.initial_date)

    # # Agora tenta encontrar o elemento dentro do iframe
    # elemFIM = wait.until(EC.presence_of_element_located((By.ID, "DATAFIM")))
    
    # elemFIM.clear()
    
    # elemFIM.send_keys(datesFilter.final_date)

    # time.sleep(1)

    # submit_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "botao")))
    # driver.execute_script("arguments[0].click();", submit_button)  
    # # submit_button.click()

    # wait = WebDriverWait(driver, 20)

    # nomePDF = "13 - Cotação - " + datesFilter.initial_date + " - " + datesFilter.final_date

    # print(f"Nome do PDF: {nomePDF}")

    # pdf_options = {
    #         "landscape": False,
    #         "displayHeaderFooter": False,
    #         "printBackground": True,
    #         "preferCSSPageSize": True,
    #     }

    # result = driver.execute_cdp_cmd("Page.printToPDF", pdf_options)

    # pdf_data = base64.b64decode(result['data'])
    # with open(f"{datesFilter.path}\\{nomePDF}.pdf", "wb") as fall:
    #     fall.write(pdf_data)

    # print(f"PDF salvo em: {datesFilter.path}\\{nomePDF}.pdf")

    # current_step += 1
    # print(f"PROGRESS:{current_step}/{total_steps}")

    # csv_link = wait.until(EC.presence_of_element_located(
    #     (By.XPATH, "//a[contains(@href, 'gerarCSVFechamentoMoedaNoPeriodo')]")
    # ))

    # print(f"Link encontrado: {csv_link.get_attribute('href')}")
    
    # # Para baixar o arquivo CSV:
    # # OPÇÃO A: Clicar no link (faz download automático)
    # csv_link.click()
    # print("Download iniciado!")

    # janelaSalvarComo = WaitOnWindow("Salvar como", timeout=30)
    # janelaSalvarComo.activate()
    # time.sleep(2)

    # location = locate_image_on_screen("./base/COTACAO-SalvarComo.png", waitFind=2)
    
    # pyautogui.moveTo(x=location.left+location.width+10,y=location.top+location.height/2,duration=0.3)
    # pyautogui.click()

    # pyautogui.write(datesFilter.path)
    # pyautogui.press('enter')
    # time.sleep(0.5)

    # location = locate_image_on_screen("./base/COTACAO-NomeArquivo.png", waitFind=2)
    
    # pyautogui.moveTo(x=location.left+location.width+10,y=location.top+location.height/2,duration=0.3)
    # pyautogui.click()

    # pyautogui.hotkey('ctrl', 'a')
    # time.sleep(0.5)
    # pyautogui.press('delete')

    # pyautogui.write(p.nomeArquivo, interval=0.05)


    # time.sleep(1)
    # pyautogui.press('enter')
    # time.sleep(3)
    
    # janela = gw.getWindowsWithTitle("Salvar Como")

    # if janela:
    #     pyautogui.press("left")
    #     time.sleep(0.2)
    #     pyautogui.press("enter")

    # time.sleep(2)

    # current_step += 1
    # print(f"PROGRESS:{current_step}/{total_steps}")
except ErroPersonalizado as error:
        print(f"Erro: {error}")
finally:
    # Fechando o navegador
    driver.quit()