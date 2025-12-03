class ParametrosDados:
    def __init__(self, nomePR="PRX004317", nomeArquivo="11 - Fretes Conta Contabil 42443"):
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
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

time.sleep(2)  # Espera para garantir que o navegador abriu corretamente

# Acessando o Google
driver.get("https://www.ibge.gov.br/explica/inflacao.php")

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

time.sleep(5)  # Espera para garantir que a página carregou completamente

elem = driver.find_element(By.CLASS_NAME, "variavel-periodo")

print(f"Elemento encontrado: {elem.text}")

dataAtual = time.strftime("%Y%m%d", time.localtime(time.time()))

nomePDF = "12 - INFLAÇÃO IBGE - " + elem.text.replace("/", "-") + " - " + dataAtual

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

# Fechando o navegador
driver.quit()