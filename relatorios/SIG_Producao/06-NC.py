class ParametrosDados:
    def __init__(self, nomePR="PRX012016", nomeArquivo="06 - NC"):
        self.nomePR = nomePR
        self.nomeArquivo = nomeArquivo

### Cabeçalho padrão de todos os scripts de RPA SIG ###
import pyautogui
import time
import pygetwindow as gw
import sys
from pathlib import Path
import argparse

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

total_steps = 9
current_step = 1

print(f"PROGRESS:{current_step}/{total_steps}")

#### Início do processo do relatório de Estoque de MP ####

time.sleep(1)  # time to switch to the correct screen

p = ParametrosDados()

nomeJanela = AbrePR(p.nomePR)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

janela = WaitOnWindow(nomeJanela)

SelecionaLayout(p.nomePR, nomeJanela, "SIG_PROD")

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Vai para o filtro de dados
location = locate_image_on_screen("./base/PRX012016-Filter.png", waitFind=2)
time.sleep(2.5)
pyautogui.moveTo(location.left+(location.width)+15, location.top+location.height+10,duration=0.3)
pyautogui.click()

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Preenche os campos de data
pyautogui.write(datesFilter.initial_date)
pyautogui.press('tab')
pyautogui.write(datesFilter.final_date)
pyautogui.press('enter')
time.sleep(0.2)

# Local de estoque

pyautogui.click(x=location.left+125, y=location.top+275)
time.sleep(0.2)
pyautogui.write("5")
pyautogui.press('enter')
time.sleep(0.5)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Clica no botão de pesquisar
pyautogui.click(x=location.left+255, y=location.top+455)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Carrega os dados
CarregandoDados()

pyautogui.moveTo(janela.width/2,janela.width/2,duration=0.3)
MouseBusy()

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Procura local para exportar
location = locate_image_on_screen("./base/PRX012016-RightClickExport.png", waitFind=5, max_attempts=10)
pyautogui.click(x=location.left+210, y=location.top+40, button='Right')
time.sleep(0.3)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Exporta para o excel
ClickOnExcel(datesFilter.path, p.nomeArquivo, p.nomePR)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Fecha da janela
janela.close()
time.sleep(1)

LimpaPR()

print(f"Processo finalizado do {p.nomePR} - {p.nomeArquivo}")