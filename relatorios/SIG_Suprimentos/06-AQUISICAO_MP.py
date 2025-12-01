class ParametrosDados:
    def __init__(self, nomePR="PRX004317", nomeArquivo="06 - Aquisição de MP"):
        self.nomePR = nomePR
        self.nomeArquivo = nomeArquivo

### Cabeçalho padrão de todos os scripts de RPA SIG ###
import pyautogui
import time
import pygetwindow as gw
import sys
from pathlib import Path
import argparse
from datetime import datetime
from dateutil.relativedelta import relativedelta

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

SelecionaLayout(p.nomePR, nomeJanela, "SIG")

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Vai para o filtro de dados
location = locate_image_on_screen("./base/PRX004317-Filter.png", waitFind=2)
time.sleep(2.5)
pyautogui.moveTo(location.left+130, location.top+10,duration=0.3)
pyautogui.click()

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

dataInicioAquisicao = (datetime.strptime(datesFilter.initial_date,"%d%m%Y") + relativedelta(years=-1)).replace(day=1).strftime("%d%m%Y")

# Preenche os campos de data
pyautogui.write(datesFilter.initial_date,interval=0.05)
pyautogui.press('enter')
time.sleep(0.1)
pyautogui.write(datesFilter.final_date,interval=0.05)
pyautogui.press('enter')
time.sleep(0.1)
pyautogui.write(dataInicioAquisicao,interval=0.05)
pyautogui.press('enter')
time.sleep(0.1)
pyautogui.write(datesFilter.final_date,interval=0.05)
pyautogui.press('enter')
time.sleep(0.2)


# Preenche campo dos tipos de materiais
pyautogui.click(x=location.left+125, y=location.top+160)
time.sleep(0.2)
pyautogui.write("14101")
pyautogui.press('tab')
time.sleep(0.3)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Clica no botão de pesquisar
pyautogui.click(x=location.left+300, y=location.top+540)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Carrega os dados
CarregandoDados()

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

MouseBusy()

# Procura local para exportar
location = locate_image_on_screen("./base/PRX004317-RightClickExport.png", waitFind=5, max_attempts=10)
pyautogui.click(x=location.left+location.width+20, y=location.top+(location.height/2), button='Right')
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