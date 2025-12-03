class ParametrosDados:
    def __init__(self, nomePR="PRX034504", nomeArquivo="03 - Numero de OPs"):
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

p = ParametrosDados()

nomeJanela = AbrePR(p.nomePR)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

janela = WaitOnWindow(nomeJanela)

SelecionaLayout(p.nomePR, nomeJanela, "NUMEROOPS")

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Vai para o filtro de dados
location = locate_image_on_screen("./base/PRX034504-Filter.png", waitFind=2)
time.sleep(2.5)
pyautogui.moveTo(location.left+145, location.top+5,duration=1)
pyautogui.click()

dataInicioProducao = (datetime.strptime(datesFilter.initial_date,"%d%m%Y") + relativedelta(months=-1)).replace(day=1).strftime("%d%m%Y")
dataFimProducao = (datetime.strptime(datesFilter.initial_date,"%d%m%Y") + relativedelta(months=2)).replace(day=1).strftime("%d%m%Y")
dataInicioEntrega = (datetime.strptime(datesFilter.initial_date,"%d%m%Y") + relativedelta(years=-1)).replace(day=1).strftime("%d%m%Y")
dataFimEntrega = (datetime.strptime(datesFilter.initial_date,"%d%m%Y") + relativedelta(months=5)).replace(day=1).strftime("%d%m%Y")

# Preenche os campos de data
pyautogui.write(dataInicioProducao)
pyautogui.press('enter')
time.sleep(0.1)
pyautogui.write(dataFimProducao)
pyautogui.press('enter')
time.sleep(0.1)
pyautogui.write(dataInicioEntrega)
pyautogui.press('enter')
time.sleep(0.1)
pyautogui.write(dataFimEntrega)
pyautogui.press('enter')
time.sleep(0.1)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Clica no botão de pesquisar
pyautogui.click(x=location.left+80, y=location.top+240)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Carrega os dados
CarregandoDados()

pyautogui.moveTo(x=janela.width/2,y=janela.height/2,duration=0.3)
MouseBusy()

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Procura local para exportar
location = locate_image_on_screen("./base/PRX034504-Tab.png", waitFind=5, max_attempts=10)
pyautogui.click(x=location.left+(location.width/2), y=location.top+location.height/2)
time.sleep(2)


pyautogui.moveTo(x=janela.width/2,y=janela.height/2,duration=0.3)
MouseBusy()


# Procura local para exportar
location = locate_image_on_screen("./base/PRX034504-DateFilter-1.png", waitFind=5, max_attempts=10)
pyautogui.click(x=location.left+(location.width/2), y=location.top+location.height/2)
time.sleep(1)

MouseBusy()

janelaFiltro = WaitOnWindow("Construtor de Filtros")

locateFiler = locate_image_on_screen("./base/PRX034504-DateFilter-2.png")

pyautogui.moveTo(x=locateFiler.left+(locateFiler.width)+20, y=locateFiler.top+(locateFiler.height/2), duration=0.5)
pyautogui.click()
time.sleep(1)

dataInicioFiltro = [(datetime.strptime(datesFilter.initial_date,"%d%m%Y")).strftime("%d"),(datetime.strptime(datesFilter.initial_date,"%d%m%Y")).strftime("%m"),(datetime.strptime(datesFilter.initial_date,"%d%m%Y")).strftime("%y")]
dataFimFiltro = [(datetime.strptime(datesFilter.final_date,"%d%m%Y")).strftime("%d"),(datetime.strptime(datesFilter.final_date,"%d%m%Y")).strftime("%m"),(datetime.strptime(datesFilter.final_date,"%d%m%Y")).strftime("%y")]

for number in dataInicioFiltro:
    print(number)
    pyautogui.write(number)
    time.sleep(0.2)
    pyautogui.press('right')
    time.sleep(0.1)

time.sleep(0.5)
pyautogui.press('right')
time.sleep(0.2)
pyautogui.press('right')
time.sleep(0.2)

for number in dataFimFiltro:
    pyautogui.write(number)
    time.sleep(0.2)
    pyautogui.press('right')
    time.sleep(0.1)


location = locate_image_on_screen("./base/PRX034504-DateFilter-3.png", waitFind=5, max_attempts=10)
pyautogui.click(x=location.left+(location.width/2), y=location.top+location.height/2)
time.sleep(0.5)

# Procura local para exportar
location = locate_image_on_screen("./base/PRX034504-RightClickExport.png", waitFind=5, max_attempts=10)
pyautogui.click(x=location.left+location.width/2, y=location.top+location.height/2, button='Right')
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