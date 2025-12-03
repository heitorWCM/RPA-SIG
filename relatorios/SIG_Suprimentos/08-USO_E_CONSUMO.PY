class ParametrosDados:
    def __init__(self, nomePR="PR03017", nomeArquivo="08 - Uso e Consumo"):
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
    MouseBusy,
    CheckBoxCheck
)

total_steps = 9
current_step = 0

#### Início do processo do relatório de Estoque de MP ####

time.sleep(1)  # time to switch to the correct screen

p = ParametrosDados()

nomeJanela = AbrePR(p.nomePR)


print(f"\n{'='*20} Janela do PR aberta {p.nomePR} {'='*20}\n")
current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Obter dimensões
tamanhoDaTela = pyautogui.size()
janela = WaitOnWindow(nomeJanela)
tamanhoDaJanela = janela.size

# Centralizar horizontalmente, topo vertical
posicao_x = 0
posicao_y = 0

janela.moveTo(int(posicao_x), int(posicao_y))
proportion = (janela.height/764)

print(f"\n{'='*20} Preparo da janela {'='*20}\n")

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Vai para o filtro de dados
location = locate_image_on_screen("./base/PR03017-Filter.png", waitFind=2)
pyautogui.click(location.left + int(220*proportion), location.top + int(10*proportion))
time.sleep(0.6)
dateSearch = datesFilter.initial_date[:-4] + datesFilter.initial_date[-2:]
pyautogui.write(dateSearch, interval=0.25)
pyautogui.press("enter")
time.sleep(0.6)
dateSearch2 = datesFilter.final_date[:-4] + datesFilter.final_date[-2:]
pyautogui.write(dateSearch2, interval=0.25)
pyautogui.press("enter")
time.sleep(0.6)


CheckBoxCheck(p.nomePR, '01')

CheckBoxCheck(p.nomePR, '02')

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")


pyautogui.click(location.left + int(225*proportion), location.top + int(185*proportion))
time.sleep(0.6)

pyautogui.write("3", interval=0.25)
pyautogui.press("enter")
time.sleep(0.6)

pyautogui.write("0", interval=0.25)
pyautogui.press("enter")
time.sleep(0.6)

pyautogui.write("3", interval=0.25)
pyautogui.press("enter")
time.sleep(0.6)

pyautogui.write("9999", interval=0.25)
pyautogui.press("enter")
time.sleep(0.6)


CheckBoxCheck(p.nomePR, '03')

CheckBoxCheck(p.nomePR, '04')

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")


pyautogui.moveTo(x=location.left + int(820*proportion), y=location.top + int(435*proportion), duration=0.5)
time.sleep(0.6)
pyautogui.click()


################################################################
### NECESSÁRIO CONTINUAR DAQUI

CFOP = WaitOnWindow("CFOP")

time.sleep(3)

locateMarcarTodos = locate_image_on_screen("./base/PR03017-CFOP-Marcar.png", waitFind=3)
pyautogui.moveTo(x=locateMarcarTodos.left + (locateMarcarTodos.width/2), y=locateMarcarTodos.top + (locateMarcarTodos.height/2), duration=0.5)
pyautogui.click()

MouseBusy()
time.sleep(1)

CFOP.activate()

locateConfirmar = locate_image_on_screen("./base/PR03017-CFOP-Confirmar.png", waitFind=3)
pyautogui.moveTo(x=locateConfirmar.left + (locateConfirmar.width/2), y=locateConfirmar.top + (locateConfirmar.height/2), duration=0.5)
pyautogui.click()

MouseBusy()
time.sleep(1)

CFOP.close()

WaitOnWindow(CFOP.title, timeout=120, untilClosed=True)

MouseBusy()

time.sleep(5)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

pyautogui.moveTo(x=location.left + int(580*proportion), y=location.top + int(620*proportion), duration=0.5)
pyautogui.click()

MouseBusy()

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

CarregandoDados(timeoutLoading=5400)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

pyautogui.moveTo(janela.width/2,janela.width/2,duration=0.3)
MouseBusy()

# Exporta para o excel
ClickOnExcel(datesFilter.path, p.nomeArquivo, p.nomePR+"A")

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Fecha da janela
janela.close()
time.sleep(1)

LimpaPR()

print(f"Processo finalizado do {p.nomePR} - {p.nomeArquivo}")