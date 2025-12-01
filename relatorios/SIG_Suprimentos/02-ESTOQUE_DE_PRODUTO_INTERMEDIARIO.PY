class ParametrosDados:
    def __init__(self, nomePR="PR07416", nomeArquivo="02 - Estoque de Produto Intermediario"):
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
    WaitOnWindow
)

total_steps = 5
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
posicao_x = (tamanhoDaTela.width - tamanhoDaJanela.width) / 2
posicao_y = 0

janela.moveTo(int(posicao_x), int(posicao_y))
proportion = (janela.height/662)

print(f"\n{'='*20} Preparo da janela {'='*20}\n")

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Vai para o filtro de dados
location = locate_image_on_screen("./base/PR07416-Filter.png", waitFind=2)
pyautogui.click(location.left + int(240*proportion), location.top + int(10*proportion))
time.sleep(0.6)
pyautogui.write("2", interval=0.1)
pyautogui.press("enter")
time.sleep(0.6)
pyautogui.write("1000", interval=0.1)
pyautogui.press("enter")
time.sleep(0.6)
pyautogui.write("2", interval=0.1)
pyautogui.press("enter")
time.sleep(0.6)
pyautogui.write("1000", interval=0.1)
pyautogui.press("enter")
time.sleep(0.6)
pyautogui.write("1", interval=0.1)
pyautogui.press("enter")
time.sleep(0.6)
print(f"Inserindo data inicial... {datesFilter.initial_date}")
dateSearch = datesFilter.final_date[:-4] + datesFilter.final_date[-2:]
pyautogui.write(dateSearch, interval=0.25)
pyautogui.press("enter")
time.sleep(1)
pyautogui.click(location.left + int(470*proportion), location.top + int(515*proportion))

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

CarregandoDados()

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