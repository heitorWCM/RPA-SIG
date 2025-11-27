import pyautogui
import time
import pygetwindow as gw

import sys
from pathlib import Path
#import os

total_steps = 9
current_step = 0

# Get absolute path and go up to project root (01 - RPA SIG)
current_file = Path(__file__).resolve()  # Make it absolute!
project_root = current_file.parent.parent.parent  # Go up 3 levels: 02-MP.py -> SIG_Suprimentos -> relatorios -> 01 - RPA SIG

sys.path.insert(0, str(project_root))

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Now your imports
from modules import (
    locate_image_on_screen,
    AbrePR,
    LimpaPR,
    SelecionaLayout,
    DeterminaDataECaminho,
    ClickOnExcel,
    CarregandoDados
)

time.sleep(2)  # time to switch to the correct screen

nomePR = "PRX012016"

nomeJanela = AbrePR(nomePR)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

SelecionaLayout(nomePR, nomeJanela, "SIG")

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

# Preenche campo dos tipos de materiais
pyautogui.click(x=location.left+(location.width)+20, y=location.top+350)
time.sleep(0.2)
pyautogui.write("1/1")
pyautogui.press('tab')
time.sleep(0.3)
pyautogui.write("1/8")
pyautogui.press('enter')
time.sleep(0.5)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Clica no botão de pesquisar
pyautogui.click(x=location.left+255, y=location.top+455)
time.sleep(15)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Carrega os dados
CarregandoDados()

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Procura local para exportar
location = locate_image_on_screen("./base/PRX012016-RightClickExport.png", waitFind=5, max_attempts=10)
pyautogui.click(x=location.left+210, y=location.top+40, button='Right')
time.sleep(0.3)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Exporta para o excel
ClickOnExcel(datesFilter.path, "02 - Estoque de MP", nomePR)

current_step += 1
print(f"PROGRESS:{current_step}/{total_steps}")

# Fecha da janela
atualJanela = gw.getWindowsWithTitle(nomeJanela)
atualJanela[0].close()
time.sleep(1)

LimpaPR()

print(f"Processo finalizado do {nomePR} - Estoque de MP")