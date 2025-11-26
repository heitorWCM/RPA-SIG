import pyautogui
import time
import pygetwindow as gw

import sys
from pathlib import Path
import os

# Get absolute path and go up to project root (01 - RPA SIG)
current_file = Path(__file__).resolve()  # Make it absolute!
project_root = current_file.parent.parent.parent  # Go up 3 levels: 02-MP.py -> SIG_Suprimentos -> relatorios -> 01 - RPA SIG

sys.path.insert(0, str(project_root))

# DEBUG: Verify (you can remove these later)
print("Project root:", project_root)
print("Modules exists?", (project_root / "modules").exists())

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

time.sleep(5)  # time to switch to the correct screen

nomePR = "PRX012016"

nomeJanela = AbrePR(nomePR)

SelecionaLayout(nomePR, nomeJanela, "SIG_PROD")

# Vai para o filtro de dados
location = locate_image_on_screen("./BaseTelas/04-Filter.png", waitFind=2)
time.sleep(2.5)
pyautogui.moveTo(location.left+(location.width)+15, location.top+location.height+10,duration=0.3)
pyautogui.click()

# Cria parâmetros de data
datesFilter = DeterminaDataECaminho(r"C:\temp", "TesteRPA", start_day=3)
datesFilter.create_folder()

# Preenche os campos de data
pyautogui.write(datesFilter.initial_date)
pyautogui.press('tab')
pyautogui.write(datesFilter.final_date)
pyautogui.press('enter')
time.sleep(0.2)

# Preenche campo dos tipos de materiais
pyautogui.click(x=location.left+(location.width)+20, y=location.top+350)
time.sleep(0.2)
pyautogui.write("2/1000")
pyautogui.press('tab')
time.sleep(0.3)
pyautogui.write("2/1000")
time.sleep(0.5)

# Clica no botão de pesquisar
pyautogui.click(x=location.left+255, y=location.top+455)
time.sleep(15)

# Carrega os dados
CarregandoDados()

# Procura local para exportar
location = locate_image_on_screen("./BaseTelas/05-RightClickExport.png", waitFind=5, max_attempts=10)
pyautogui.click(x=location.left+210, y=location.top+40, button='Right')
time.sleep(0.3)


# Exporta para o excel
ClickOnExcel(datesFilter.path, "03 - Estoque de INTERMEDIARIO", nomePR)

# Fecha da janela
atualJanela = gw.getWindowsWithTitle(nomeJanela)
atualJanela[0].close()
time.sleep(1)

LimpaPR()

print(f"Processo finalizado do {nomePR} - Estoque de INTERMEDIARIO")