####################################################################
#
#   Módulo para abrir um PR no Prosyst ERP
#
####################################################################

import pygetwindow as gw
import time
import pyautogui
import sys
import modules.LocateImageOnScreen as LocateImageOnScreen

# Função para garantir que a janela principal do Prosyst ERP esteja ativa
def JanelaPrincipal():
    # Verifica a janela ativa
    janela = gw.getActiveWindow()

    if janela.title != "Prosyst ERP":
        print("A janela ativa não é o Prosyst ERP. Tentando ativar a janela correta...")

        try:
            janela = gw.getWindowsWithTitle("Prosyst ERP")[0]
            janela.activate()
                
            # Maximiza a janela se não estiver maximizada
            if not janela.isMaximized:
                janela.maximize()
                time.sleep(1)
            
            pyautogui.click(janela.width/2,janela.height/2)
            
            time.sleep(0.5)

            return janela
        except IndexError:
            print("Janela 'Prosyst ERP' não encontrada. Encerrando o script.")
            sys.exit()

# Função para limpar o campo do PR
def LimpaPR():

    JanelaPrincipal()

    location = LocateImageOnScreen.locate_image_on_screen("./modules/AbrePR-IMG/00-Clean.png", lookForPresence=True, waitFind=2)

    if location:
        pyautogui.click(x=location.left+(location.width/2), y=location.top+(location.height/2))
        pyautogui.moveTo(650, 300, duration=0.3)
        pyautogui.click()

# Função principal para abrir o PR
def AbrePR(nome_do_pr):

    # Validação do nome do PR
    if not nome_do_pr.startswith("PR"):
        print("Nome do PR inválido. Deve começar com 'PR'. Encerrando o script.")
        sys.exit()
    
    LimpaPR()
    
    # Localiza a barra de pesquisa do PR e insere o nome do PR
    location = LocateImageOnScreen.locate_image_on_screen("./modules/AbrePR-IMG/01-Search.png")

    pyautogui.click(x=location.left+20, y=location.top+5)
    pyautogui.write(nome_do_pr)
    pyautogui.press('enter')

    time.sleep(2)

    # Clica no PR encontrado

    pyautogui.move(location.left+10, location.top+20)
    pyautogui.doubleClick()

    # Espera até que a janela do PR seja aberta
    start = time.time()

    timeout = 30  # segundos

    while True:
        active = gw.getActiveWindow()

        if active and active.title != "Prosyst ERP":
            print("Window found!")
            return active.title
        
        if time.time() - start >= timeout:
            print("Timeout reached — quitting script.")
            sys.exit()

        time.sleep(0.5)
    
    

