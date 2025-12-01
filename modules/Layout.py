####################################################################
#
#   Módulo para abrir um PR no Prosyst ERP
#
####################################################################

import modules.WaitOnWindow as wow
import time
import pyautogui
import modules.LocateImageOnScreen as LocateImageOnScreen


def SelecionaLayout(nome_do_pr, janela_pr, nome_layout):

    # Janela selecionada do PR
    try:
        wow.WaitOnWindow(janela_pr, timeout=60).activate()
    except:
        print(f"Janela do PR {janela_pr} não encontrada...")
        return

    # Localiza botão de layout e clica
    location = LocateImageOnScreen.locate_image_on_screen("./modules/Layout-IMG/00-LayoutButton.png", waitFind=5)
    pyautogui.click(x=location.left+(location.width/2), y=location.top+(location.height/2))

    # Localiza o padrão de layout e clica
    location = LocateImageOnScreen.locate_image_on_screen(f"./modules/Layout-IMG/{nome_do_pr}-{nome_layout}.png", waitFind=1)
    pyautogui.click(x=location.left+(location.width/2), y=location.top+(location.height/2))
    time.sleep(5)