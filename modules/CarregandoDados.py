import pyautogui
import time
from modules.WaitWhileImageExists import wait_while_image_exists

def CarregandoDados():
        
    # Carrega os dados
    wait_while_image_exists(["./modules/CarregandoDados-IMG/00-Carregando.png","./modules/CarregandoDados-IMG/01-Carregando.png"], timeout=900)
    time.sleep(2)

    return