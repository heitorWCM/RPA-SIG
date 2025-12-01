import time
from modules.LocateImageOnScreen import locate_image_on_screen
from modules.WaitWhileImageExists import wait_while_image_exists

def CarregandoDados():
    
    time.sleep(0.5)

    location = locate_image_on_screen(["./modules/CarregandoDados-IMG/00-Carregando.png","./modules/CarregandoDados-IMG/01-Carregando.png","./modules/CarregandoDados-IMG/02-Carregando.png","./modules/CarregandoDados-IMG/03-Carregando.png","./modules/CarregandoDados-IMG/04-Carregando.png"], waitFind=1, max_attempts=30, lookForPresence=True)

    if location:
        wait_while_image_exists(["./modules/CarregandoDados-IMG/00-Carregando.png","./modules/CarregandoDados-IMG/01-Carregando.png","./modules/CarregandoDados-IMG/02-Carregando.png","./modules/CarregandoDados-IMG/03-Carregando.png","./modules/CarregandoDados-IMG/04-Carregando.png"], timeout=900)
        time.sleep(2)
    else:
        print("No loading screen detected, proceeding...")

    return