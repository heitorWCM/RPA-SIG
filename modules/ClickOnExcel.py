import pyautogui
import time
import modules.LocateImageOnScreen as LocateImageOnScreen
import modules.WaitWhileImageExists as wwie
import pygetwindow as gw
import modules.WaitOnWindow as wow
import sys

# Função principal para clicar e salvar o excel
def ClickOnExcel(path, fileName, prName):
    
    locationWithOutResults = LocateImageOnScreen.locate_image_on_screen(["./modules/ClickOnExcel-IMG/15-WithOutResults.png","./modules/ClickOnExcel-IMG/16-WithOutResults.png"], waitFind=1, lookForPresence=True, max_attempts=3)

    if not locationWithOutResults:    
        if prName.startswith("PRX"):
            SaveExcelPRX(path, fileName)
        else:
            SaveExcelStandard(path, fileName, prName)
        
        print("Excel saved successfully.")
    else:
        with open(path + "\\" + fileName+".txt", "a", encoding="utf-8") as f:
            agora = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time()))
            f.write(f"[{agora}] - Não encontrado nenhum dado\n")
        print("Report has no results, skipping excel export.")
        pyautogui.press('enter')

#Funcção para salvar o excel do PRX
def SaveExcelPRX(path, fileName):
    
    startTrying =time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time()))
    print("Saving excel for PRX")
    print(f"Running: {startTrying}")

    location = LocateImageOnScreen.locate_image_on_screen("./modules/ClickOnExcel-IMG/00-Export.png")
    pyautogui.moveTo(location.left+(location.width/2), location.top+(location.height/2),duration=0.5)
    time.sleep(1)

    # Clica na opção de exportar para excel
    location = LocateImageOnScreen.locate_image_on_screen("./modules/ClickOnExcel-IMG/01-Excel.png")
    actualPosisiton = pyautogui.position()
    pyautogui.moveTo(x=location.left+(location.width/2), y=actualPosisiton.y,duration=0.5)
    pyautogui.moveTo(x=location.left+(location.width/2), y=location.top+(location.height/2),duration=0.2)
    pyautogui.click()

    # Seleciona o local para salvar o arquivo
    location = LocateImageOnScreen.locate_image_on_screen("./modules/ClickOnExcel-IMG/02-SavePath.png", waitFind=1)
    pyautogui.click(x=location.left+location.width+20, y=location.top+(location.height/2))
    time.sleep(0.2)
    pyautogui.write(path, interval=0.1)
    time.sleep(0.3)
    pyautogui.press('enter')
    time.sleep(0.3)

    # Seleciona campo do nome do arquivo, insere o nome do arquivo e salva
    location = LocateImageOnScreen.locate_image_on_screen("./modules/ClickOnExcel-IMG/03-FileName.png")
    pyautogui.click(x=location.left+location.width+20, y=location.top+(location.height/2))
    time.sleep(0.5)
    pyautogui.write(fileName, interval=0.05)
    time.sleep(0.2)
    pyautogui.press('tab')
    pyautogui.press('tab')
    time.sleep(0.3)
    pyautogui.press('Enter')

    # Caso tenha que sobresecrever o arquivo existente
    location = LocateImageOnScreen.locate_image_on_screen(["./modules/ClickOnExcel-IMG/04-SaveAs.png","./modules/ClickOnExcel-IMG/05-SaveAs.png"], waitFind=3, lookForPresence=True, max_attempts=5)
    if location:
        pyautogui.press('enter')

    # Espera o arquivo ser salvo
    if LocateImageOnScreen.locate_image_on_screen(["./modules/ClickOnExcel-IMG/06-OpenFile.png","./modules/ClickOnExcel-IMG/07-OpenFile.png"], waitFind=3, lookForPresence=True, max_attempts=10):
        pyautogui.press('tab')
        time.sleep(0.3)
        pyautogui.press('enter')

#Função para salvar o excel padrão
def SaveExcelStandard(path, fileName, prName):
    tentativas = 0  # Move outside the while loop

    while True:
        try:
            startTrying =time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time()))
            print(f"\n{'='*6} Attempting to activate report window {'='*6}\n")
            print(f"{startTrying} - Attempt {tentativas + 1}")
            nome_da_janela = f"PROSYST [\\PROSYST\\WPROSYST\\{prName}.RPT ]"
            print(nome_da_janela)
            
            relatorio = wow.WaitOnWindow(nome_da_janela, timeout=120)

            relatorio.activate()
            time.sleep(1)
            relatorio.maximize()
            print("Janela do relatório ativada e maximizada.")
            break
        except Exception as e:
            print(f"Erro ao ativar a janela do relatório: {str(e)}. Tentando novamente... {tentativas}")
            tentativas += 1
            time.sleep(5)
            if tentativas >= 8:
                print("Número máximo de tentativas atingido. Abortando operação.")
                sys.exit()

    locationOption = [None, None]
    for i in range(2):

        relatorio.activate()

        time.sleep(2)

        location = LocateImageOnScreen.locate_image_on_screen("./modules/ClickOnExcel-IMG/08-Export.png")
        
        pyautogui.moveTo(location.left + location.width/2, location.top + location.height/2, duration=1)

        pyautogui.click(location.left + location.width/2, location.top + location.height/2)

        time.sleep(5)

        try:
            wow.WaitOnWindow("Export").activate()
        except:
            print("Export window not found...")
            sys.exit()

        time.sleep(1)

        while True:
            try:
                location = LocateImageOnScreen.locate_image_on_screen("./modules/ClickOnExcel-IMG/09-TypeBox.png", waitFind=1, max_attempts=1, lookForPresence=True)
                if location:
                    print("Selection of file type located.")
                    break
                else:
                    pyautogui.press('tab')
                    time.sleep(0.5)
            except Exception as e:
                print(f"Erro ao localizar a imagem: {str(e)}. Tentando novamente...")
                time.sleep(1)

        while True:
            try:
                if i == 0:
                    imageLocation = ["./modules/ClickOnExcel-IMG/10-PDFType.png","./modules/ClickOnExcel-IMG/11-PDFType.png"]
                else:
                    imageLocation = "./modules/ClickOnExcel-IMG/12-ExcelType.png"
                
                locationOption[i] = LocateImageOnScreen.locate_image_on_screen(imageLocation, waitFind=1, max_attempts=2, lookForPresence=True)
                
                if locationOption[i]:
                    print("File type selected.")
                    break
                else:
                    if i == 0:
                        pyautogui.press('up')
                        time.sleep(1)
                    else:
                        pyautogui.press('down')
                        time.sleep(1)

            except Exception as e:
                print(f"Erro ao localizar a imagem: {str(e)}. Tentando novamente...")
                time.sleep(1)

        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.press('Enter')
        time.sleep(5)


        nome_da_saida = wow.WaitOnWindow(["Export Options","Excel Format Options"])
        
        if nome_da_saida:
            pyautogui.press('enter')
            time.sleep(2)
        else:
            print("Export Options window not found, retrying...")
            sys.exit()

        location = LocateImageOnScreen.locate_image_on_screen(["./modules/ClickOnExcel-IMG/13-SaveAs.png","./modules/ClickOnExcel-IMG/14-SaveAs.png"], waitFind=5)

        pyautogui.click(x=location.left + 150, y=location.top + 50)
        pyautogui.write(path,interval=0.05)
        time.sleep(0.2)
        pyautogui.press('enter')
        time.sleep(0.5)

        pyautogui.click(x=location.left + 125, y=location.top + 395)
        pyautogui.write(fileName,interval=0.05)
        time.sleep(0.2)
        pyautogui.press('enter')
        time.sleep(1)

        if  wow.WaitOnWindow("File already exists", wait=1, timeout=10):
            pyautogui.press('enter')

    time.sleep(10)

    relatorio.activate()
    time.sleep(1)

    print("Closing report window.")
    relatorio.close()

    time.sleep(2)