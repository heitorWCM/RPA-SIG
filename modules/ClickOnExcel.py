import pyautogui
import time
import modules.LocateImageOnScreen as LocateImageOnScreen
import modules.WaitWhileImageExists as wwie
import pygetwindow as gw
import sys

# Função principal para clicar e salvar o excel
def ClickOnExcel(path, fileName, prName):
    if prName.startswith("PRX"):
        SaveExcelPRX(path, fileName)
    else:
        SaveExcelStandard(path, fileName, prName)
    
    print("Excel saved successfully.")

#Funcção para salvar o excel do PRX
def SaveExcelPRX(path, fileName):
    
    print("Saving excel for PRX")

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
    pyautogui.write(fileName, interval=0.1)
    pyautogui.press('tab')
    pyautogui.press('tab')
    time.sleep(0.3)
    pyautogui.press('Enter')

    # Caso tenha que sobresecrever o arquivo existente
    location = LocateImageOnScreen.locate_image_on_screen(["./modules/ClickOnExcel-IMG/04-SaveAs.png","./modules/ClickOnExcel-IMG/05-SaveAs.png"], waitFind=3, lookForPresence=True)
    if location:
        pyautogui.press('enter')

    # Espera o arquivo ser salvo
    wwie.wait_while_image_exists(["./modules/ClickOnExcel-IMG/06-OpenFile.png","./modules/ClickOnExcel-IMG/07-OpenFile.png"], timeout=300)
    pyautogui.press('tab')
    time.sleep(0.3)
    pyautogui.press('enter')

#Função para salvar o excel padrão
def SaveExcelStandard(path, fileName, prName):

    while True:
        tentativas = 0

        try:
            relatorio =  gw.getWindowsWithTitle("PROSYST [\PROSYST\WPROSYST\{prName}.RPT ]")
            relatorio.activate()
            time.sleep(1)
            relatorio.maximize()
            print("Janela do relatório ativada e maximizada.")
            break
        except Exception as e:
            print(f"Erro ao ativar a janela do relatório: {str(e)}. Tentando novamente...")
            tentativas += 1
            time.sleep(2)
            if tentativas >= 5:
                print("Número máximo de tentativas atingido. Abortando operação.")
                sys.exit()

    for i in range(2):

        location = LocateImageOnScreen.locate_image_on_screen("./modules/ClickOnExcel-IMG/08-Export.png")

        pyautogui.click(location.left + location.width/2, location.top + location.height/2)

        time.sleep(5)

        gw.getActiveWindowTitle("Export")

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
                location = LocateImageOnScreen.locate_image_on_screen("./modules/ClickOnExcel-IMG/10-PDFType.png", waitFind=1, max_attempts=1, lookForPresence=True)
                if location:
                    print("Excel file type selected.")
                    break
                else:
                    if i == 0:
                        pyautogui.press('up')
                        time.sleep(0.5)
                    else:
                        pyautogui.press('down')
                        time.sleep(0.5)

            except Exception as e:
                print(f"Erro ao localizar a imagem: {str(e)}. Tentando novamente...")
                time.sleep(1)

        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.press('tab')
        time.sleep(0.5)
        pyautogui.press('Enter')
        time.sleep(2)

        gw.getActiveWindowTitle("Export Options")
        pyautogui.press('Enter')

        location = LocateImageOnScreen.locate_image_on_screen("./modules/ClickOnExcel-IMG/12-SaveAs.png", waitFind=5)

        pyautogui.click(x=location.left + 150, y=location.top + 50)
        pyautogui.write(path,interval=0.1)
        pyautogui.press('enter')
        time.sleep(0.5)

        pyautogui.click(x=location.left + 125, y=location.top + 395)
        pyautogui.write(fileName,interval=0.1)
        pyautogui.press('enter')
        time.sleep(5)

        if gw.getActiveWindow().title == "File already exists":
            pyautogui.press('enter')

    time.sleep(10)

    relatorio.close()

    time.sleep(2)