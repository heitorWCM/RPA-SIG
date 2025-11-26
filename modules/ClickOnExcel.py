import pyautogui
import time
import modules.LocateImageOnScreen as LocateImageOnScreen
import modules.WaitWhileImageExists as wwie

# Função principal para clicar e salvar o excel
def ClickOnExcel(path, fileName, PRName):
    if PRName.startswith("PRX"):
        SaveExcelPRX(path, fileName)
    else:
        SaveExcelStandard(path, fileName)
    
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
    pyautogui.write(path)
    time.sleep(0.3)
    pyautogui.press('enter')
    time.sleep(0.3)

    # Seleciona campo do nome do arquivo, insere o nome do arquivo e salva
    location = LocateImageOnScreen.locate_image_on_screen("./modules/ClickOnExcel-IMG/03-FileName.png")
    pyautogui.click(x=location.left+location.width+20, y=location.top+(location.height/2))
    time.sleep(0.5)
    pyautogui.write(fileName)
    pyautogui.press('tab')
    pyautogui.press('tab')
    time.sleep(0.3)
    pyautogui.press('Enter')

    # Caso tenha que sobresecrever o arquivo existente
    location = LocateImageOnScreen.locate_image_on_screen("./modules/ClickOnExcel-IMG/04-SaveAs.png", waitFind=3, lookForPresence=True)
    if location:
        pyautogui.press('enter')

    # Espera o arquivo ser salvo
    wwie.wait_while_image_exists("./modules/ClickOnExcel-IMG/05-OpenFile.png", timeout=300)
    pyautogui.press('tab')
    time.sleep(0.3)
    pyautogui.press('enter')

#Função para salvar o excel padrão
def SaveExcelStandard(path, fileName):
    print("Saving standard excel")
    print("Function not implemented yet.")