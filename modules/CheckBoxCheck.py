import modules.LocateImageOnScreen as LocateImageOnScreen
import time
import pyautogui

def CheckBoxCheck(prName, number):
    
    img_ok  = f"./modules/CheckBoxCheck-IMG/{prName}-CB{number}-OK.png"
    img_nok = f"./modules/CheckBoxCheck-IMG/{prName}-CB{number}-NOK.png"

    locationBox = LocateImageOnScreen.locate_image_on_screen(img_ok,waitFind=2,lookForPresence=True)

    if not locationBox:
        locationCB = LocateImageOnScreen.locate_image_on_screen(img_nok, waitFind=2)
        pyautogui.click(locationCB.left + -15, locationCB.top + (locationCB.height/2))

        print(f"CheckBox {number} marcado.")

        time.sleep(0.4)