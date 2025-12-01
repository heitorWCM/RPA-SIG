import modules.LocateImageOnScreen as LocateImageOnScreen
import os
import time
import pyautogui


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(BASE_DIR, "CheckBoxCheck-IMG")

def CheckBoxCheck(prName, number):
    
    img_ok  = os.path.join(IMG_DIR, f"{prName}-CB_{number}_OK.png")
    img_nok = os.path.join(IMG_DIR, f"{prName}-CB_{number}_NOK.png")

    locationBox = LocateImageOnScreen.locate_image_on_screen(img_ok,waitFind=2,lookForPresence=True)

    if not locationBox:
        locationCB = LocateImageOnScreen.locate_image_on_screen(img_nok, waitFind=2)
        pyautogui.click(locationCB.left + -15, locationCB.top + (locationCB.height/2))

        print(f"CheckBox {number} marcado.")

        time.sleep(0.4)