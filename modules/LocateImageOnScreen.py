import pyautogui
import time
import sys

def locate_image_on_screen(image_path, waitFind=0, lookForPresence=False, max_attempts=3):
    
    tryTurn = 0

    while tryTurn < max_attempts:
        try:
            location = pyautogui.locateOnScreen(image_path)
            
            if location:
                print("Found at:", location)
                center = pyautogui.center(location)
                print("Center point:", center)
                return location
                
        except pyautogui.ImageNotFoundException:
            print(f"Image not found - {image_path} (Attempt {tryTurn + 1}/{max_attempts})")
        
        tryTurn += 1
        
        if tryTurn < max_attempts:
            time.sleep(waitFind)
    
    # After max attempts reached
    print(f"Failed to find image after {max_attempts} attempts - {image_path}")
    if lookForPresence:
        return False
    else:
        sys.exit()