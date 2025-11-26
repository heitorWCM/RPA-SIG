import pyautogui
import time
import sys

def wait_while_image_exists(image_path, timeout=30):
    start = time.time()
    tempoAtual = time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(start))

    print(f"Start searching - {tempoAtual}")
    
    tryTurn=0

    while True:
        try:
            is_on_screen = pyautogui.locateOnScreen(image_path) #, confidence=confidence)
        except pyautogui.ImageNotFoundException:
            # Treat as NOT found
            is_on_screen = None

        # If image is gone → exit the loop normally
        if is_on_screen is None:
            print("Image disappeared — continuing program.")
            end = time.time()
            tempoAtual = time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(end))
            print(f"Found - {tempoAtual}")
            return
        else:
            print(f"Image still on screen {tryTurn} — {image_path} — waiting...")

        # Check timeout
        if time.time() - start >= timeout:
            print("Timeout reached — quitting script.")
            sys.exit()

        tryTurn+=1

        time.sleep(5)