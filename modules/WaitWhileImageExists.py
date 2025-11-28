import pyautogui
import time
import sys

def wait_while_image_exists(image_path, timeout=30):
    # If user passes a single image → convert to list
    if isinstance(image_path, str):
        image_list = [image_path]
    else:
        image_list = image_path  # already a list

    start = time.time()
    tempoAtual = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(start))
    print(f"Start searching - {tempoAtual}")

    tryTurn = 0

    while True:
        any_image_found = False

        for img in image_list:
            try:
                found = pyautogui.locateOnScreen(img)
            except pyautogui.ImageNotFoundException:
                found = None

            if found is not None:
                print(f"[{tryTurn}] Image still on screen: {img}")
                any_image_found = True

        # None of the images were found → exit
        if not any_image_found:
            end = time.time()
            tempoAtual = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(end))
            print(f"Image(s) disappeared — continuing program. ({tempoAtual})")
            return

        # Timeout
        if time.time() - start >= timeout:
            print("Timeout reached — quitting script.")
            sys.exit()

        tryTurn += 1
        time.sleep(5)
