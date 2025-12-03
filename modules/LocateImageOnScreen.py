import pyautogui
import time
import sys

def locate_image_on_screen(image_path, waitFind=0, lookForPresence=False, max_attempts=3, regionArea=False):

    # Convert a single image into a list
    if isinstance(image_path, str):
        image_list = [image_path]
    else:
        image_list = image_path

    tryTurn = 0

    while tryTurn < max_attempts:
        
        messagePop = ''

        for img in image_list:
            try:
                if not regionArea:
                    location = pyautogui.locateOnScreen(img)
                else:
                    location = pyautogui.locateOnScreen(img, region=regionArea)

                if location:
                    print(f"Found: {img} at {location}")
                    center = pyautogui.center(location)
                    print("Center point:", center)
                    return location  # Return as soon as ANY image is found

            except pyautogui.ImageNotFoundException:
                messagePop = messagePop + f"Image not found - {img} (Attempt {tryTurn + 1}/{max_attempts})\n"

        tryTurn += 1

        print(messagePop)

        # Wait before next attempt
        if tryTurn < max_attempts:
            time.sleep(waitFind)

    # After all attempts and no image found
    print(f"Failed to find ANY image after {max_attempts} attempts: {image_list}")

    if lookForPresence:
        return False
    else:
        sys.exit()
