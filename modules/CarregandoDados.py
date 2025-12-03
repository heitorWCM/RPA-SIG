import time
import os
import pyautogui
from pathlib import Path
from modules.LocateImageOnScreen import locate_image_on_screen
from modules.WaitWhileImageExists import wait_while_image_exists

# Get the directory where this module is located
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(BASE_DIR, "CarregandoDados-IMG")

def get_loading_images():
    """
    Automatically get all PNG images from the CarregandoDados-IMG folder
    """
    images = []
    
    # Check if directory exists
    if not os.path.exists(IMG_DIR):
        print(f"Warning: Image directory not found: {IMG_DIR}")
        return images
    
    # Get all .png files in the directory, sorted
    for file in sorted(Path(IMG_DIR).glob("*.png")):
        images.append(str(file))
    
    if not images:
        print(f"Warning: No PNG images found in {IMG_DIR}")
    else:
        print(f"Loaded {len(images)} loading screen images")
    
    return images


def CarregandoDados(timeoutLoading=900):

    loading_images = get_loading_images()

    screen_w, screen_h = pyautogui.size()

    region_w = screen_w // 2     # 50% da largura
    region_h = screen_h // 2     # 50% da altura

    # Coordenadas de início (canto superior esquerdo) para centralizar
    region_x = (screen_w - region_w) // 2
    region_y = (screen_h - region_h) // 2

    central_region = (region_x, region_y, region_w, region_h)

    location = locate_image_on_screen(loading_images, waitFind=0.2, max_attempts=5, lookForPresence=True, regionArea=central_region)

    if location:
        wait_while_image_exists(loading_images, timeout=timeoutLoading)
        time.sleep(2)
    else:
        print("No loading screen detected, proceeding...")

    return