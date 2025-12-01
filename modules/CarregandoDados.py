import time
import os
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


def CarregandoDados():
    
    time.sleep(0.5)

    loading_images = get_loading_images()

    location = locate_image_on_screen(loading_images, waitFind=1, max_attempts=30, lookForPresence=True)

    if location:
        wait_while_image_exists(loading_images, timeout=900)
        time.sleep(2)
    else:
        print("No loading screen detected, proceeding...")

    return