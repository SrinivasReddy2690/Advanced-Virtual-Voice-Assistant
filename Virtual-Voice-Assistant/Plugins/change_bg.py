import os
import random
import ctypes

def change_bg(directory):

    image_files = [os.path.join(directory, file) for file in os.listdir(directory) if file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.jfif'))]
    
    if not image_files:
        print("No image files found in the specified directory.")
        return
    
    random_image = random.choice(image_files)
    SPI_SETDESKWALLPAPER = 20
    ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, random_image, 3)
    print(f"Background changed to {random_image} successfully.")
