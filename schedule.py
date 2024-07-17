import os
import subprocess
import time
import pyautogui

def open_ppt_and_exit(filepath, slideshow_duration):
    try:
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"File not found at path: {filepath}")
        
        # Command to open PowerPoint
        command = f'start "" "{filepath}"'
        
        # Open PowerPoint
        subprocess.Popen(command, shell=True)
        
        # Wait for PowerPoint to open (adjust this delay if needed)
        time.sleep(5)
        
        # Send F5 key to start slideshow
        pyautogui.press('f5')
        
        # Wait for the specified duration before exiting slideshow
        time.sleep(slideshow_duration)
        
        # Close PowerPoint by sending Alt+F4 keys
        pyautogui.hotkey('alt', 'f4')
        
    except FileNotFoundError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"Error opening PowerPoint: {e}")

if __name__ == "__main__":
    ppt_path = r"C:\Users\chethan j s\Downloads\Chethan_Marketing.pptx"
    slideshow_duration_seconds = 60  # Change this to your desired duration in seconds
    open_ppt_and_exit(ppt_path, slideshow_duration_seconds)
