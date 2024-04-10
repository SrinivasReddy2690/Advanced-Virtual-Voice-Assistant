import math
import psutil
import time
from random import randint,seed
import subprocess
# import AppOpener
from pynput.keyboard import Key, Controller
from PIL import ImageGrab
import wmi
import os
import shutil
import pyautogui

class SystemTasks:
    def __init__(self):
        self.keyboard = Controller()

    def write(self, text):
        self.keyboard.type(text)

    def select(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('a')
        self.keyboard.release('a')
        self.keyboard.release(Key.ctrl)

    def hitEnter(self):
        self.keyboard.press(Key.enter)
        self.keyboard.release(Key.enter)

    def delete(self):
        self.select()
        self.keyboard.press(Key.backspace)
        self.keyboard.release(Key.backspace)

    def copy(self):
        self.select()
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('c')
        self.keyboard.release('c')
        self.keyboard.release(Key.ctrl)

    def paste(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('v')
        self.keyboard.release('v')
        self.keyboard.release(Key.ctrl)

    def pause(self):
        self.keyboard.press(Key.space)
        self.keyboard.release(Key.space)
    
    def resume(self):
        self.keyboard.press(Key.space)
        self.keyboard.release(Key.space)
    
    def new_file(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('n')
        self.keyboard.release('n')
        self.keyboard.release(Key.ctrl)

    def save(self, name):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('s')
        self.keyboard.release('s')
        self.keyboard.release(Key.ctrl)
        time.sleep(0.2)
        self.write(name)
        self.hitEnter()


class TabOpt:
    def __init__(self):
        self.keyboard = Controller()

    def switchTab(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press(Key.tab)
        self.keyboard.release(Key.tab)
        self.keyboard.release(Key.ctrl)

    def closeTab(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('w')
        self.keyboard.release('w')
        self.keyboard.release(Key.ctrl)

    def newTab(self):
        self.keyboard.press(Key.ctrl)
        self.keyboard.press('t')
        self.keyboard.release('t')
        self.keyboard.release(Key.ctrl)


class WindowOpt:
    def __init__(self):
        self.keyboard = Controller()

    def closeWindow(self):
        self.keyboard.press(Key.alt_l)
        self.keyboard.press(Key.f4)
        self.keyboard.release(Key.f4)
        self.keyboard.release(Key.alt_l)

    def minimizeWindow(self):
        for i in range(2):
            self.keyboard.press(Key.cmd)
            self.keyboard.press(Key.down)
            self.keyboard.release(Key.down)
            self.keyboard.release(Key.cmd)
            time.sleep(0.05)

    def maximizeWindow(self):
        self.keyboard.press(Key.cmd)
        self.keyboard.press(Key.up)
        self.keyboard.release(Key.up)
        self.keyboard.release(Key.cmd)

    def switchWindow(self):
        self.keyboard.press(Key.alt_l)
        self.keyboard.press(Key.tab)
        self.keyboard.release(Key.tab)
        self.keyboard.release(Key.alt_l)

def screenshot():
    time.sleep(1)  
    pyautogui.hotkey('win', 'fn', 'f10')
    time.sleep(1)
    screenshot = pyautogui.screenshot()
    saved_directory = r"C:\Users\banka\Documents\MAJOR PROJECT\Virtual-Voice-Assistant\Plugins\Screenshot_Assistant"
    if screenshot is not None:
        os.makedirs(saved_directory, exist_ok=True)
        timestamp = time.strftime("%Y%m%d-%H%M%S")
        screenshot_path = os.path.join(saved_directory, f"screenshot_{timestamp}.png")
        screenshot.save(screenshot_path)
        print(f"Screenshot saved: {screenshot_path}")
    else:
        print("No screenshot captured.")

def systemInfo():
    c = wmi.WMI()
    my_system_1 = c.Win32_LogicalDisk()[0]
    my_system_2 = c.Win32_ComputerSystem()[0]
    info = f"Total Disk Space: {round(int(my_system_1.Size)/(1024**3),2)} GB\n" \
           f"Free Disk Space: {round(int(my_system_1.Freespace)/(1024**3),2)} GB\n" \
           f"Manufacturer: {my_system_2.Manufacturer}\n" \
           f"Model: {my_system_2. Model}\n" \
           f"Owner: {my_system_2.PrimaryOwnerName}\n" \
           f"Number of Processors: {psutil.cpu_count()}\n" \
           f"System Type: {my_system_2.SystemType}"
    return info


def convert_size(size_bytes):
    if size_bytes == 0:
        return "0B"
    size_name = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)
    return "%s %s" % (s, size_name[i])


def system_stats():
    cpu_stats = str(psutil.cpu_percent())
    battery_percent = psutil.sensors_battery().percent
    memory_in_use = convert_size(psutil.virtual_memory().used)
    total_memory = convert_size(psutil.virtual_memory().total)
    stats = f"Currently {cpu_stats} percent of CPU, {memory_in_use} of RAM out of total {total_memory} is being used and " \
                f"battery level is at {battery_percent}%"
    return stats


import subprocess

def app_path(app):
    app_paths = {
        'access': r'C:\Program Files\Microsoft Office\root\Office16\ACCICONS.exe',
        'powerpoint': r'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.exe',
        'word': r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.exe',
        'excel': r'C:\Program Files\Microsoft Office\root\Office16\EXCEL.exe',
        'outlook': r'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.exe',
        'onenote': r'C:\Program Files\Microsoft Office\root\Office16\ONENOTE.exe',
        'publisher': r'C:\Program Files\Microsoft Office\root\Office16\MSPUB.exe',
        'sharepoint': r'C:\Program Files\Microsoft Office\root\Office16\GROOVE.exe',
        'infopath designer': r'C:\Program Files\Microsoft Office\root\Office16\INFOPATH.exe',
        'infopath filler': r'C:\Program Files\Microsoft Office\root\Office16\INFOPATH.exe',
        'notepad': 'notepad.exe' 
    }
    try:
        return app_paths[app]
    except KeyError:
        return None

def open_app(query):
    apps = ('access', 'powerpoint', 'word', 'excel', 'outlook', 'onenote', 'publisher', 'sharepoint', 'infopath designer',
            'infopath filler', 'notepad')
    for app in apps:
        if app in query:
            path = app_path(app)
            if path:
                subprocess.Popen(path)
                return True
            else:
                print(f"Path for {app} not found.")
                return False
    print("Application not found.")
    return False


class SystemTaskNotepad:
    def __init__(self, directory="."):
        self.directory = directory

    def writee(self, filename, content):
        with open(os.path.join(self.directory, filename), "w") as file:
            file.write(content)
        return filename

    def savee(self, filename):
        file_path = os.path.join(self.directory, filename)
        if os.path.exists(file_path):
            saved_directory = os.path.join(self.directory, "saved_notes")
            os.makedirs(saved_directory, exist_ok=True)
            shutil.move(file_path, os.path.join(saved_directory, filename))
            print(f"File saved: {filename}")
            return os.path.join(saved_directory, filename)
        else:
            print(f"File not found: {filename}")
            return None

def take_note(note):
    seed(1) 
    subprocess.Popen('notepad.exe')
    time.sleep(0.2)
    sys_task = SystemTaskNotepad()
    filename = sys_task.writee(f'note_{randint(1, 100)}.txt', note) 
    saved_file = sys_task.savee(filename)
    if saved_file:
        subprocess.Popen(['notepad.exe', saved_file])


 
