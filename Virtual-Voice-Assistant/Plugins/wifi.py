import subprocess
import sys
import ctypes

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def turn_off_wifi():
    def _turn_wifi(state):
        subprocess.run(["Powershell", "-Command", f"netsh interface set interface Wi-Fi {state}"], shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    if is_admin():
        _turn_wifi("disabled")
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        _turn_wifi("disabled") 

def turn_on_wifi():
    def _turn_wifi(state):
        subprocess.run(["Powershell", "-Command", f"netsh interface set interface Wi-Fi {state}"], shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    if is_admin():
        _turn_wifi("enabled")
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        _turn_wifi("enabled") 

# turn_off_wifi()


