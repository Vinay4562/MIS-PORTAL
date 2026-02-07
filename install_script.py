
import os
import sys
import shutil
import winshell
from win32com.client import Dispatch

def create_installer_script():
    # Define paths
    app_name = "MIS_PORTAL"
    exe_name = "MIS_PORTAL.exe"
    
    # Determine source path (handle frozen state)
    if getattr(sys, 'frozen', False):
        # If run as a bundled exe, look in the temp extraction folder
        base_path = sys._MEIPASS
    else:
        # If run as a script, look in the dist folder
        base_path = os.path.abspath("dist")
        
    exe_source = os.path.join(base_path, exe_name)
    
    # Destination in Local AppData
    install_dir = os.path.join(os.environ['LOCALAPPDATA'], app_name)
    exe_dest = os.path.join(install_dir, exe_name)
    
    print(f"Installing {app_name}...")
    
    # 1. Create Directory
    if not os.path.exists(install_dir):
        os.makedirs(install_dir)
        print(f"Created directory: {install_dir}")
        
    # 2. Copy Executable
    if os.path.exists(exe_source):
        shutil.copy2(exe_source, exe_dest)
        print(f"Copied executable to {exe_dest}")
    else:
        print(f"Error: Source executable not found at {exe_source}")
        return

    # 3. Create Desktop Shortcut
    desktop = winshell.desktop()
    path = os.path.join(desktop, f"{app_name}.lnk")
    
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.TargetPath = exe_dest
    shortcut.WorkingDirectory = install_dir
    shortcut.IconLocation = exe_dest
    shortcut.save()
    
    print(f"Shortcut created at {path}")
    print("Installation complete!")

if __name__ == "__main__":
    try:
        create_installer_script()
    except Exception as e:
        print(f"Installation failed: {e}")
        input("Press Enter to exit...")
