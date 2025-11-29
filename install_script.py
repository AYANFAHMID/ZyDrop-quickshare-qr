import sys
import os

# --- Global Configuration ---
APP_NAME = "ZyDrop"
MENU_LABEL = "Share with ZyDrop (QR)"
ICON_WIN = "imageres.dll,-102"  # Native Windows share icon
SCRIPT_NAME = "script.pyw"  # Main application script (using .pyw to hide console)

def install_linux():
    """Creates a .desktop file on Linux for file manager integration."""
    print("\nInstalling for Linux...")
    
    home_dir = os.path.expanduser("~")
    apps_dir = os.path.join(home_dir, ".local", "share", "applications")
    
    if not os.path.exists(apps_dir):
        os.makedirs(apps_dir)
    
    script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), SCRIPT_NAME)
    if not os.path.exists(script_path):
        print(f"Error: Script not found at {script_path}")
        return

    # .desktop file content
    # %F allows passing the selected file path to the script
    desktop_entry = f"""[Desktop Entry]
Type=Application
Name={APP_NAME}
GenericName=LAN File Share
Comment=Transfer files via QR
Exec=python3 "{script_path}" %F
Icon=network-wireless
Terminal=false
Categories=Utility;Network;
NoDisplay=false
"""
    
    desktop_file = os.path.join(apps_dir, "zydrop.desktop")
    
    try:
        with open(desktop_file, "w") as f:
            f.write(desktop_entry)

        # Execution permissions
        os.chmod(script_path, 0o755)
        print(f"Installed at: {desktop_file}")
        print("It will now appear in 'Open with...' on right-click.")
    except Exception as e:
        print(f"Error during installation: {e}")

def install_windows():
    """Modifies the Windows Registry (Requires Admin)."""
    import ctypes
    import winreg
    print("\n🪟 Installing for Windows...")

    # --- Admin Check ---
    def is_admin():
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    if not is_admin():
        print("Requesting administrator permissions...")
        # Relaunch the script as admin
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        sys.exit()

    # --- Installation Logic ---
    script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), SCRIPT_NAME)
    
    # Use pythonw.exe to avoid the black console window when opening the app
    python_exe = sys.executable.replace("python.exe", "pythonw.exe")
    
    # Registry keys for Files (*) and Folders (Directory)
    keys_to_add = [
        (r"*\shell\ZyDrop", MENU_LABEL),
        (r"Directory\shell\ZyDrop", MENU_LABEL)
    ]

    print("--- Modifying Registry ---")
    try:
        for key_path, label in keys_to_add:
            # 1. Create the menu entry
            key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)
            winreg.SetValue(key, "", winreg.REG_SZ, label)
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, ICON_WIN)
            winreg.CloseKey(key)

            # 2. Assign the command
            cmd_key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, f"{key_path}\\command")
            cmd = f'"{python_exe}" "{script_path}" "%1"'
            winreg.SetValue(cmd_key, "", winreg.REG_SZ, cmd)
            winreg.CloseKey(cmd_key)
            print(f"✔ Registered: {key_path}")
            
        print("\nInstallation successful. Try right-clicking a file or folder!")
    except Exception as e:
        print(f"Critical error: {e}")
    
    input("Press Enter to exit...")

def uninstall_linux():
    """Removes the .desktop file from the user's applications directory."""
    print("\nUninstalling for Linux...")
    
    home_dir = os.path.expanduser("~")
    desktop_file = os.path.join(home_dir, ".local", "share", "applications", "zydrop.desktop")
    
    if not os.path.exists(desktop_file):
        print(f"No installation found at {desktop_file}. Nothing to do.")
        return

    try:
        os.remove(desktop_file)
        print(f"Uninstalled successfully. Removed: {desktop_file}")
        print("The 'Open with...' entry for ZyDrop has been removed.")
    except Exception as e:
        print(f"Error during uninstallation: {e}")

def uninstall_windows():
    """Removes the context menu entries from the Windows Registry (Requires Admin)."""
    import ctypes
    import winreg
    print("\n🪟 Uninstalling for Windows...")

    # --- Admin Check ---
    def is_admin():
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    if not is_admin():
        print("Requesting administrator permissions for uninstallation...")
        # Relaunch the script as admin
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        sys.exit()

    # --- Uninstallation Logic ---
    keys_to_remove = [
        r"*\shell\ZyDrop",
        r"Directory\shell\ZyDrop"
    ]

    print("--- Modifying Registry ---")
    all_successful = True
    for key_path in keys_to_remove:
        try:
            # Recursively delete the key (deletes subkeys like 'command' as well)
            winreg.DeleteKeyEx(winreg.HKEY_CLASSES_ROOT, key_path)
            print(f"✔ Removed: HKEY_CLASSES_ROOT\\{key_path}")
        except FileNotFoundError:
            print(f"Key not found (already uninstalled?): HKEY_CLASSES_ROOT\\{key_path}")
        except Exception as e:
            print(f"Failed to remove key {key_path}: {e}")
            all_successful = False
    
    if all_successful:
        print("\nUninstallation successful. The context menu entries have been removed.")
    else:
        print("\nUninstallation completed with errors.")
    
    input("Press Enter to exit...")

if __name__ == "__main__":
    print(f"--- {APP_NAME} Context Menu Manager ---")
    print("1. Install context menu entry")
    print("2. Uninstall context menu entry")
    print("------------------------------------")
    
    choice = input("Please select an option (1 or 2): ")

    if choice == '1':
        if sys.platform.startswith("win"): install_windows()
        elif sys.platform.startswith("linux"): install_linux()
        else: print(f"Unsupported OS. Please run {SCRIPT_NAME} manually.")
    elif choice == '2':
        if sys.platform.startswith("win"): uninstall_windows()
        elif sys.platform.startswith("linux"): uninstall_linux()
        else: print(f"Unsupported OS. No automatic uninstallation available.")
    else:
        print("Invalid option. Please run the script again and select 1 or 2.")