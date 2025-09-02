# addin_client_64.py

import os
import sys
import winreg
import pythoncom
import win32com.server.localserver
import win32com.server.register
from addin_common import WPSAddin, log_message, register_wps_addin_entry

# Set the 64-bit specific CLSID
class WPSAddin64(WPSAddin):
    # This CLSID must be a different value from the 32-bit version
    _reg_clsid_ = "{9b3e157b-7b0f-4318-8d2b-65c36398b1e4}" # A new, unique CLSID

    _reg_class_spec_ = "addin_client_64.WPSAddin64"

def is_pyinstaller_bundle():
    """Check if running as PyInstaller bundle"""
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')

def register_com_server_pyinstaller(cls):
    """Register COM server for PyInstaller executable"""
    # This logic from your original code is already good for 64-bit environments
    import winreg
    clsid = cls._reg_clsid_
    progid = cls._reg_progid_
    desc = cls._reg_desc_
    exe_path = sys.executable 
    log_message(f"Registering PyInstaller COM server with executable: {exe_path}")
    try:
        with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
        with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, f'"{exe_path}" /embedding')
        try:
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\InprocServer32")
            log_message("Removed stale InprocServer32 key for PyInstaller mode")
        except FileNotFoundError:
            pass
        with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\ProgID") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, progid)
        with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, progid) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
            with winreg.CreateKeyEx(key, "CLSID") as clsid_key:
                winreg.SetValueEx(clsid_key, "", 0, winreg.REG_SZ, clsid)
        log_message("PyInstaller COM server registered successfully")
        return True
    except Exception as e:
        log_message(f"Failed to register PyInstaller COM server: {e}")
        return False

def register_server(cls):
    """Register COM server - handles both Python and PyInstaller"""
    log_message("Starting 64-bit registration process...")
    if is_pyinstaller_bundle():
        com_success = register_com_server_pyinstaller(cls)
    else:
        # 64-bit Python does not need InprocServer32 registration for COM servers
        # that are out-of-process (LocalServer32), which is the case for PyInstaller.
        # But if you're running it directly, it will use PythonCOM's LocalServer.
        # This part of the logic is a bit tricky, but for a 64-bit PyInstaller,
        # the LocalServer32 approach is the correct one.
        log_message("Python script mode not supported for this 64-bit client.")
        return False
    
    if not com_success:
        log_message("FATAL: COM server registration failed")
        print("FATAL: Failed to register the COM server. Please run as Administrator.")
        return False
    
    clsid = cls._reg_clsid_
    progid = cls._reg_progid_
    desc = cls._reg_desc_
    
    if register_wps_addin_entry(clsid, progid, desc):
        log_message("WPS add-in registration successful")
        print("SUCCESS: Add-in registered successfully!")
        print("Please restart WPS Office to see the add-in.")
        return True
    else:
        log_message("WPS add-in registration failed")
        print("FAILED: Could not register WPS Office add-in entry.")
        return False

def unregister_server(cls):
    """Enhanced unregistration"""
    import winreg
        
    # Unregister COM server
    if not is_pyinstaller_bundle():
        try:
            import win32com.server.register
            win32com.server.register.UnregisterServer(cls._reg_clsid_)
            log_message("Python COM server unregistered successfully")
        except Exception as e:
            log_message(f"Could not unregister Python COM server: {e}")

        # Remove WPS Office entries
    clsid = cls._reg_clsid_
    progid = cls._reg_progid_
        
    # Remove COM server entries
    com_paths_to_remove = [
        f"CLSID\\{clsid}\\LocalServer32",
        f"CLSID\\{clsid}\\InprocServer32", 
        f"CLSID\\{clsid}\\ProgID",
        f"CLSID\\{clsid}",
        f"{progid}\\CLSID",
            progid
    ]
        
    for path in com_paths_to_remove:
        try:
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, path)
            log_message(f"Removed COM entry: {path}")
        except FileNotFoundError:
            pass
        except Exception as e:
            log_message(f"Could not remove COM entry {path}: {e}")
        
        # Remove WPS Office entries
    wps_addin_paths = [
        f"Software\\Kingsoft\\Office\\Addins\\{progid}",
        f"Software\\WPS\\Office\\Addins\\{progid}",
        f"Software\\WPS Office\\Addins\\{progid}"
    ]
        
    for path in wps_addin_paths:
        try:
            winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER, path, 0, 0)
            log_message(f"Removed: HKCU\\{path}")
        except FileNotFoundError:
            pass
        except Exception as e:
            log_message(f"Could not remove {path}: {e}")

    print("Unregistration complete.")

def run_com_server():
    """Run as COM server when called by Windows"""
    log_message("Starting COM server...")
    try:
        import win32com.server.localserver
        win32com.server.localserver.serve([WPSAddin64._reg_clsid_])
    except Exception as e:
        log_message(f"Error running COM server: {e}")

if __name__ == '__main__':
    log_message("Executing 64-bit main block.")
    if len(sys.argv) > 1:
        if sys.argv[1].lower() == '/regserver':
            register_server(WPSAddin64)
        elif sys.argv[1].lower() == '/unregserver':
            unregister_server(WPSAddin64)
        elif sys.argv[1].lower() == '/embedding':
            run_com_server()
    else:
        print("WPS Office AI Assistant Add-in (64-bit)")
        print("Usage:")
        print("  /regserver   - Register the add-in")
        print("  /unregserver - Unregister the add-in")
        print(f"Running as: {'PyInstaller bundle' if is_pyinstaller_bundle() else 'Python script'}")
        input("\nPress Enter to exit...")