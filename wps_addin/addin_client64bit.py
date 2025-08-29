"""
64-bit specific WPS Add-in implementation.
Handles COM server registration and execution for 64-bit WPS Office installations.
"""

import sys
import os
import winreg
import win32com.client
import win32com.server.localserver
import pythoncom
from addin_base_client import WPSAddinBase, log_message

class WPSAddin64(WPSAddinBase):
    """64-bit WPS Add-in implementation"""
    
    def __init__(self):
        super().__init__()
        log_message("64-bit WPS Add-in initialized")

def is_pyinstaller_bundle():
    """Check if running as PyInstaller bundle"""
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')

def register_com_server_python_64bit(cls):
    """Register COM server for 64-bit Python execution (InprocServer32 only)"""
    clsid = cls._reg_clsid_
    desc = cls._reg_desc_
    
    # Path to pythoncom DLL
    pythoncom_dll = pythoncom.__file__
    
    try:
        # Register COM via pywin32 as usual
        import win32com.server.register
        win32com.server.register.UseCommandLine(cls)
        
        # Overwrite to only keep InprocServer32
        with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\InprocServer32") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, pythoncom_dll)
            winreg.SetValueEx(key, "ThreadingModel", 0, winreg.REG_SZ, "both")
        
        # Ensure LocalServer32 is removed if it exists
        try:
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32")
            log_message("Removed stale LocalServer32 key for 64-bit Python mode")
        except FileNotFoundError:
            pass
        
        log_message("64-bit Python COM server registered successfully (InprocServer32 only)")
        return True
    except Exception as e:
        log_message(f"Failed to register 64-bit Python COM server: {e}")
        return False

def register_com_server_pyinstaller_64bit(cls):
    """Register COM server for 64-bit PyInstaller executable"""
    clsid = cls._reg_clsid_
    progid = cls._reg_progid_
    desc = cls._reg_desc_
    
    # Get the pyinstaller .exe executable path
    exe_path = sys.executable
    
    log_message(f"Registering 64-bit PyInstaller COM server with executable: {exe_path}")
    
    try:
        # Register main CLSID root
        with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
        
        # Register LocalServer32 (not InprocServer32 for exe)
        with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, f'"{exe_path}" /embedding')
            
        # Ensure InprocServer32 is removed if it exists
        try:
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\InprocServer32")
            log_message("Removed stale InprocServer32 key for 64-bit PyInstaller mode")
        except FileNotFoundError:
            pass
        
        # Register ProgID
        with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\ProgID") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, progid)
        
        # Register ProgID mapping
        with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, progid) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, desc)
            with winreg.CreateKeyEx(key, "CLSID") as clsid_key:
                winreg.SetValueEx(clsid_key, "", 0, winreg.REG_SZ, clsid)
        
        log_message("64-bit PyInstaller COM server registered successfully")
        return True
    except Exception as e:
        log_message(f"Failed to register 64-bit PyInstaller COM server: {e}")
        return False

def register_server_64bit(cls):
    """Register 64-bit COM server - handles both Python and PyInstaller"""
    log_message("Starting 64-bit registration process...")
    
    # Step 1: Register COM server (different methods for Python vs PyInstaller)
    if is_pyinstaller_bundle():
        com_success = register_com_server_pyinstaller_64bit(cls)
    else:
        com_success = register_com_server_python_64bit(cls)
        
    if not com_success:
        log_message("FATAL: 64-bit COM server registration failed")
        print("FATAL: Failed to register the 64-bit COM server. Please run as Administrator.")
        return False
        
    # Step 2: Register WPS Office add-in entries
    from wps_addin.addin_client_registry_utils import register_wps_addin_entry
    
    clsid = cls._reg_clsid_
    progid = cls._reg_progid_
    desc = cls._reg_desc_
    
    if register_wps_addin_entry(clsid, progid, desc):
        log_message("64-bit WPS add-in registration successful")
        print("SUCCESS: 64-bit Add-in registered successfully!")
        print("Please restart WPS Office to see the add-in.")
        return True
    else:
        log_message("64-bit WPS add-in registration failed")
        print("FAILED: Could not register 64-bit WPS Office add-in entry.")
        return False

def unregister_server_64bit(cls):
    """Enhanced 64-bit unregistration"""
    # Unregister COM server
    if not is_pyinstaller_bundle():
        try:
            win32com.server.register.UnregisterServer(cls._reg_clsid_)
            log_message("64-bit Python COM server unregistered successfully")
        except Exception as e:
            log_message(f"Could not unregister 64-bit Python COM server: {e}")

    from wps_addin.addin_client_registry_utils import unregister_wps_addin_entry
    unregister_wps_addin_entry(cls._reg_clsid_, cls._reg_progid_)
    print("64-bit unregistration complete.")

def run_com_server_64bit():
    """Run as 64-bit COM server when called by Windows"""
    log_message("Starting 64-bit COM server...")
    try:
        win32com.server.localserver.serve([WPSAddin64._reg_clsid_])
    except Exception as e:
        log_message(f"Error running 64-bit COM server: {e}")

def check_environment_64bit():
    """Check and setup 64-bit environment"""
    if not is_pyinstaller_bundle():
        script_dir = os.path.dirname(os.path.abspath(__file__))
        if script_dir not in sys.path:
            sys.path.insert(0, script_dir)
            log_message(f"Added script directory to Python path: {script_dir}")
    
    log_message(f"Running 64-bit as: {'PyInstaller bundle' if is_pyinstaller_bundle() else 'Python script'}")
    log_message(f"Executable: {sys.executable}")