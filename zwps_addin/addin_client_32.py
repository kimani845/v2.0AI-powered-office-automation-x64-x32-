# addin_client_32.py

import os
import sys
import winreg
import pythoncom
import win32com.server.register
from addin_common import WPSAddin, log_message, register_wps_addin_entry

# Set the 32-bit specific CLSID
class WPSAddin32(WPSAddin):
    # This CLSID must be a different value from the 64-bit version
    _reg_clsid_ = "{cf0b4f12-56e5-4818-b400-b3f2660e0a3c}" 

    _reg_class_spec_ = "addin_client_32.WPSAddin32"

def is_pyinstaller_bundle():
    """Check if running as PyInstaller bundle"""
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')


def register_com_server_pyinstaller(cls):
    """Register COM server for PyInstaller executable"""
    import winreg
        
    clsid = cls._reg_clsid_
    progid = cls._reg_progid_
    desc = cls._reg_desc_
        
        # Get the pyinstaller .exe executable path
    exe_path = sys.executable 
        
    log_message(f"Registering PyInstaller COM server with executable: {exe_path}")
        
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
            log_message("Removed stale InprocServer32 key for PyInstaller mode")
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
            
        log_message("PyInstaller COM server registered successfully")
        return True
    except Exception as e:
        log_message(f"Failed to register PyInstaller COM server: {e}")
        return False

def register_com_server_python(cls):
    """Register COM server for regular Python execution (InprocServer32 only)"""
    # This logic from your original code is already good for 32-bit environments
    # import winreg, pythoncom, os
    clsid = cls._reg_clsid_
    desc = cls._reg_desc_
    pythoncom_dll = pythoncom.__file__ 
    try:
        win32com.server.register.UseCommandLine(cls)
        with winreg.CreateKeyEx(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\InprocServer32") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, pythoncom_dll)
            winreg.SetValueEx(key, "ThreadingModel", 0, winreg.REG_SZ, "both")
            
        # Delete LocalServer32
        try:
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32")
            log_message("Removed stale LocalServer32 key for Python mode")
        except FileNotFoundError:
            pass
        log_message("Python COM server registered successfully (InprocServer32 only)")
        return True
    except Exception as e:
        log_message(f"Failed to register Python COM server: {e}")
        return False

def register_server(cls):
    """Register COM server - handles both Python and PyInstaller"""
    log_message("Starting 32-bit registration process...")
    if is_pyinstaller_bundle():
        # You'll need a 32-bit specific PyInstaller registration function here if you need one, but the existing one is often fine.
        com_success = register_com_server_pyinstaller(cls) # Assumes this is in common or 32-bit module
    else:
        com_success = register_com_server_python(cls)
    
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

# def unregister_server(cls):
#     # This can also be a common function, or you can have separate versions
#     # for cleaner separation. The current unregister logic is mostly platform-agnostic.
#     # ... (Keep the unregistration logic here)
#     pass

def run_com_server():
    """Run as COM server when called by Windows"""
    log_message("Starting COM server...")
    try:
        import win32com.server.localserver
        win32com.server.localserver.serve([WPSAddin32._reg_clsid_])
    except Exception as e:
        log_message(f"Error running COM server: {e}")

if __name__ == '__main__':
    log_message("Executing 32-bit main block.")
    if len(sys.argv) > 1:
        if sys.argv[1].lower() == '/regserver':
            register_server(WPSAddin32)
        elif sys.argv[1].lower() == '/unregserver':
            unregister_server(WPSAddin32)
        elif sys.argv[1].lower() == '/embedding':
            run_com_server()
    else:
        print("WPS Office AI Assistant Add-in (32-bit)")
        print("Usage:")
        print("  /regserver   - Register the add-in")
        print("  /unregserver - Unregister the add-in")
        print(f"Running as: {'PyInstaller bundle' if is_pyinstaller_bundle() else 'Python script'}")
        input("\nPress Enter to exit...")