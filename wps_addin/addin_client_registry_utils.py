"""
Shared registry utilities for WPS Add-in registration.
Contains common registry operations used by both 32-bit and 64-bit implementations.
"""
# Completed
import winreg
from addin_base_client import log_message

def register_wps_addin_entry(clsid, progid, description):
    """
    Creates the specific registry entry that WPS Office looks for.
    """
    log_message(f"Attempting to create WPS-specific entry for ProgID: {progid}")
    
    # Try multiple possible WPS registry paths
    wps_addin_paths = [
        r"Software\Kingsoft\Office\Addins",
        r"Software\WPS\Office\Addins",
        r"Software\WPS Office\Addins"
    ]
    
    registration_succeeded = False
    
    for base_path in wps_addin_paths:
        try:
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, base_path, 0, winreg.KEY_CREATE_SUB_KEY) as parent_key:
                log_message(f"Successfully opened/created parent key at HKCU\\{base_path}")
                with winreg.CreateKeyEx(parent_key, progid) as key:
                    winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, description)
                    winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, "AI Assistant")
                    winreg.SetValueEx(key, "LoadBehavior", 0, winreg.REG_DWORD, 3)
                    winreg.SetValueEx(key, "CLSID", 0, winreg.REG_SZ, clsid)
                    winreg.SetValueEx(key, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
                    log_message(f"Successfully created WPS Add-in entry '{progid}' at {base_path}")
                    registration_succeeded = True
                    break  # Success, no need to try other paths
        except Exception as e:
            log_message(f"Could not register at {base_path}: {e}")
            continue
    
    return registration_succeeded

def unregister_wps_addin_entry(clsid, progid):
    """Remove WPS Office and COM registry entries"""
    
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