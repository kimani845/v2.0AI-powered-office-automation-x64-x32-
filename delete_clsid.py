import os
import sys
import ctypes
import winreg
import getpass # For robust input in admin mode

# The CLSID of the OLD add-in entry you want to delete.
# Ensure this is the correct CLSID you intend to remove.
CLSID_TO_DELETE = "{bdb57ff2-79b9-4205-9447-f5fe85f37312}"

# --- Helper for Running as Administrator ---
def run_as_admin():
    """Re-launch the script with admin privileges if not already elevated."""
    if ctypes.windll.shell32.IsUserAnAdmin():
        return True

    print("[INFO] Elevation required. Prompting for Administrator access...")

    # Re-run the script with admin rights
    # Use sys.executable to ensure the correct Python (or bundled exe) is launched
    # If run from an exe, sys.executable is the exe itself.
    params = " ".join([f'"{arg}"' for arg in sys.argv])
    try:
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, params, None, 1)
    except Exception as e:
        print(f"[ERROR] Failed to elevate: {e}")
        return False

    # Exit the current, non-elevated process
    sys.exit(0)

# --- Recursive Registry Deletion ---
def _delete_key_recursive(hive, subkey_path, view_flags):
    """
    Recursively deletes a registry key and all its subkeys.
    This function is called by delete_clsid_entry for each specific path.
    """
    try:
        # Open the key with delete and enumerate subkeys access
        # The access flags need to be combined carefully. KEY_WOW64_32KEY / KEY_WOW64_64KEY are for redirection.
        # KEY_ALL_ACCESS includes KEY_ENUMERATE_SUB_KEYS, KEY_SET_VALUE, KEY_READ, KEY_DELETE
        # For DeleteKey, the parent key needs KEY_SET_VALUE or KEY_ALL_ACCESS to delete a subkey.
        key_access = winreg.KEY_ALL_ACCESS | view_flags
        key_handle = winreg.OpenKey(hive, subkey_path, 0, key_access)
        
        # Enumerate and recursively delete all subkeys first
        # Enumerate keys by index until an OSError (no more keys) occurs
        while True:
            try:
                # Use index 0 as enumeration order is not guaranteed after deletion,
                # and EnumKey gets the key name by index. If a key is deleted, index 0
                # will point to the next available key.
                subkey_name = winreg.EnumKey(key_handle, 0) 
                full_subkey_path = f"{subkey_path}\\{subkey_name}"
                print(f"[INFO] Deleting subkey: HIVE:{hive_to_string(hive)}\\{full_subkey_path}")
                # Recursive call to delete the subkey
                _delete_key_recursive(hive, full_subkey_path, view_flags) 
            except OSError as e:
                # Error code 259 (No more data) indicates no more subkeys
                if e.winerror == 259:
                    break
                else:
                    print(f"[ERROR] OSError enumerating subkeys under {subkey_path}: {e}")
                    break # Exit loop on other OSError
            except Exception as e:
                print(f"[ERROR] Error enumerating subkeys under {subkey_path}: {e}")
                break # Exit loop on unexpected error

        winreg.CloseKey(key_handle) # Close the key before attempting to delete it

        # Now delete the key itself (it should now be empty)
        winreg.DeleteKey(hive, subkey_path)
        print(f"[DELETED] Key: HIVE:{hive_to_string(hive)}\\{subkey_path} (View: {'32-bit' if view_flags == winreg.KEY_WOW64_32KEY else '64-bit'})")
        return True
    except FileNotFoundError:
        print(f"[INFO] Key not found (already deleted or never existed): HIVE:{hive_to_string(hive)}\\{subkey_path} (View: {'32-bit' if view_flags == winreg.KEY_WOW64_32KEY else '64-bit'})")
        return True # Considered successful if it's already gone
    except PermissionError:
        print(f"[ERROR] Permission denied to delete key: HIVE:{hive_to_string(hive)}\\{subkey_path}. Ensure script runs as Administrator.")
        return False
    except Exception as e:
        print(f"[ERROR] An unexpected error occurred deleting key HIVE:{hive_to_string(hive)}\\{subkey_path}: {e}")
        return False

def hive_to_string(hive):
    if hive == winreg.HKEY_CLASSES_ROOT: return "HKEY_CLASSES_ROOT"
    if hive == winreg.HKEY_CURRENT_USER: return "HKEY_CURRENT_USER"
    if hive == winreg.HKEY_LOCAL_MACHINE: return "HKEY_LOCAL_MACHINE"
    return str(hive)

def delete_clsid_entry(clsid_to_delete):
    """
    Attempts to delete a CLSID entry and its related entries in various registry locations.
    """
    print(f"\n--- Attempting to delete CLSID: {clsid_to_delete} and related entries ---")
    
    # Registry paths to target for CLSID deletion
    # We use HKEY_LOCAL_MACHINE as the base for SOFTWARE\Classes\CLSID
    # and rely on view_flags for redirection.
    clsid_paths = [
        # HKEY_LOCAL_MACHINE - 64-bit view (for 64-bit applications)
        (winreg.HKEY_LOCAL_MACHINE, f"SOFTWARE\\Classes\\CLSID\\{clsid_to_delete}", winreg.KEY_WOW64_64KEY, "64-bit HKLM"),
        # HKEY_LOCAL_MACHINE - 32-bit view (for 32-bit applications on 64-bit OS)
        (winreg.HKEY_LOCAL_MACHINE, f"SOFTWARE\\Classes\\WOW6432Node\\CLSID\\{clsid_to_delete}", winreg.KEY_WOW64_32KEY, "32-bit HKLM (WOW6432Node)"),
        # HKEY_CURRENT_USER (usually not subject to WoW64 redirection, uses default view)
        (winreg.HKEY_CURRENT_USER, f"Software\\Classes\\CLSID\\{clsid_to_delete}", 0, "HKCU") # 0 for default view_flags
    ]

    all_successful = True
    for hive, path, view_flags, description in clsid_paths:
        print(f"\nAttempting to delete from {description} view...")
        # For HKEY_CURRENT_USER, if view_flags is 0, we should ensure it's not passed to OpenKey with a WOW64 flag
        current_view_flags = view_flags if view_flags != 0 else winreg.KEY_READ | winreg.KEY_SET_VALUE
        if not _delete_key_recursive(hive, path, view_flags):
            all_successful = False
            print(f"[ERROR] Failed to delete CLSID from {description} view.")

    # Also check for WPS Office Add-ins entry for this CLSID (if it points to it)
    print("\n--- Checking and deleting associated WPS Office Add-in entries ---")
    wps_addins_base_path = "Software\\Kingsoft\\Office\\Addins"
    try:
        # Open HKEY_CURRENT_USER for WPS add-in paths with read and enumerate access
        # Default view flags (0) for HKCU is usually sufficient.
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, wps_addins_base_path, 0, winreg.KEY_READ | winreg.KEY_ENUMERATE_SUB_KEYS) as wps_key:
            wps_keys_to_delete = []
            i = 0
            while True:
                try:
                    addin_name = winreg.EnumKey(wps_key, i)
                    addin_full_path = f"{wps_addins_base_path}\\{addin_name}"
                    
                    # Open each specific add-in key to read its CLSID
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, addin_full_path, 0, winreg.KEY_READ) as addin_specific_key:
                        try:
                            clsid_value, _ = winreg.QueryValueEx(addin_specific_key, "CLSID")
                            if clsid_value.lower() == clsid_to_delete.lower():
                                wps_keys_to_delete.append(addin_full_path)
                        except FileNotFoundError:
                            pass # No CLSID value for this add-in
                        except Exception as e:
                            print(f"[WARNING] Error reading CLSID for WPS add-in entry {addin_name}: {e}")
                    i += 1
                except OSError as e: # End of enumeration
                    if e.winerror == 259: # No more data
                        break
                    else:
                        print(f"[ERROR] OSError enumerating WPS add-in entries: {e}")
                        break
                except Exception as e:
                    print(f"[ERROR] Error during WPS add-in enumeration: {e}")
                    break
        
        for key_path in wps_keys_to_delete:
            print(f"[INFO] Deleting WPS Office Add-in entry linked to old CLSID: HKEY_CURRENT_USER\\{key_path}")
            # Delete the WPS add-in entry. For HKCU, view_flags 0 is appropriate.
            if not _delete_key_recursive(winreg.HKEY_CURRENT_USER, key_path, 0):
                all_successful = False

    except FileNotFoundError:
        print("[INFO] WPS Office Add-ins base key not found. No WPS-specific entries to delete.")
    except PermissionError:
        print("[ERROR] Permission denied to access WPS Office Add-ins base key. Ensure script runs as Administrator.")
        all_successful = False
    except Exception as e:
        print(f"[ERROR] An unexpected error occurred while checking WPS Add-in entries: {e}")
        all_successful = False

    if all_successful:
        print(f"\n✅ All detected registry entries for CLSID {clsid_to_delete} have been deleted.")
    else:
        print(f"\n❌ Some registry entries for CLSID {clsid_to_delete} could not be fully deleted. Please review the errors.")

def main():
    if not run_as_admin():
        sys.exit(1)

    print(f"\nWARNING: This script will attempt to delete ALL registry entries associated with the CLSID: {CLSID_TO_DELETE}.")
    print("This action is irreversible and can impact applications that use this CLSID.")
    
    # Changed from getpass.getpass to input() for broader compatibility in various terminals
    confirm = input("Type 'DELETE' (case-sensitive) to confirm deletion: ") 
    
    if confirm == "DELETE":
        delete_clsid_entry(CLSID_TO_DELETE)
    else:
        print("\nDeletion cancelled by user.")

if __name__ == "__main__":
    main()
