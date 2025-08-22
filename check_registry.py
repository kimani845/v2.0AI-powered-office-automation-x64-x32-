import winreg
import sys

# The GUID (or CLSID) of your COM add-in
ADDIN_GUID = "{bdb1ed0a-14d7-414d-a68d-a2df20b5685a}"
# The ProgID/WPS Addin Entry Name (must match what's in addin_client.py)
WPS_ADDIN_ENTRY_NAME = "AIAddin.Connect"

def check_clsid_registration(guid, arch_name, key_flag):
    """
    Checks if the primary CLSID key for the COM add-in is registered
    in the specified registry view (32-bit or 64-bit).
    """
    print(f"\n--- Checking CLSID registration for GUID: {guid} ({arch_name} view) ---")
    
    base_key = winreg.HKEY_LOCAL_MACHINE
    key_path = f"SOFTWARE\\Classes\\CLSID\\{guid}"
    
    try:
        key_handle = winreg.OpenKey(base_key, key_path, 0, winreg.KEY_READ | key_flag)
        
        print(f"[SUCCESS] Found the main registry key at: HKEY_LOCAL_MACHINE\\{key_path} (Using {arch_name} view)")
        
        try:
            default_value, _ = winreg.QueryValueEx(key_handle, "")
            print(f"  -> Description (_reg_desc_): '{default_value}'")
        except FileNotFoundError:
            print("  -> No default description found for the CLSID key.")

        try:
            inproc_handle = winreg.OpenKey(key_handle, "InprocServer32")
            print("[SUCCESS] Found the 'InprocServer32' subkey.")
            
            server_dll, _ = winreg.QueryValueEx(inproc_handle, "")
            print(f"  -> Server DLL: '{server_dll}'")
            winreg.CloseKey(inproc_handle)
            
        except FileNotFoundError:
            print("[ERROR] The main CLSID key was found, but the critical 'InprocServer32' subkey is MISSING!")
            
        try:
            progid_handle = winreg.OpenKey(key_handle, "ProgID")
            print("[SUCCESS] Found the 'ProgID' subkey.")
            
            progid_value, _ = winreg.QueryValueEx(progid_handle, "")
            print(f"  -> Program ID (_reg_progid_): '{progid_value}'")
            winreg.CloseKey(progid_handle)
            
        except FileNotFoundError:
            print("[WARNING] The 'ProgID' subkey is missing. This might cause issues for applications relying on it.")
            
        winreg.CloseKey(key_handle)
        return True # Registration found

    except FileNotFoundError:
        print(f"\n[FAILURE] The registry key was NOT FOUND at: HKEY_LOCAL_MACHINE\\{key_path} (Using {arch_name} view)")
        print("  This means the registration command did not write this specific COM key.")
        return False # Registration not found

    except Exception as e:
        print(f"\n[ERROR] An unexpected error occurred while accessing the {arch_name} registry: {e}")
        return False # Error occurred

def check_wps_addin_entry_registration(wps_entry_name, expected_clsid):
    """
    Checks if the WPS Office add-in entry is registered under HKEY_CURRENT_USER.
    """
    print(f"\n--- Checking WPS Office Add-in entry for ProgID: {wps_entry_name} (HKEY_CURRENT_USER) ---")
    wps_addin_key_path = f"Software\\Kingsoft\\Office\\Addins\\{wps_entry_name}"

    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, wps_addin_key_path, 0, winreg.KEY_READ)
        print(f"[SUCCESS] Found WPS Office Add-in registry key at: HKEY_CURRENT_USER\\{wps_addin_key_path}")

        desc, _ = winreg.QueryValueEx(key, "Description")
        friendly_name, _ = winreg.QueryValueEx(key, "FriendlyName")
        load_behavior, _ = winreg.QueryValueEx(key, "LoadBehavior")
        clsid_value, _ = winreg.QueryValueEx(key, "CLSID")

        print(f"  -> Description: '{desc}'")
        print(f"  -> FriendlyName: '{friendly_name}'")
        print(f"  -> LoadBehavior: {load_behavior}")
        print(f"  -> CLSID: '{clsid_value}' (Expected: '{expected_clsid}')")
        
        if clsid_value.lower() == expected_clsid.lower():
            print("[SUCCESS] CLSID value matches the expected COM Add-in GUID.")
        else:
            print("[ERROR] CLSID value DOES NOT match the expected COM Add-in GUID!")

        winreg.CloseKey(key)
        return True
    except FileNotFoundError:
        print(f"[FAILURE] WPS Office Add-in registry key NOT FOUND at: HKEY_CURRENT_USER\\{wps_addin_key_path}")
        print("  This means the WPS-specific add-in entry was not created.")
        return False
    except Exception as e:
        print(f"[ERROR] An unexpected error occurred during WPS Office Add-in entry verification: {e}")
        return False

if __name__ == "__main__":
    print("Starting COM Add-in Registry Check...")

    found_clsid_64bit = check_clsid_registration(ADDIN_GUID, "64-bit", winreg.KEY_WOW64_64KEY)
    found_clsid_32bit = check_clsid_registration(ADDIN_GUID, "32-bit", winreg.KEY_WOW64_32KEY)
    found_wps_entry = check_wps_addin_entry_registration(WPS_ADDIN_ENTRY_NAME, ADDIN_GUID)

    print("\n--- Final Summary ---")
    com_ok = found_clsid_64bit and found_clsid_32bit
    wps_ok = found_wps_entry

    if com_ok:
        print(f"✅ COM Add-in {ADDIN_GUID} appears to be registered in both 64-bit and 32-bit system views.")
    else:
        print(f"❌ COM Add-in {ADDIN_GUID} registration is incomplete in system views. Review above errors.")

    if wps_ok:
        print(f"✅ WPS Office Add-in entry '{WPS_ADDIN_ENTRY_NAME}' is present in HKEY_CURRENT_USER.")
    else:
        print(f"❌ WPS Office Add-in entry '{WPS_ADDIN_ENTRY_NAME}' is MISSING in HKEY_CURRENT_USER.")

    if com_ok and wps_ok:
        print("\n🎉 All critical registry entries for your add-in (COM component and WPS entry) appear to be correctly set!")
        print("If the add-in is still not appearing, consider the troubleshooting steps below.")
    else:
        print("\n🚨 Important: There are still missing or incorrect registry entries. Please address the errors above.")
        print("Ensure you have run your PyInstaller-generated executables with '/regserver' as Administrator for both 32-bit and 64-bit versions.")

    input("\nPress Enter to exit.")
