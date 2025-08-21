import winreg
import sys

# The GUID (or CLSID) of your COM add-in from your addin_client.py
# Make sure this matches exactly, including the curly braces.
ADDIN_GUID = "{bdb1ed0a-14d7-414d-a68d-a2df20b5685a}"

def check_clsid_registration(guid, arch_name, key_flag):
    """
    Checks if the primary CLSID key for the COM add-in is registered
    in the specified registry view (32-bit or 64-bit).
    """
    print(f"\n--- Checking CLSID registration for GUID: {guid} ({arch_name} view) ---")
    
    # Construct the full path to the key in the registry
    # HKEY_CLASSES_ROOT is a merged view, but we'll try HKLM directly with flags
    # for explicit 32/64-bit checks.
    base_key = winreg.HKEY_LOCAL_MACHINE
    key_path = f"SOFTWARE\\Classes\\CLSID\\{guid}"
    
    try:
        # Try to open the key using the specified key_flag for architecture
        key_handle = winreg.OpenKey(base_key, key_path, 0, winreg.KEY_READ | key_flag)
        
        print(f"[SUCCESS] Found the main registry key at: HKEY_LOCAL_MACHINE\\{key_path} (Using {arch_name} view)")
        
        # Read the default value (which should be your _reg_desc_)
        try:
            default_value, _ = winreg.QueryValueEx(key_handle, "")
            print(f"  -> Description (_reg_desc_): '{default_value}'")
        except FileNotFoundError:
            print("  -> No default description found for the CLSID key.")

        # Check for the important InprocServer32 subkey
        try:
            inproc_handle = winreg.OpenKey(key_handle, "InprocServer32")
            print("[SUCCESS] Found the 'InprocServer32' subkey.")
            
            # Read the default value, which points to the COM server DLL
            server_dll, _ = winreg.QueryValueEx(inproc_handle, "")
            print(f"  -> Server DLL: '{server_dll}'")
            winreg.CloseKey(inproc_handle)
            
        except FileNotFoundError:
            print("[ERROR] The main CLSID key was found, but the critical 'InprocServer32' subkey is MISSING!")
            
        # Check for the ProgID subkey
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

if __name__ == "__main__":
    print("Starting COM Add-in Registry Check...")

    # Check 64-bit registry view (KEY_WOW64_64KEY is the default for 64-bit Python)
    print("\nAttempting to check 64-bit registry view:")
    found_64bit = check_clsid_registration(ADDIN_GUID, "64-bit", winreg.KEY_WOW64_64KEY)

    # Check 32-bit registry view (KEY_WOW64_32KEY for WoW64 redirector)
    print("\nAttempting to check 32-bit registry view (via WoW64 redirector):")
    found_32bit = check_clsid_registration(ADDIN_GUID, "32-bit", winreg.KEY_WOW64_32KEY)

    print("\n--- Summary ---")
    if found_64bit:
        print(f"✅ COM Add-in {ADDIN_GUID} IS registered in the 64-bit registry view.")
    else:
        print(f"❌ COM Add-in {ADDIN_GUID} IS NOT registered in the 64-bit registry view.")

    if found_32bit:
        print(f"✅ COM Add-in {ADDIN_GUID} IS registered in the 32-bit registry view.")
    else:
        print(f"❌ COM Add-in {ADDIN_GUID} IS NOT registered in the 32-bit registry view.")

    if not found_64bit and not found_32bit:
        print("\nACTION REQUIRED: The COM Add-in is not registered in either 64-bit or 32-bit views.")
        print("Please ensure your registration script was run correctly and with sufficient permissions (e.g., as Administrator).")
    elif found_64bit and not found_32bit:
        print("\nNOTE: Registered in 64-bit only. If you need it for 32-bit Office, you might need to run the 32-bit `python.exe` version of your registration script.")
    elif not found_64bit and found_32bit:
        print("\nNOTE: Registered in 32-bit only. If you need it for 64-bit Office, you might need to run the 64-bit `python.exe` version of your registration script.")
    else:
        print("\nGood news! The COM Add-in appears to be registered in both 64-bit and 32-bit registry views.")

    # Keep the window open to see the output
    input("\nPress Enter to exit.")
