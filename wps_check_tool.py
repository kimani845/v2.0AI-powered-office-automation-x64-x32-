# """
# WPS Office Add-in Diagnostic Tool
# Run this to diagnose why your add-in isn't showing buttons
# """
# import winreg
# import os
# import sys
# import win32com.client
# import pythoncom

# def print_status(message, status="INFO"):
#     """Prints a formatted message with a status indicator."""
#     if status == "SUCCESS":
#         print(f"✅ {message}")
#     elif status == "FAILURE":
#         print(f"❌ {message}")
#     elif status == "WARNING":
#         print(f"⚠️ {message}")
#     else:
#         print(f"• {message}")

# def check_wps_addin_registry():
#     """Check all possible WPS add-in registry locations, including 32/64-bit views."""
#     print("\n=== WPS ADD-IN REGISTRY CHECK ===")
    
#     addin_name = "WPSAIAddin.Connect"
#     base_paths = [
#         r"Software\Kingsoft\Office\Addins",
#         r"Software\WPS\Office\Addins",
#         r"Software\WPS Office\Addins"
#     ]
    
#     found_any = False
    
#     # Check both 32-bit and 64-bit registry views
#     for view in [winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY]:
#         try:
#             with winreg.OpenKeyEx(winreg.HKEY_CURRENT_USER, r"Software", 0, winreg.KEY_READ | view) as root_key:
#                 for base_path in base_paths:
#                     full_path = f"{base_path}\\{addin_name}"
#                     try:
#                         with winreg.OpenKey(root_key, full_path) as key:
#                             print_status(f"Found entry: HKCU\\{full_path} (view: {view})", "SUCCESS")
#                             found_any = True
                            
#                             try:
#                                 load_behavior = winreg.QueryValueEx(key, "LoadBehavior")[0]
#                                 print_status(f"  LoadBehavior: {load_behavior}")
#                                 if load_behavior == 3:
#                                     print_status("  LoadBehavior is correct (3 = Load at startup)", "SUCCESS")
#                                 elif load_behavior == 2:
#                                     print_status("  LoadBehavior is 2 (Add-in failed to load)", "WARNING")
#                                 else:
#                                     print_status(f"  LoadBehavior is {load_behavior} (unexpected)", "WARNING")
#                             except FileNotFoundError:
#                                 print_status("  LoadBehavior not set", "FAILURE")
                            
#                             try:
#                                 winreg.QueryValueEx(key, "CLSID")[0]
#                                 print_status("  CLSID value exists", "SUCCESS")
#                             except FileNotFoundError:
#                                 print_status("  CLSID value not set", "FAILURE")
                            
#                     except FileNotFoundError:
#                         print_status(f"Entry not found: HKCU\\{full_path} (view: {view})", "INFO")
#                     except Exception as e:
#                         print_status(f"Error accessing HKCU\\{full_path}: {e}", "FAILURE")

#         except PermissionError:
#             print_status(f"Permission denied for registry view: {view}. Run as administrator.", "WARNING")
#         except FileNotFoundError:
#             pass # No base 'Software' key for this view, which is normal on some systems.
            
#     if not found_any:
#         print_status("No WPS add-in registry entries found!", "FAILURE")
#         return False
#     return True

# def check_com_registration():
#     """Check COM server registration for both in-process and local servers."""
#     print("\n=== COM SERVER REGISTRATION CHECK ===")
    
#     clsid = "{cf0b4f12-56e5-4818-b400-b3f2660e0a3c}"
#     progid = "WPSAIAddin.Connect"
    
#     # Check ProgID registration first
#     try:
#         with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid) as key:
#             print_status(f"ProgID registered: {progid}", "SUCCESS")
#             # Get the CLSID from the ProgID
#             clsid_from_progid = winreg.QueryValueEx(key, "CLSID")[0]
#             if clsid_from_progid == clsid:
#                 print_status("  ProgID's CLSID matches expected CLSID", "SUCCESS")
#             else:
#                 print_status(f"  ProgID's CLSID does NOT match expected CLSID ({clsid_from_progid} vs {clsid})", "FAILURE")
#     except FileNotFoundError:
#         print_status(f"ProgID not registered: {progid}", "FAILURE")
#         return

#     # Check CLSID registration for both LocalServer32 and InprocServer32
#     try:
#         with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}") as key:
#             print_status(f"CLSID registered: {clsid}", "SUCCESS")
            
#             found_server = False
#             # Check for LocalServer32 (executable)
#             try:
#                 with winreg.OpenKey(key, "LocalServer32") as server_key:
#                     exe_path = winreg.QueryValueEx(server_key, "")[0].strip('"')
#                     print_status(f"  LocalServer32 path: {exe_path}")
#                     if os.path.exists(exe_path):
#                         print_status("  Executable exists at this path", "SUCCESS")
#                         found_server = True
#                     else:
#                         print_status("  Executable NOT found at this path", "FAILURE")
#             except FileNotFoundError:
#                 print_status("  LocalServer32 key not found", "INFO")
            
#             # Check for InprocServer32 (DLL)
#             try:
#                 with winreg.OpenKey(key, "InprocServer32") as server_key:
#                     dll_path = winreg.QueryValueEx(server_key, "")[0].strip('"')
#                     print_status(f"  InprocServer32 path: {dll_path}")
#                     if os.path.exists(dll_path):
#                         print_status("  DLL exists at this path", "SUCCESS")
#                         found_server = True
#                     else:
#                         print_status("  DLL NOT found at this path", "FAILURE")
#             except FileNotFoundError:
#                 print_status("  InprocServer32 key not found", "INFO")
            
#             if not found_server:
#                 print_status("No server path (LocalServer32 or InprocServer32) found under CLSID!", "FAILURE")
                
#     except FileNotFoundError:
#         print_status(f"CLSID not registered: {clsid}", "FAILURE")

# def check_wps_installation():
#     """Check WPS Office installation, including 32/64-bit views."""
#     print("\n=== WPS OFFICE INSTALLATION CHECK ===")
#     install_paths = [
#         r"SOFTWARE\Kingsoft\Office",
#         r"SOFTWARE\WPS\Office",
#         r"SOFTWARE\WPS Office"
#     ]
#     wps_found = False
    
#     for view in [winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY]:
#         for path in install_paths:
#             try:
#                 with winreg.OpenKeyEx(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ | view) as key:
#                     print_status(f"WPS registry found: HKLM\\{path} (view: {view})", "SUCCESS")
#                     wps_found = True
#                     try:
#                         install_root = winreg.QueryValueEx(key, "InstallRoot")[0]
#                         print_status(f"  Install path: {install_root}")
#                         if os.path.exists(install_root):
#                             print_status("  Installation directory exists", "SUCCESS")
#                         else:
#                             print_status("  Installation directory NOT found", "FAILURE")
#                     except FileNotFoundError:
#                         print_status("  InstallRoot not found in registry", "WARNING")
#             except FileNotFoundError:
#                 pass
#             except PermissionError:
#                 print_status(f"Permission denied for registry path: HKLM\\{path} (view: {view}). Run as administrator.", "WARNING")
    
#     if not wps_found:
#         print_status("WPS Office installation not detected!", "FAILURE")
#         return False
    
#     return True

# def check_wps_com_objects():
#     """Test WPS COM object access by trying to connect to a running instance and then by creating a new one."""
#     print("\n=== WPS COM OBJECT TEST ===")
#     com_names = [
#         "kwps.Application",
#         "WPS.Application", 
#         "Kingsoft.WPS.Application"
#     ]
    
#     # Try to connect to a running instance
#     print("Attempting to connect to an active WPS application...")
#     for com_name in com_names:
#         try:
#             app = win32com.client.GetActiveObject(com_name)
#             if app:
#                 print_status(f"Connected to active {com_name} successfully", "SUCCESS")
#                 try:
#                     version = getattr(app, 'Version', 'Unknown')
#                     print_status(f"  Version: {version}")
#                 except:
#                     print_status("  Version: Could not determine", "WARNING")
#                 return True
#         except pythoncom.com_error as e:
#             # -2147221021 is COM_E_NOT_REGISTERED, -2147467259 is E_FAIL (no running instance)
#             if e.hresult not in [-2147221021, -2147467259]:
#                 print_status(f"Failed to connect to active {com_name}: {e}", "WARNING")
#             else:
#                 pass # This is expected if the app isn't running
#         except Exception as e:
#             print_status(f"Failed to connect to active {com_name}: {e}", "WARNING")

#     print("\nNo active WPS application found. Attempting to launch a new, hidden instance...")
#     for com_name in com_names:
#         try:
#             app = win32com.client.Dispatch(com_name)
#             app.Visible = False # Run in the background
#             print_status(f"Launched and connected to new {com_name} instance successfully", "SUCCESS")
#             try:
#                 version = getattr(app, 'Version', 'Unknown')
#                 print_status(f"  Version: {version}")
#             except:
#                 print_status("  Version: Could not determine", "WARNING")
            
#             app.Quit()
#             return True
#         except Exception as e:
#             print_status(f"Failed to launch {com_name} instance: {e}", "FAILURE")

#     print_status("Failed to connect to or launch any WPS COM object.", "FAILURE")
#     print_status("Please ensure WPS Office is correctly installed and registered.", "INFO")
#     return False

# def check_ribbon_xml():
#     """Check if ribbon.xml exists and is valid."""
#     print("\n=== RIBBON XML CHECK ===")
    
#     script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
#     ribbon_path = os.path.join(script_dir, "wps_addin", "ribbon.xml")
    
#     if not os.path.exists(ribbon_path):
#         print_status(f"ribbon.xml not found at: {ribbon_path}", "FAILURE")
#         return False
    
#     print_status(f"ribbon.xml found at: {ribbon_path}", "SUCCESS")
    
#     try:
#         with open(ribbon_path, 'r', encoding='utf-8') as f:
#             content = f.read()
        
#         print_status(f"File readable, size: {len(content)} characters")
        
#         # Check namespace
#         if 'xmlns="http://schemas.microsoft.com/office/2009/07/customui"' in content:
#             print_status("Uses Microsoft Office namespace", "SUCCESS")
#         elif 'xmlns="http://schemas.kingsoft.com/office/2009/07/customui"' in content:
#             print_status("Uses Kingsoft/WPS namespace", "SUCCESS")
#         else:
#             print_status("Namespace not recognized - this might be the issue!", "WARNING")
#             print("  Try changing to: http://schemas.microsoft.com/office/2009/07/customui")
        
#         # Check basic structure
#         if '<customUI' in content and '<ribbon' in content and '<tabs' in content:
#             print_status("Contains required customUI, ribbon, and tabs elements", "SUCCESS")
#             return True
#         else:
#             print_status("Missing one or more required elements (<customUI>, <ribbon>, <tabs>)", "FAILURE")
#             return False
        
#     except Exception as e:
#         print_status(f"Error reading ribbon.xml: {e}", "FAILURE")
#         return False

# def suggest_fixes():
#     """Suggest potential fixes based on common issues."""
#     print("\n=== SUGGESTED FIXES ===")
#     print("Based on the results above, here are some common solutions:")
#     print("1. If registry access failed, try running this script as an **administrator**.")
#     print("2. If LoadBehavior is **2**, the add-in failed to load at startup. Check WPS Office's Trust Center settings to ensure COM add-ins are allowed. Also, verify the add-in's files haven't been moved or deleted.")
#     print("3. If COM registration failed, the add-in's DLL or EXE might not be properly registered. Re-run your add-in's installer or registration command.")
#     print("4. If the ribbon XML check failed, a syntax error or incorrect file path is likely. Review the `ribbon.xml` file for any typos or structural issues.")
#     print("5. Always **restart** WPS Office completely (including ending processes in Task Manager) after making changes to ensure they are picked up.")

# def main():
#     print("WPS Office Add-in Diagnostic Tool")
#     print("=" * 50)
    
#     registry_ok = check_wps_addin_registry()
#     com_ok = check_com_registration()
#     wps_ok = check_wps_installation()
#     wps_com_ok = check_wps_com_objects()
#     xml_ok = check_ribbon_xml()
    
#     print("\n" + "=" * 50)
#     print("=== SUMMARY ===")
#     print(f"Registry entries: {'✅ OK' if registry_ok else '❌ FAILED'}")
#     print(f"COM registration: {'✅ OK' if com_ok else '❌ FAILED'}")
#     print(f"WPS installation: {'✅ OK' if wps_ok else '❌ FAILED'}")
#     print(f"WPS COM access: {'✅ OK' if wps_com_ok else '❌ FAILED'}")
#     print(f"Ribbon XML: {'✅ OK' if xml_ok else '❌ FAILED'}")
    
#     suggest_fixes()

# if __name__ == "__main__":
#     main()
#     input("\nPress Enter to exit...")