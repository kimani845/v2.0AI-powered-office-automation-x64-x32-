"""
WPS Office Add-in Diagnostic Tool
Run this to diagnose why your add-in isn't showing buttons
"""
import winreg
import os
import sys

def check_wps_addin_registry():
    """Check all possible WPS add-in registry locations"""
    print("=== WPS ADD-IN REGISTRY CHECK ===")
    
    addin_name = "WPSAIAddin.Connect"
    base_paths = [
        r"Software\Kingsoft\Office\Addins",
        r"Software\WPS\Office\Addins",
        r"Software\WPS Office\Addins"
    ]
    
    found_entries = []
    
    for base_path in base_paths:
        full_path = f"{base_path}\\{addin_name}"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, full_path) as key:
                print(f"✓ FOUND: HKCU\\{full_path}")
                
                # Check all values
                try:
                    load_behavior = winreg.QueryValueEx(key, "LoadBehavior")[0]
                    print(f"  LoadBehavior: {load_behavior}")
                    
                    if load_behavior == 3:
                        print("  ✓ LoadBehavior is correct (3 = Load at startup)")
                    elif load_behavior == 2:
                        print("  ✗ LoadBehavior is 2 (Add-in failed to load)")
                    else:
                        print(f"  ⚠ LoadBehavior is {load_behavior} (unexpected)")
                        
                except FileNotFoundError:
                    print("  ✗ LoadBehavior not set")
                
                try:
                    clsid = winreg.QueryValueEx(key, "CLSID")[0]
                    print(f"  CLSID: {clsid}")
                except FileNotFoundError:
                    print("  ✗ CLSID not set")
                
                try:
                    description = winreg.QueryValueEx(key, "Description")[0]
                    print(f"  Description: {description}")
                except FileNotFoundError:
                    print("  ✗ Description not set")
                    
                found_entries.append((base_path, load_behavior if 'load_behavior' in locals() else None))
                
        except FileNotFoundError:
            print(f"✗ NOT FOUND: HKCU\\{full_path}")
    
    if not found_entries:
        print("✗ No WPS add-in registry entries found!")
        return False
    else:
        print(f"✓ Found {len(found_entries)} registry entries")
        return True

def check_com_registration():
    """Check COM server registration"""
    print("\n=== COM SERVER REGISTRATION CHECK ===")
    
    clsid = "{cf0b4f12-56e5-4818-b400-b3f2660e0a3c}"
    progid = "WPSAIAddin.Connect"
    
    # Check CLSID registration
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}") as key:
            print(f"✓ CLSID registered: {clsid}")
            
            # Check LocalServer32
            try:
                with winreg.OpenKey(key, "LocalServer32") as server_key:
                    exe_path = winreg.QueryValueEx(server_key, "")[0]
                    print(f"  LocalServer32: {exe_path}")
                    if os.path.exists(exe_path.strip('"')):
                        print("  ✓ Executable exists")
                    else:
                        print("  ✗ Executable NOT found")
            except FileNotFoundError:
                print("  ✗ LocalServer32 not found")
                
    except FileNotFoundError:
        print(f"✗ CLSID not registered: {clsid}")
    
    # Check ProgID registration
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid) as key:
            print(f"✓ ProgID registered: {progid}")
    except FileNotFoundError:
        print(f"✗ ProgID not registered: {progid}")

def check_wps_installation():
    """Check WPS Office installation"""
    print("\n=== WPS OFFICE INSTALLATION CHECK ===")
    
    install_paths = [
        r"SOFTWARE\Kingsoft\Office",
        r"SOFTWARE\WPS\Office",
        r"SOFTWARE\WPS Office"
    ]
    
    wps_found = False
    
    for path in install_paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                print(f"✓ WPS registry found: {path}")
                wps_found = True
                
                try:
                    install_root = winreg.QueryValueEx(key, "InstallRoot")[0]
                    print(f"  Install path: {install_root}")
                    if os.path.exists(install_root):
                        print("  ✓ Installation directory exists")
                    else:
                        print("  ✗ Installation directory NOT found")
                except FileNotFoundError:
                    print("  ⚠ InstallRoot not found in registry")
                    
        except FileNotFoundError:
            print(f"✗ Not found: {path}")
    
    if not wps_found:
        print("✗ WPS Office installation not detected!")
        return False
    
    return True

def check_wps_com_objects():
    """Test WPS COM object access"""
    print("\n=== WPS COM OBJECT TEST ===")
    
    com_names = [
        "kwps.Application",
        "WPS.Application", 
        "Kingsoft.WPS.Application"
    ]
    
    for com_name in com_names:
        try:
            import win32com.client
            app = win32com.client.GetActiveObject(com_name)
            if app:
                print(f"✓ {com_name} - Connected successfully")
                try:
                    version = getattr(app, 'Version', 'Unknown')
                    print(f"  Version: {version}")
                except:
                    print("  Version: Could not determine")
                return True
        except Exception as e:
            print(f"✗ {com_name} - {e}")
    
    print("⚠ No active WPS application found. Start WPS Writer and try again.")
    return False

def check_ribbon_xml():
    """Check if ribbon.xml exists and is valid"""
    print("\n=== RIBBON XML CHECK ===")
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    ribbon_path = os.path.join(script_dir,"wps_addin", "ribbon.xml")
    
    if not os.path.exists(ribbon_path):
        print(f"✗ ribbon.xml not found at: {ribbon_path}")
        return False
    
    print(f"✓ ribbon.xml found at: {ribbon_path}")
    
    try:
        with open(ribbon_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        print(f"✓ File readable, size: {len(content)} characters")
        
        # Check namespace
        if 'xmlns="http://schemas.microsoft.com/office/2009/07/customui"' in content:
            print("✓ Uses Microsoft Office namespace")
        elif 'xmlns="http://schemas.kingsoft.com/office/2009/07/customui"' in content:
            print("✓ Uses Kingsoft/WPS namespace")
        else:
            print("⚠ Namespace not recognized - this might be the issue!")
            print("  Try changing to: http://schemas.microsoft.com/office/2009/07/customui")
        
        # Check basic structure
        if '<customUI' in content:
            print("✓ Contains customUI element")
        else:
            print("✗ Missing customUI element")
            
        if '<ribbon' in content:
            print("✓ Contains ribbon element")
        else:
            print("✗ Missing ribbon element")
            
        return True
        
    except Exception as e:
        print(f"✗ Error reading ribbon.xml: {e}")
        return False

def suggest_fixes():
    """Suggest potential fixes"""
    print("\n=== SUGGESTED FIXES ===")
    print("1. If LoadBehavior is 2, the add-in tried to load but failed:")
    print("   - Check WPS Office security settings")
    print("   - File → Options → Trust Center → Add-ins")
    print("   - Ensure COM add-ins are allowed")
    
    print("\n2. Try manually enabling the add-in:")
    print("   - In WPS: File → Options → Add-ins")
    print("   - Manage: COM Add-ins → Go")
    print("   - Check if 'AI Office Automation' is listed")
    print("   - If unchecked, check it manually")
    
    print("\n3. If the add-in appears in COM Add-ins but buttons don't show:")
    print("   - The issue is likely in the ribbon XML")
    print("   - Try a different XML namespace")
    print("   - Verify the XML structure")
    
    print("\n4. Restart WPS Office completely:")
    print("   - Close all WPS applications")
    print("   - End any WPS processes in Task Manager")
    print("   - Restart WPS Writer")
    
    print("\n5. Check WPS Office version compatibility:")
    print("   - Older WPS versions have different COM interfaces")
    print("   - Try registering with different WPS registry paths")

def main():
    print("WPS Office Add-in Diagnostic Tool")
    print("=" * 50)
    
    registry_ok = check_wps_addin_registry()
    com_ok = check_com_registration()
    wps_ok = check_wps_installation()
    wps_com_ok = check_wps_com_objects()
    xml_ok = check_ribbon_xml()
    
    print("\n=== SUMMARY ===")
    print(f"Registry entries: {'✓' if registry_ok else '✗'}")
    print(f"COM registration: {'✓' if com_ok else '✗'}")
    print(f"WPS installation: {'✓' if wps_ok else '✗'}")
    print(f"WPS COM access: {'✓' if wps_com_ok else '✗'}")
    print(f"Ribbon XML: {'✓' if xml_ok else '✗'}")
    
    suggest_fixes()

if __name__ == "__main__":
    main()
    input("\nPress Enter to exit...")