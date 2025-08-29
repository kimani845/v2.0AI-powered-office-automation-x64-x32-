# DONE AND RUNNING
"""
Entry point for 64-bit WPS Add-in
Run this script with 64-bit Python to register/use the 64-bit version
"""

import sys
import os
from addin_base_client import log_message

def main():
    """Main entry point for 64-bit operations"""
    log_message("64-bit WPS Add-in Entry Point")
    
    # Force import of 64-bit implementation
    from addin_client64bit import (
        WPSAddin64 as WPSAddin,
        register_server_64bit as register_server,
        unregister_server_64bit as unregister_server,
        run_com_server_64bit as run_com_server,
        check_environment_64bit as check_environment
    )
    
    check_environment()
    
    if len(sys.argv) > 1:
        if sys.argv[1].lower() == '/regserver':
            print("Registering 64-bit WPS Add-in...")
            register_server(WPSAddin)
        elif sys.argv[1].lower() == '/unregserver':
            print("Unregistering 64-bit WPS Add-in...")
            unregister_server(WPSAddin)
        elif sys.argv[1].lower() == '/embedding':
            # This is called by Windows when WPS tries to instantiate the COM object
            run_com_server()
    else:
        print("WPS Office AI Assistant Add-in (64-bit)")
        print("Usage:")
        print("  python run_64bit.py /regserver   - Register the 64-bit add-in")
        print("  python run_64bit.py /unregserver - Unregister the 64-bit add-in")
        print(f"Running as: 64-bit {'PyInstaller bundle' if getattr(sys, 'frozen', False) else 'Python script'}")
        input("\nPress Enter to exit...")

if __name__ == '__main__':
    main()