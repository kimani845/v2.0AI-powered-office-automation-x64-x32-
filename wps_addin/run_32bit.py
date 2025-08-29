# DONE AND RUNNING

"""
Entry point for 32-bit WPS Add-in
Run this script with 32-bit Python to register/use the 32-bit version
"""

import sys
import os
from addin_base_client import log_message

def main():
    """Main entry point for 32-bit operations"""
    log_message("32-bit WPS Add-in Entry Point")
    
    # Force import of 32-bit implementation
    from addin_client_32bit import (
        WPSAddin32 as WPSAddin,
        register_server_32bit as register_server,
        unregister_server_32bit as unregister_server,
        run_com_server_32bit as run_com_server,
        check_environment_32bit as check_environment
    )
    
    check_environment()
    
    if len(sys.argv) > 1:
        if sys.argv[1].lower() == '/regserver':
            print("Registering 32-bit WPS Add-in...")
            register_server(WPSAddin)
        elif sys.argv[1].lower() == '/unregserver':
            print("Unregistering 32-bit WPS Add-in...")
            unregister_server(WPSAddin)
        elif sys.argv[1].lower() == '/embedding':
            # This is called by Windows when WPS tries to instantiate the COM object
            run_com_server()
    else:
        print("WPS Office AI Assistant Add-in (32-bit)")
        print("Usage:")
        print("  python run_32bit.py /regserver   - Register the 32-bit add-in")
        print("  python run_32bit.py /unregserver - Unregister the 32-bit add-in")
        print(f"Running as: 32-bit {'PyInstaller bundle' if getattr(sys, 'frozen', False) else 'Python script'}")
        input("\nPress Enter to exit...")

if __name__ == '__main__':
    main()