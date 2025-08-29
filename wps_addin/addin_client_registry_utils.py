"""
Main launcher that detects system architecture and launches appropriate add-in.
This file serves as the entry point and determines whether to use 32-bit or 64-bit logic.
"""

import sys
import platform
import os
from wps_addin.addin_base_client import log_message

def detect_architecture():
    """Detect if we're running on 32-bit or 64-bit system"""
    architecture = platform.architecture()[0]
    machine = platform.machine()
    
    log_message(f"System architecture: {architecture}")
    log_message(f"Machine type: {machine}")
    
    # Check if we're running 32-bit Python on 64-bit system
    is_64bit_system = machine.endswith('64') or 'AMD64' in machine
    is_64bit_python = sys.maxsize > 2**32
    
    log_message(f"64-bit system: {is_64bit_system}")
    log_message(f"64-bit Python: {is_64bit_python}")
    
    return is_64bit_python

def main():
    """Main entry point that routes to appropriate architecture-specific implementation"""
    log_message("WPS Office AI Assistant Add-in Launcher")
    
    is_64bit = detect_architecture()
    
    if is_64bit:
        log_message("Detected 64-bit environment - using 64-bit implementation")
        from wps_addin.addin_client64bit import (
            WPSAddin64 as WPSAddin, 
            register_server_64bit as register_server,
            unregister_server_64bit as unregister_server,
            run_com_server_64bit as run_com_server,
            check_environment_64bit as check_environment
        )
    else:
        log_message("Detected 32-bit environment - using 32-bit implementation")
        from wps_addin.addin_client_32bit import (
            WPSAddin32 as WPSAddin,
            register_server_32bit as register_server,
            unregister_server_32bit as unregister_server,
            run_com_server_32bit as run_com_server,
            check_environment_32bit as check_environment
        )
    
    # Main command logic
    check_environment()
    
    if len(sys.argv) > 1:
        if sys.argv[1].lower() == '/regserver':
            register_server(WPSAddin)
        elif sys.argv[1].lower() == '/unregserver':
            unregister_server(WPSAddin)
        elif sys.argv[1].lower() == '/embedding':
            # This is called by Windows when WPS tries to instantiate the COM object
            run_com_server()
    else:
        arch_type = "64-bit" if is_64bit else "32-bit"
        print(f"WPS Office AI Assistant Add-in ({arch_type})")
        print("Usage:")
        print("  /regserver   - Register the add-in")
        print("  /unregserver - Unregister the add-in")
        print(f"Running as: {arch_type} {'PyInstaller bundle' if getattr(sys, 'frozen', False) else 'Python script'}")
        input("\nPress Enter to exit...")

if __name__ == '__main__':
    main()