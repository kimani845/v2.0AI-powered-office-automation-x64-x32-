"""
Backward compatibility wrapper for the original wps_addin.py file.
This ensures existing references continue to work while using the new architecture-aware launcher.
"""

from wps_addin.addin_clientlauncher import main

# For backward compatibility, expose the WPSAddin class
def get_wps_addin_class():
    """Get the appropriate WPS Add-in class based on system architecture"""
    import platform
    import sys
    
    is_64bit = sys.maxsize > 2**32
    
    if is_64bit:
        from wps_addin.addin_client64bit import WPSAddin64
        return WPSAddin64
    else:
        from wps_addin.addin_client_32bit import WPSAddin32
        return WPSAddin32

# Create the class instance for COM registration
WPSAddin = get_wps_addin_class()

if __name__ == '__main__':
    main()