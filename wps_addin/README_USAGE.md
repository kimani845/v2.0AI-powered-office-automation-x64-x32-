# WPS Add-in Usage Guide

## Running During Development

### For 32-bit WPS Office:
```bash
# Using 32-bit Python
python run_32bit.py /regserver    # Register
python run_32bit.py /unregserver  # Unregister
```

### For 64-bit WPS Office:
```bash
# Using 64-bit Python
python run_64bit.py /regserver    # Register
python run_64bit.py /unregserver  # Unregister
```

## Building Executables

### Prerequisites:
- Install PyInstaller: `pip install pyinstaller`
- Ensure you have both 32-bit and 64-bit Python environments if you want to build both

### Build Both Versions:
```bash
# Run the batch script to build both
build_scripts.bat
```

### Build Individual Versions:
```bash
# Build 32-bit only
pyinstaller build_32bit.spec --clean

# Build 64-bit only
pyinstaller build_64bit.spec --clean
```

## Using the Built Executables

After building, you'll have:
- `dist/WPSAddin_32bit.exe` - For 32-bit WPS Office
- `dist/WPSAddin_64bit.exe` - For 64-bit WPS Office

### Registration:
```bash
# Register 32-bit version (run as Administrator)
dist\WPSAddin_32bit.exe /regserver

# Register 64-bit version (run as Administrator)
dist\WPSAddin_64bit.exe /regserver
```

### Unregistration:
```bash
# Unregister 32-bit version
dist\WPSAddin_32bit.exe /unregserver

# Unregister 64-bit version
dist\WPSAddin_64bit.exe /unregserver
```

## Architecture Detection

The system automatically detects:
- Whether you're running 32-bit or 64-bit Python
- Whether you're running as a PyInstaller bundle or Python script
- Applies the appropriate COM registration method

## File Structure

```
├── wps_addin_base.py          # Shared functionality and base class
├── wps_addin_32bit.py         # 32-bit specific implementation
├── wps_addin_64bit.py         # 64-bit specific implementation
├── wps_registry_utils.py      # Shared registry utilities
├── wps_addin_launcher.py      # Auto-detecting launcher (legacy)
├── wps_addin.py              # Backward compatibility wrapper
├── run_32bit.py              # 32-bit entry point
├── run_64bit.py              # 64-bit entry point
├── build_32bit.spec          # PyInstaller spec for 32-bit
├── build_64bit.spec          # PyInstaller spec for 64-bit
└── build_scripts.bat         # Batch script to build both versions
```

## Important Notes

1. **Administrator Rights**: Registration requires Administrator privileges
2. **WPS Restart**: After registration, restart WPS Office to see the add-in
3. **Architecture Matching**: Use the version that matches your WPS Office installation (32-bit or 64-bit)
4. **Backend Server**: Ensure your backend server is running at `http://127.0.0.1:8000`