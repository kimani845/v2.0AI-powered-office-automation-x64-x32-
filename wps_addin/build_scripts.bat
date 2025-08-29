@echo off
echo Building WPS Add-in for both 32-bit and 64-bit architectures...

echo.
echo ========================================
echo Building 32-bit version...
echo ========================================
pyinstaller build_32bit.spec --clean

echo.
echo ========================================
echo Building 64-bit version...
echo ========================================
pyinstaller build_64bit.spec --clean

echo.
echo ========================================
echo Build Complete!
echo ========================================
echo 32-bit executable: dist\WPSAddin_32bit.exe
echo 64-bit executable: dist\WPSAddin_64bit.exe
echo.
echo To register:
echo   For 32-bit WPS: dist\WPSAddin_32bit.exe /regserver
echo   For 64-bit WPS: dist\WPSAddin_64bit.exe /regserver
echo.
pause