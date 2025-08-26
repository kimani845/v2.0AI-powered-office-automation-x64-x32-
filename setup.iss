; Inno Setup Script for AI Assistant for WPS (Full Application)
; Version 2.0 - Bundles add-in, backend server, and data files.

[Setup]
; Use a unique GUID for your project
AppId={{336AC79A-E98E-4D86-A780-9587723C0C30}}
AppName=WPS AI Assistant
AppVersion=1.0
AppPublisher=Bushbaby
; The installer must run as admin to write to Program Files and register the COM server
PrivilegesRequired=admin
; {autopf} correctly chooses "Program Files" or "Program Files (x86)"
DefaultDirName={autopf}\AI Assistant for WPS
DefaultGroupName=AI Assistant for WPS
OutputBaseFilename=AI-Assistant-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; These files must be in your project's root directory
Source: "VC_redist.x86.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: not Is64BitInstallMode
Source: "VC_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: Is64BitInstallMode

; --- Add-in and Backend Executables (from 'dist' folder) ---
; NOTE: We install the backend server for both 32-bit and 64-bit users.
Source: "dist\AI_Backend_Server.exe"; DestDir: "{app}"

; Install the correct 32-bit or 64-bit add-in client
Source: "dist\AI_Addin_Client_32.exe"; DestDir: "{app}"; Flags: ignoreversion; Check: not Is64BitInstallMode
Source: "dist\AI_Addin_Client_64.exe"; DestDir: "{app}"; Flags: ignoreversion; Check: Is64BitInstallMode

; --- Ribbon, Images, and other assets (from 'wps_addin' folder) ---
; The add-in client needs these files to find its UI resources.
Source: "wps_addin\ribbon.xml"; DestDir: "{app}"
Source: "wps_addin\*.png"; DestDir: "{app}\images\"; Flags: recursesubdirs createallsubdirs

; --- Application Data (from 'app' folder) ---
; This copies the 'agents' folder and anything else inside the 'app' directory.
Source: "app\*"; DestDir: "{app}\data\"; Flags: recursesubdirs createallsubdirs


[Icons]
; This creates a shortcut in the Windows Startup folder to automatically launch
; the backend server when any user logs in.
Name: "{commonstartup}\AI Assistant Backend"; Filename: "{app}\AI_Backend_Server.exe"; WorkingDir: "{app}"

; Optional: Create a Start Menu shortcut to uninstall the application
Name: "{group}\Uninstall AI Assistant"; Filename: "{uninstallexe}"


[Run]
; --- Install System Dependencies ---
Filename: "{tmp}\VC_redist.x86.exe"; Parameters: "/install /quiet /norestart"; Check: not Is64BitInstallMode
Filename: "{tmp}\VC_redist.x64.exe"; Parameters: "/install /quiet /norestart"; Check: Is64BitInstallMode

; --- Register the correct COM Add-in after all files are copied ---
Filename: "{app}\AI_Addin_Client_32.exe"; Parameters: "/regserver"; Flags: postinstall runhidden; Check: not Is64BitInstallMode
Filename: "{app}\AI_Addin_Client_64.exe"; Parameters: "/regserver"; Flags: postinstall runhidden; Check: Is64BitInstallMode


[UninstallRun]
; --- Unregister the correct COM Add-in when the user uninstalls ---
Filename: "{app}\AI_Addin_Client_32.exe"; Parameters: "/unregserver"; Flags: runhidden; Check: not Is64BitInstallMode
Filename: "{app}\AI_Addin_Client_64.exe"; Parameters: "/unregserver"; Flags: runhidden; Check: Is64BitInstallMode