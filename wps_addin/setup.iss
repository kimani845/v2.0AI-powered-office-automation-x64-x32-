; Inno Setup Script for AI WPS Add-in (32-bit and 64-bit)

[Setup]
; Use a new GUID for this combined installer
AppId={{YOUR-NEW-UNIQUE-GUID-HERE}}
AppName=AI Assistant for WPS
AppVersion=1.0
AppPublisher=Your Company Name
; Let Inno Setup decide the default Program Files directory based on architecture
DefaultDirName={autopf}\AI Assistant for WPS
DefaultGroupName=AI Assistant for WPS
OutputBaseFilename=AI-Assistant-Setup-Combined
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; This tells the installer to run in 64-bit mode on a 64-bit OS
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; --- 32-bit files: Install ONLY on 32-bit Windows ---
Source: "VC_redist.x86.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: not Is64BitInstallMode
Source: "dist32\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Check: not Is64BitInstallMode

; --- 64-bit files: Install ONLY on 64-bit Windows ---
Source: "VC_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: Is64BitInstallMode
Source: "dist64\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Check: Is64BitInstallMode

[Run]
; --- 32-bit tasks: Run ONLY on 32-bit Windows ---
; 1. Install 32-bit VC++ Redistributable
Filename: "{tmp}\VC_redist.x86.exe"; Parameters: "/install /quiet /norestart"; StatusMsg: "Installing required 32-bit system components..."; Check: not Is64BitInstallMode
; 2. Register 32-bit COM server
Filename: "{app}\addin_client.exe"; Parameters: "/regserver"; Flags: postinstall runhidden; Check: not Is64BitInstallMode

; --- 64-bit tasks: Run ONLY on 64-bit Windows ---
; 1. Install 64-bit VC++ Redistributable
Filename: "{tmp}\VC_redist.x64.exe"; Parameters: "/install /quiet /norestart"; StatusMsg: "Installing required 64-bit system components..."; Check: Is64BitInstallMode
; 2. Register 64-bit COM server
Filename: "{app}\addin_client.exe"; Parameters: "/regserver"; Flags: postinstall runhidden; Check: Is64BitInstallMode

[UninstallRun]
; --- Unregister the correct version on uninstall ---
Filename: "{app}\addin_client.exe"; Parameters: "/unregserver"; Flags: runhidden; Check: not Is64BitInstallMode
Filename: "{app}\addin_client.exe"; Parameters: "/unregserver"; Flags: runhidden; Check: Is64BitInstallMode