; Inno Setup Script for AI Assistant for WPS (Full Application)
; Version 2.1 - Corrected privilege elevation flags

[Setup]
AppId={336AC79A-E98E-4D86-A780-9587723C0C30}
AppName=AI Assistant for WPS
AppVersion=1.0
AppPublisher=Bushbaby
PrivilegesRequired=admin
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
Source: "VC_redist.x86.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: not Is64BitInstallMode
Source: "VC_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: Is64BitInstallMode
Source: "dist\AI_Backend_Server.exe"; DestDir: "{app}"
Source: "dist\AI_Addin_Client_32.exe"; DestDir: "{app}"; Flags: ignoreversion; Check: not Is64BitInstallMode
Source: "dist\AI_Addin_Client_64.exe"; DestDir: "{app}"; Flags: ignoreversion; Check: Is64BitInstallMode
Source: "wps_addin\ribbon.xml"; DestDir: "{app}"
Source: "wps_addin\*.png"; DestDir: "{app}\images\"; Flags: recursesubdirs createallsubdirs
Source: "app\*"; DestDir: "{app}\data\"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{commonstartup}\AI Assistant Backend"; Filename: "{app}\AI_Backend_Server.exe"; WorkingDir: "{app}"
Name: "{group}\Uninstall AI Assistant"; Filename: "{uninstallexe}"

[Run]
; --- Install System Dependencies ---
Filename: "{tmp}\VC_redist.x86.exe"; Parameters: "/install /quiet /norestart"; Check: not Is64BitInstallMode
Filename: "{tmp}\VC_redist.x64.exe"; Parameters: "/install /quiet /norestart"; Check: Is64BitInstallMode

; --- Register the correct COM Add-in after all files are copied ---
; FIXED: Removed 'runhidden' to allow shellexec to elevate privileges
Filename: "{app}\AI_Addin_Client_32.exe"; Parameters: "/regserver"; Flags: postinstall shellexec; Check: not Is64BitInstallMode
Filename: "{app}\AI_Addin_Client_64.exe"; Parameters: "/regserver"; Flags: postinstall shellexec; Check: Is64BitInstallMode


[UninstallRun]
; --- Unregister the correct COM Add-in when the user uninstalls ---
; FIXED: Removed 'runhidden' to allow shellexec to elevate privileges
Filename: "{app}\AI_Addin_Client_32.exe"; Parameters: "/unregserver"; Flags: shellexec; Check: not Is64BitInstallMode
Filename: "{app}\AI_Addin_Client_64.exe"; Parameters: "/unregserver"; Flags: shellexec; Check: Is64BitInstallMode