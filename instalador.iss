; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{C77B00F8-657A-4DE8-834E-8626BC064A08}
AppName=AoYind 3
AppVersion=3.0
;AppVerName=AoYind 3 3.0
AppPublisher=YindSoft
AppPublisherURL=http://www.aoyind.com/
AppSupportURL=http://www.aoyind.com/
AppUpdatesURL=http://www.aoyind.com/
DefaultDirName={pf}\AoYind 3
DefaultGroupName=AoYind 3
OutputDir=C:\Users\Javier\Documents\Ao\Cliente\Instalador
OutputBaseFilename=aoyind3
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 0,6.1

[Files]
Source: "C:\Users\Javier\Documents\Ao\Cliente\Paquete Instalador\AoYindCliente.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\Javier\Documents\Ao\Cliente\Ruritania.TTF"; DestDir: "{fonts}"; FontInstall: "Ruritania"; Flags: onlyifdoesntexist uninsneveruninstall 
Source: "C:\Users\Javier\Documents\Ao\Cliente\Paquete Instalador\msvbvm60.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "C:\Users\Javier\Documents\Ao\Cliente\Paquete Instalador\zlib.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: "C:\Users\Javier\Documents\Ao\Cliente\Paquete Instalador\dx8vb.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "C:\Users\Javier\Documents\Ao\Cliente\Paquete Instalador\ijl11.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: "C:\Users\Javier\Documents\Ao\Cliente\Paquete Instalador\aamd532.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: "C:\Users\Javier\Documents\Ao\Cliente\Paquete Instalador\msvcrt.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: "C:\Users\Javier\Documents\Ao\Cliente\Paquete Instalador\VB6ES.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: "C:\Users\Javier\Documents\Ao\Cliente\Paquete Instalador\CSWSK32.ocx"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace sharedfile regserver 
Source: "C:\Users\Javier\Documents\Ao\Cliente\Paquete Instalador\MSCOMCTL.ocx"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace sharedfile regserver 

Source: "C:\Users\Javier\Documents\Ao\Cliente\Paquete Instalador\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\AoYind 3"; Filename: "{app}\AoYindCliente.exe"
Name: "{group}\{cm:UninstallProgram,AoYind 3}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\AoYind 3"; Filename: "{app}\AoYindCliente.exe"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\AoYind 3"; Filename: "{app}\AoYindCliente.exe"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\AoYindCliente.exe"; Description: "{cm:LaunchProgram,AoYind 3}"; Flags: nowait postinstall skipifsilent

