[Setup]
AppName=Katip
AppVersion=4.0.1
AppPublisher=Berkesas
AppPublisherURL=https://berkesas.github.io/katip
DefaultDirName={commonpf}\Katip
AppCopyright=Copyright (C) 2012-2025 Nazar Mammedov
OutputBaseFilename=katip_setup_x64
DefaultGroupName=Katip
ArchitecturesAllowed=x64compatible

#define VERSION "4.0.1.00"
#define PROGID "Katip"
#define CLSID "{7eef24e6-8944-4840-b2dc-59f6eaede092}"
#define CLASS_FULL_NAME "Katip"
#define COMPANY_NAME "Berkesas"
#define ASSEMBLY_FULL_NAME "Katip, Version=4.0.1.0, Culture=neutral, PublicKeyToken=null"

[Files]
Source: "languages.txt"; DestDir: "{commonappdata}\Katip"; Permissions: everyone-full;
Source: "settings.ini"; DestDir: "{commonappdata}\Katip"; Permissions: everyone-full;
Source: "error.log"; DestDir: "{commonappdata}\Katip"; Permissions: everyone-full;
Source: "LANGUAGES.md"; DestDir: "{commonappdata}\Katip"; Permissions: everyone-full;
Source: "dictionaries\*"; DestDir: "{commonappdata}\Katip\dictionaries"; Flags: recursesubdirs; Permissions: everyone-full;
Source: "locale\*"; DestDir: "{commonappdata}\Katip\locale"; Flags: recursesubdirs; Permissions: everyone-full; 
Source: "katip4.dotm"; DestDir: "{userappdata}\Microsoft\Word\STARTUP"
Source: "hunspellvba.dll"; DestDir: "{app}"
Source: "hunspell-1.7-0.dll"; DestDir: "{app}"

[Icons]
Name: "{group}\Uninstall"; Filename: "{uninstallexe}"; WorkingDir: "{app}"

[Registry]
; Only add {app} to PATH if it doesn't already exist
Root: HKCU; Subkey: "Environment"; \
  ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{app}"; Flags: preservestringtype; \
  Check: not IsAppInUserPath()

[Code]
function GetUserPath(): String;
var
  UserPath: String;
begin
  if RegQueryStringValue(HKEY_CURRENT_USER, 'Environment', 'Path', UserPath) then
    Result := UserPath
  else
    Result := '';
end;

function IsAppInUserPath(): Boolean;
var
  OldPath: String;
  NewPath: String;
begin
  NewPath:= ExpandConstant('{app}');
  OldPath := GetUserPath();
  Result := Pos(';'+NewPath, OldPath) > 0;
end;


