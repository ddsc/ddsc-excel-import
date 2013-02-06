[Setup]
AppName=DDSC Excel Plugin
AppVersion=1.0.0.0
AppCopyright=Fugro GeoServices B.V.
AppId={{0C7AE33C-DA71-43B9-999B-47372EC6BBE0}
DefaultDirName={pf}\DDSC
AppPublisher=Fugro GeoServices B.V.
AppPublisherURL=http://www.fugro.nl/contact/fgsbv/
AppContact=M. Kodde
AppSupportPhone=070-3170928
VersionInfoVersion=1.0.0.0
VersionInfoCompany=Fugro GeoServices B.V.
VersionInfoDescription=DDSC Excel Plugin
VersionInfoCopyright=Fugro GeoServices B.V.
VersionInfoProductName=DDSC Excel Plugin
VersionInfoProductVersion=1.0.0.0
OutputBaseFileName="DDSCExcelPlugin"

[Files]
Source: "..\bin\Debug\DDSCExcelPlugin.xll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\bin\Debug\DDSCExcelPlugin.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\bin\Debug\DDSCExcelPlugin.dna"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\..\template\DDSC_import_location_timeseries.xls"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\..\template\DDSC_import_location_timeseries.xlt"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: "HKCU"; Subkey: "{code:GetExcelVersion}"; ValueType: string; ValueName: "{code:GetValueName|{code:GetExcelVersion}}"; ValueData: """{app}\DDSCExcelPlugin.xll"""; Flags: createvalueifdoesntexist

[Code]
function InitializeSetup(): Boolean;
var
  WindowHandle: HWND;
begin
  WindowHandle := FindWindowByClassName('XLMAIN');
  if (WindowHandle = 0) then
    result := True
  else
    begin
      MsgBox('Excel is running, please close and try again!', mbInformation, MB_OK);
      result := False;    
    end;
end;

function GetExcelVersion(Param: String):String;
var
  OfficeVersions: TArrayOfString;
  I: Integer;
  Key: String;
begin
// List office versions
  if RegGetSubkeyNames(HKEY_CURRENT_USER, 'Software\Microsoft\Office', OfficeVersions) then
    begin
      // Check all the office version
      for I := 0 to GetArrayLength(OfficeVersions)-1 do
        begin
          // Check if Excel is installed and has Options
          Key := 'Software\Microsoft\Office\' + OfficeVersions[I] + '\Excel\Options';
          if RegKeyExists(HKEY_CURRENT_USER, Key) then
            break;
        end;
        //MsgBox('ExcelVersion=' + Key, mbInformation, MB_OK);
        result := Key;
    end;
end;

function GetValueName(Key: String):String;
var
  ValueName: string;
  Keys: TArrayOfString;
  J: Integer;
  I: Integer;
begin
  if RegKeyExists(HKEY_CURRENT_USER, Key) then
    begin
      // List all the add-ins currently being shown - for this read all value name
      if RegGetValueNames(HKEY_CURRENT_USER, Key, Keys) then
        begin
          // Process each value name
          I := 0;
          for J := 0 to GetArrayLength(Keys)-1 do
            begin
              // Check if it is really an ADD-IN
              if (Length(Keys[J]) >= 4) AND (Copy(Keys[J], 1, 4) = 'OPEN') then
                I := I + 1;
            end;
            if (I = 0) then
              ValueName := 'OPEN'
            else
              ValueName := 'OPEN' + IntToStr(I);
            //MsgBox('ValueName=' + ValueName, mbConfirmation, MB_OK);
            result := ValueName;
        end;
    end;
end;

