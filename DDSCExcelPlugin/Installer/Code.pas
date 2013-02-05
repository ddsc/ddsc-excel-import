procedure CheckRegistry(Install: Boolean);
var
  OfficeVersions: TArrayOfString;
  I: Integer;
  J: Integer;
  Installed: Boolean;
  Key: String;
  Value: String;
  Keys: TArrayOfString;
  AlertMessage: String;
  CurrentValue: String;
  NewKey: String;
  KeyNumber: Integer;
  NewKeys: TArrayOfString;
begin
  // Location of Add-In
  Value := ExpandConstant('"{app}\DDSCExcelPlugin.xll"');
  
  // List office versions
  if RegGetSubkeyNames(HKEY_CURRENT_USER, 'Software\Microsoft\Office', OfficeVersions) then
    begin
      // Check all the office version
      for I := 0 to GetArrayLength(OfficeVersions)-1 do
        begin
          // Initialize Installation Info
          KeyNumber := 0;
          Installed := false;
          // Check if Excel is installed and has Options
          Key := 'Software\Microsoft\Office\' + OfficeVersions[I] + '\Excel\Options';
          if RegKeyExists(HKEY_CURRENT_USER, Key) then
            begin
              // List all the add-ins currently being shown - for this read all value name
              if RegGetValueNames(HKEY_CURRENT_USER, Key, Keys) then
                begin
                  // Process each value name
                  for J := 0 to GetArrayLength(Keys)-1 do
                    begin
                      // Check if it is really an ADD-IN
                      if (Length(Keys[J]) >= 4) AND (Copy(Keys[J], 1, 4) = 'OPEN') then
                        begin
                          // Read the add-in path
                          if RegQueryStringValue(HKEY_CURRENT_USER, Key, Keys[J], CurrentValue) then
                            begin
                              // Check if it is the add-in we are installing
                              if CompareText(Value, CurrentValue) = 0 then
                                Installed := true;
                              else
                                begin
                                  // Store all other add-ins in another array
                                  // Will be used for uninstall
                                  SetArrayLength(NewKeys, KeyNumber + 1);
                                  NewKeys[KeyNumber] := CurrentValue;
                                  KeyNumber := KeyNumber + 1;
                                end
                            end
                        end
                    end
                end
                if Installed then
                  begin
                    // Are we trying to uninstall?
                    if not Install then
                      begin
                        // Re-serialize all other add-ins
                        for J := 0 to KeyNumber-1 do
                          begin
                            NewKey := 'OPEN';
                            if J > 0 then NewKey := NewKey + IntToStr(J);
                              RegWriteStringValue(HKEY_CURRENT_USER, Key, NewKey, NewKeys[J]);
                          end
                          // Delete additional keys
                          repeat
                            NewKey := 'OPEN';
                            if J > 0 then NewKey := NewKey + IntToStr(J);
                          until (Not RegDeleteValue(HKEY_CURRENT_USER, Key, NewKey));
                      end
                      AlertMessage := 'Installed';
                  end
                else
                  // Not installed
                  begin
                    // We are trying to install - add to the last
                    if Install then
                      begin
                        NewKey := 'OPEN';
                        if KeyNumber > 0 then
                          NewKey := NewKey + IntToStr(KeyNumber);
                          RegWriteStringValue(HKEY_CURRENT_USER, Key, NewKey, Value);
                      end
                      AlertMessage := 'Not-Installed';
                  end
                  AlertMessage := 'For office version ' + OfficeVersions[I] + ', add-in is ' + AlertMessage;
                  //MsgBox(AlertMessage, mbInformation, MB_OK);
            end
        end
    end
end;
      
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  case CurUninstallStep of
    usPostUninstall:
      begin
        CheckRegistry(false);
      end;
  end;
end;