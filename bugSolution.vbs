Function GetValue(key)
  Dim SetValue
  On Error GoTo ErrorHandler
  SetValue = GetSetting(Section, key)
  On Error GoTo 0
  GetValue = SetValue
  Exit Function
ErrorHandler:
  If Err.Number <> 0 Then
    ' Log the error or handle it appropriately
    WScript.Echo "Error getting setting for key '" & key & "': " & Err.Description
    Err.Clear
    GetValue = ""
  End If
End Function