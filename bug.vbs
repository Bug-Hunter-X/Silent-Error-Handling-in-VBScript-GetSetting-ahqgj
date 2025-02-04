Function GetValue(key)
  On Error Resume Next
  SetValue = GetSetting(Section, key)
  If Err.Number <> 0 Then
    Err.Clear
    SetValue = ""
  End If
  GetValue = SetValue
End Function