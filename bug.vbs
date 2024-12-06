Function GetValue(key)
  If Err.Number <> 0 Then
    Err.Clear
  End If
  On Error Resume Next
  GetValue = Application("Get" & key)
  If Err.Number <> 0 Then
    Err.Clear
    GetValue = Null
  End If
  On Error GoTo 0
End Function