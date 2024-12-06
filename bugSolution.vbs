Function GetValue(key)
  Dim result
  On Error GoTo ErrorHandler
  result = Application("Get" & key)
  Exit Function
ErrorHandler:
  If Err.Number <> 0 Then
    'Log error details or handle it appropriately.
    'Example:  WScript.Echo "Error getting value for key '" & key & "': " & Err.Description
    result = Null 'Or another appropriate default value
  End If
  GetValue = result
End Function