This example demonstrates a potential issue with error handling in VBScript.  The GetValue function uses On Error Resume Next to suppress errors, but this can hide critical issues.  The solution provides a more robust approach to error handling.