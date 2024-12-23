Function MyFunction(param1, param2)
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise vbError, , "Parameters cannot be empty", 1 'Added a context number for better debugging
  End If
  ' ... rest of the function ...
End Function

' Example of how to handle the error:
On Error GoTo ErrorHandler

Dim result
result = MyFunction( , ) 'This will cause an error
MsgBox "This line should not be reached"
Exit Sub

ErrorHandler:
Select Case Err.Number
  Case vbError
    MsgBox Err.Description & " Error Number: " & Err.Number
  Case Else
    MsgBox "An unexpected error occurred."
End Select
Err.Clear
