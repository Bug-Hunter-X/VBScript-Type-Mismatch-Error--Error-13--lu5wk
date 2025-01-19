Function f(a, b)
  On Error Resume Next
  If IsNumeric(a) And IsNumeric(b) Then
    f = CDbl(a) + CDbl(b) 'Explicit type conversion
  Else
    f = Null ' Handle error gracefully
    Err.Clear ' Clear the error object
  End If
  On Error GoTo 0
End Function

MsgBox f(1, "2") ' Returns 3
MsgBox f("a", 2) ' Returns Null