Function f(a, b)
  If IsNumeric(a) And IsNumeric(b) Then
    f = a + b
  Else
    Err.Raise 13, , "Type mismatch"
  End If
End Function

MsgBox f(1, "2")