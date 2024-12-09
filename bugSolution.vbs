Function MyFunction(param1, param2)
  If IsEmpty(param1) Or IsEmpty(param2) Then
    'Instead of Err.Raise, provide a default value or handle gracefully
    param1 = IIf(IsEmpty(param1), "", param1) 
    param2 = IIf(IsEmpty(param2), "", param2)
  End If
  ' ... rest of the function ...
End Function