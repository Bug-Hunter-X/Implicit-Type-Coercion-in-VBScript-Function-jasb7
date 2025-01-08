Function MyFunction(param1 As Variant, param2 As Variant)
  ' Explicit declaration of variable types
  Dim result As Variant
  If IsNumeric(param1) And IsNumeric(param2) Then
    result = param1 + param2
  Else
    result = "Error: Parameters must be numeric"
  End If
  MyFunction = result
End Function