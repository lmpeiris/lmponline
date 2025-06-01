Attribute VB_Name = "Module1"
Dim xy As String



Function fact(ByVal l As Integer) As Double
xy = "1"
GoTo count
count:
xy = Val(l * (l - 1) * xy)
l = l - 2
GoTo processer
processer:
If l < 2 Then
fact = Val(xy)
Exit Function
Else
GoTo count
End If
End Function

Function sinh(w As Double) As Double
xy = (Exp(w)) - Exp(-1 * w)
sinh = Val(xy) / 2
End Function

Function cosh(w As Double) As Double
xy = (Exp(w)) + Exp(-1 * w)
cosh = Val(xy) / 2
End Function

Function tanh(w As Double) As Double
xy = (Exp(w)) + Exp(-1 * w)
tanh = Val(xy) / (Exp(w)) - Exp(-1 * w)
End Function

