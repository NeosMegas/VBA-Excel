Attribute VB_Name = "modInterpolateY_X"
Option Explicit

Function Y_X(xValues As Range, yValues As Range, x As Double) As Double
Dim i As Integer
If xValues.Count <> yValues.Count Then
Y_X = 0
Exit Function
End If

If xValues(1) < xValues(xValues.Count) Then

For i = 1 To xValues.Count - 1
DoEvents
If x < xValues(1).Value Then
i = 1
Exit For
End If
If x >= xValues(xValues.Count).Value Then
i = xValues.Count - 1
Exit For
End If
If x >= xValues(i).Value And x < xValues(i + 1).Value Then
Exit For
End If
Next

ElseIf xValues(1) > xValues(xValues.Count) Then

For i = xValues.Count - 1 To 1 Step -1
DoEvents
If x < xValues(i + 1) Then
i = xValues.Count - 1
Exit For
End If
If x >= xValues(1) Then
i = 1
Exit For
End If
If x < xValues(i).Value And x >= xValues(i + 1).Value Then
Exit For
End If
Next

End If

Y_X = yValues(i) + (x - xValues(i)) / (xValues(i + 1) - xValues(i)) * (yValues(i + 1) - yValues(i))
End Function
