Attribute VB_Name = "modTextFunctions"
Public Function RegExpExtract(Text As String, Pattern As String, Optional item As Integer = 1) As String
    On Error GoTo ErrHandl
    Dim regex As Object, matches As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = Pattern
    regex.Global = True
    If regex.TEST(Text) Then
        Set matches = regex.Execute(Text)
        RegExpExtract = matches.item(item - 1)
        Exit Function
    End If
ErrHandl:
    RegExpExtract = CVErr(xlErrValue)
End Function

Public Function SplitString(s As String, delimeter As String, Optional index As Integer = 0) As Variant
    Dim str() As String
    str = Split(s, delimeter)
    If index = 0 Then
        SplitString = UBound(str) + 1
    ElseIf index <= UBound(str) + 1 Then
        SplitString = str(index - 1)
    Else
        SplitString = -1
    End If
End Function
