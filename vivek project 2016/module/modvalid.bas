Attribute VB_Name = "Module2"
Public Function key(KeyAscii As Integer, Textbox As String) As Integer
    If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 8) Then
    KeyAscii = 0
End If
'    keyascii = validation1
If ((Len(Textbox) = 0) And KeyAscii = 32) Or (Right(Textbox, 1) = " " And KeyAscii = 32) Then
    KeyAscii = 0
End If
key = KeyAscii
End Function

Public Function keynum(KeyAscii As Integer, Textbox As String) As Integer
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
    KeyAscii = 0
End If
keynum = KeyAscii
End Function
Public Function keyboth(KeyAscii As Integer, Textbox As String) As Integer
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90)) Then
    KeyAscii = 0
End If
keyboth = KeyAscii
End Function

'Public Function maxsizevalid(max As Integer)
'If Len(txtc_name.Text) < 7 Then
'    MsgBox ("text length must be grater then 6")
'    txtc_name.SetFocus
'    selected
'    Exit Sub
'End If
'End Function
'
'

