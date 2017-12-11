Attribute VB_Name = "Module1"
Public cnn As New ADODB.Connection
Public user As String
Private Sub main()
    cnn.ConnectionString = "Provider=microsoft.jet.oledb.4.0; data source=D:\vivek project 2016\database\MAutomation.mdb"
    cnn.Open
    'Form4.Show
    'Form3.Show
    Form1.Show
    'frmunit.Show
    'frmcompany.Show
    'frmcategory.Show
    'frmmedicine.Show
    'frmsupplier.Show
    'frmmedicindetail.Show
    'frmorder.Show
    'frmstock.Show
    'frmsale.Show
    'MDIForm1.Show
End Sub
'Public Function validation1(keyascii As Integer, Optional o1 As Integer) As Integer
'
'
'If Not ((keyascii >= 97 And keyascii <= 122) Or (keyascii >= 65 And keyascii <= 90) Or keyascii = 32 Or keyascii = 8) Then
'    keyascii = 0
'End If
'    validation1 keyascii
'End Function
