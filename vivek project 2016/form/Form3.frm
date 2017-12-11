VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmsupplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12810
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   12810
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3135
      Left            =   3840
      TabIndex        =   19
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5530
      _Version        =   393216
      BackColorBkg    =   -2147483633
   End
   Begin VB.Frame Confirmation 
      Caption         =   "Confirmation"
      Height          =   975
      Left            =   8160
      TabIndex        =   16
      Top             =   3600
      Width           =   2655
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Manupilation 
      Caption         =   "Manupilation"
      Height          =   975
      Left            =   5280
      TabIndex        =   13
      Top             =   3600
      Width           =   2655
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Supplier Details"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12615
      Begin VB.TextBox txtStreet 
         Height          =   375
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtColony 
         Height          =   375
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   5
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txtCity 
         Height          =   375
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   6
         Top             =   3480
         Width           =   2175
      End
      Begin VB.ComboBox ComboComp 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   1440
         MaxLength       =   35
         TabIndex        =   9
         Top             =   4920
         Width           =   2175
      End
      Begin VB.TextBox txtContact 
         Height          =   375
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   7
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox txtSID 
         Height          =   405
         Left            =   1440
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtSName 
         Height          =   375
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Street 
         Caption         =   "Street No/ H.No"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Colony"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "City"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Email ID"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Contact No."
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Company"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Supplier ID"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Supplier Name"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmsupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstSupplier As New ADODB.Recordset
Dim rstSave As New ADODB.Recordset
Dim rstKey As New ADODB.Recordset
Dim rstCompany As New ADODB.Recordset
Dim rstsearch As New ADODB.Recordset
Dim rstcheck  As New ADODB.Recordset
Dim addb As Boolean
Private Sub clearall()
txtSID.Text = ""
txtSName.Text = ""
txtStreet.Text = ""
txtColony.Text = ""
txtCity.Text = ""
txtContact.Text = ""
txtEmail.Text = ""
End Sub
Private Sub bordertext(i As Integer)
    If i = 1 Then
        txtSName.BorderStyle = vbFixedSingle
        txtSName.BackColor = &H80000009
        txtSName.Enabled = True
        
        ComboComp.BackColor = &H80000009
        ComboComp.Enabled = True
        
        
        txtStreet.BorderStyle = vbFixedSingle
        txtStreet.BackColor = &H80000009
        txtStreet.Enabled = True
        
        txtColony.BorderStyle = vbFixedSingle
        txtColony.BackColor = &H80000009
        txtColony.Enabled = True
        
        txtCity.BorderStyle = vbFixedSingle
        txtCity.BackColor = &H80000009
        txtCity.Enabled = True
        
        txtEmail.BorderStyle = vbFixedSingle
        txtEmail.BackColor = &H80000009
        txtEmail.Enabled = True
        txtContact.BorderStyle = vbFixedSingle
        txtContact.BackColor = &H80000009
        txtContact.Enabled = True
    Else
        txtSName.BorderStyle = 0
        txtSName.BackColor = &H8000000F
        txtSName.Enabled = False
        
        ComboComp.BackColor = &H8000000F
        ComboComp.Enabled = False
        
        txtStreet.BorderStyle = 0
        txtStreet.BackColor = &H8000000F
        txtStreet.Enabled = False
        
        txtColony.BorderStyle = 0
        txtColony.BackColor = &H8000000F
        txtColony.Enabled = False
        
        txtCity.BorderStyle = 0
        txtCity.BackColor = &H8000000F
        txtCity.Enabled = False
        
        txtEmail.BorderStyle = 0
        txtEmail.BackColor = &H8000000F
        txtEmail.Enabled = False
        txtContact.BorderStyle = 0
        txtContact.BackColor = &H8000000F
        txtContact.Enabled = False
    End If
End Sub

Private Sub cmdadd_Click()
bordertext (1)
clearall
If rstKey.State = adStateOpen Then rstKey.Close
rstKey.Open "select supplier_id from key_id", cnn, adOpenKeyset, adLockOptimistic
txtSID.Text = rstKey.Fields("supplier_id").Value + 1
If pk = 1000 Then
    MsgBox ("Company Limit Is 999 ")
    Frame1.Enabled = False
    Confirmation.Enabled = False
    Manupilation.Enabled = True
    showData
    Exit Sub
End If
combocompanyName

addb = True
Manupilation.Enabled = False
Confirmation.Enabled = True
MSFlexGrid1.Enabled = False
Frame1.Enabled = True
txtSName.SetFocus
End Sub

Private Sub cmdCancel_Click()
Confirmation.Enabled = False
Manupilation.Enabled = True
MSFlexGrid1.Enabled = True
Frame1.Enabled = False
showData
bordertext (0)
addb = True
End Sub

Private Sub cmdedit_Click()
MSFlexGrid1.Enabled = False
If txtSName.Text = "" Then
    MsgBox "Select Option from the grid"
    MSFlexGrid1.Enabled = True
    Exit Sub
'    Manupilation.Enabled = False
'    Confirmation.Enabled = True
'
Else
    bordertext (1)
    txtSName.SetFocus
    MSFlexGrid1_Click
    combocompanyName
    Manupilation.Enabled = False
    Confirmation.Enabled = True
    addb = False
End If
End Sub
Private Sub saveData()
If addb = False Then
   If rstSave.Fields(1).Value = txtSName.Text And rstSave.Fields(2).Value = rstCompany.Fields(0).Value And rstSave.Fields(3).Value = txtStreet.Text And rstSave.Fields(4).Value = txtColony.Text And rstSave.Fields(5).Value = txtCity.Text And rstSave.Fields(6).Value = txtContact.Text And rstSave.Fields(7).Value = txtEmail.Text And rstSave.Fields(2).Value = ComboComp.Text Then
    MsgBox "No changes made"
    Exit Sub
    End If
End If
rstSave.Fields(1).Value = StrConv(Trim(txtSName.Text), vbProperCase)
rstSave.Fields(2).Value = StrConv(Trim(rstCompany.Fields(0).Value), vbProperCase)
rstSave.Fields(3).Value = StrConv(Trim(txtStreet.Text), vbProperCase)
rstSave.Fields(4).Value = StrConv(Trim(txtColony.Text), vbProperCase)
rstSave.Fields(5).Value = StrConv(Trim(txtCity.Text), vbProperCase)
rstSave.Fields(6).Value = txtContact.Text
rstSave.Fields(7).Value = StrConv(Trim(txtEmail.Text), vbLowerCase)

End Sub

Private Sub cmdSave_Click()
'supplier'

If Not Len(txtSName.Text) >= 2 Then
    MsgBox ("Name must have atleast 2 character"), vbOKOnly + vbCritical, "Medical Store Automation"
    txtSName.SetFocus
    Exit Sub
End If
If Not Len(txtStreet.Text) >= 1 Then
    MsgBox ("Street must have atleast 1 character"), vbOKOnly + vbCritical, "Medical Store Automation"
    txtStreet.SetFocus
    Exit Sub
ElseIf Len(txtStreet.Text) = 1 Then
        If Not (Asc(txtStreet.Text) >= 48 And Asc(txtStreet.Text) <= 57) Then
            MsgBox ("Invalid Street No/H.No"), vbOKOnly + vbCritical, "Medical Store Automation"
            txtStreet.SetFocus
            Exit Sub
        End If
        
'''''''''''''8888888
Dim c%, cn%, k%
ElseIf Len(txtStreet.Text) > 2 Then
        For k = 1 To Len(txtStreet.Text)
                If IsNumeric(Right(Left(txtStreet.Text, k), 1)) Then
                    c = c + 1
                End If
                If (Asc(Right(Left(txtStreet.Text, k), 1)) >= 97 And Asc(Right(Left(txtStreet.Text, k), 1)) <= 122) Or (Asc(Right(Left(txtStreet.Text, k), 1)) >= 65 And Asc(Right(Left(txtStreet.Text, k), 1)) <= 90) Then
                    cn = cn + 1
                End If
        Next
    If c = Len(txtStreet.Text) Or cn = Len(txtStreet.Text) Then
        MsgBox ("Invalid Street No/H.No"), vbOKOnly + vbCritical, "Medical Store Automation"
            txtStreet.SetFocus
            Exit Sub
    End If
    

End If


If Not Len(txtColony.Text) >= 4 Then
    MsgBox ("Colony must have atleast 4 character"), vbOKOnly + vbCritical, "Medical Store Automation"
    txtColony.SetFocus
    Exit Sub
End If
If Not Len(txtCity.Text) >= 4 Then
    MsgBox ("City must have atleast 4 character"), vbOKOnly + vbCritical, "Medical Store Automation"
    txtCity.SetFocus
    Exit Sub
End If



If rstcheck.State = adStateOpen Then rstcheck.Close

rstcheck.Open "select * from querysup where SupplierName='" & txtSName.Text & "' and CompanyName='" & ComboComp.Text & "'and StreetNo='" & txtStreet.Text & "' and Colony='" & txtColony.Text & "'and city='" & txtCity.Text & "' and Contact='" & txtContact.Text & "' and EmailId='" & txtEmail.Text & "';", cnn, adOpenKeyset, adLockOptimistic
If rstcheck.RecordCount > 0 Then
        MsgBox "Supplier of this company already exist", vbCritical, "Medical Store Automation"
        txtSName.SetFocus
        Exit Sub
End If



Dim cc As Integer
If Len(txtContact.Text) < 1 Then
MsgBox ("Invalid number"), vbOKOnly + vbCritical, "Medical Store Automation"
    txtContact.SetFocus
    Exit Sub
'''
End If
cc = Left(txtContact.Text, 1)
If rstsearch.State = adStateOpen Then rstsearch.Close
rstsearch.Open "select * from supplier where contact='" & txtContact.Text & "'", cnn, adOpenKeyset, adLockOptimistic
If Not Len(txtContact.Text) = 10 Then
    MsgBox ("Contact should have 10 numbers"), vbOKOnly + vbCritical, "Medical Store Automation"
    txtContact.SetFocus
    Exit Sub

ElseIf Not (cc = 7 Or cc = 8 Or cc = 9) Then
        MsgBox ("start with 7 ,8,9 ")
        txtContact.SetFocus
    Exit Sub
ElseIf rstsearch.RecordCount > 0 And addb = True Then

        MsgBox ("Contact number already exist"), vbOKOnly + vbCritical, "Medical Store Automation"
        txtContact.SetFocus
    Exit Sub
ElseIf Len(txtContact.Text) = 10 Then
    Dim j%
  '  MsgBox (InStr(1, txtContact.Text, "9"))
    For i = 1 To 9
'    MsgBox Left(txtContact.Text, 1)
'    MsgBox Left(txtContact.Text, 1 + i)
        If Left(txtContact.Text, 1) = Right(Left(txtContact.Text, i + 1), 1) Then
            j = j + 1
        End If
    Next
    If j = 9 Then
    MsgBox ("Invalid contact number"), vbOKOnly + vbCritical, "Medical Store Automation"
    txtContact.SetFocus
    Exit Sub
    End If
End If


If Not Len(txtEmail.Text) >= 11 Then
    MsgBox ("Email must have atleast 11 character"), vbOKOnly + vbCritical, "Medical Store Automation"
    txtEmail.SetFocus
    Exit Sub
End If

If rstsearch.State = adStateOpen Then rstsearch.Close
rstsearch.Open "select * from supplier where emailid='" & txtEmail.Text & "'", cnn, adOpenKeyset, adLockOptimistic
'''''''''
If ((InStr(InStr(1, txtEmail.Text, "@"), txtEmail.Text, ".") - InStr(1, txtEmail.Text, "@")) < 3) Then
MsgBox ("Must have atleast 2 charector between @ and . "), vbOKOnly + vbCritical, "Medical Store Automation"
    txtEmail.SetFocus
    MsgBox InStr(1, txtEmail.Text, ".")
    MsgBox InStr(1, txtEmail.Text, "@")
    Exit Sub
End If
Dim cmp As String, em As String
em = txtEmail.Text
em = Right(em, Len(em) - InStr(em, "@"))
em = Right(em, Len(em) - InStr(em, "."))

cmp = Right(txtEmail.Text, 4)
If Not ((em = "co.in") Or (em = "com") Or (em = "edu") Or (em = "in") Or (em = "org")) Then
    MsgBox ("Email must end like .com /.in/.org/.edu/.co.in"), vbOKOnly + vbCritical, "Medical Store Automation"
    txtEmail.SetFocus
    Exit Sub


ElseIf rstsearch.RecordCount > 0 And addb = True Then

    MsgBox ("email already exist"), vbOKOnly + vbCritical, "Medical Store Automation"
    txtEmail.SetFocus
    Exit Sub
End If

If InStr(1, txtEmail.Text, "@") = 0 Then

    MsgBox ("atleast one @ is needed"), vbOKOnly + vbCritical, "Medical Store Automation"
    txtEmail.SetFocus
    Exit Sub
End If

If (InStr(InStr(1, txtEmail.Text, "@") + 1, txtEmail.Text, "@")) Then

    MsgBox ("only one @ allowed"), vbOKOnly + vbCritical, "Medical Store Automation"
    txtEmail.SetFocus
    Exit Sub
End If
''''''


Dim cmpt As Integer
cmpt = Asc(Left(txtEmail.Text, 1))
If Not (cmpt >= 97 And cmpt <= 122) Or (cmpt >= 65 And cmpt <= 90) Then
    MsgBox ("must start with charector")
    txtEmail.SetFocus
    Exit Sub
End If



If rstSave.State = adStateOpen Then rstSave.Close
If rstCompany.State = adStateOpen Then rstCompany.Close
rstSave.Open "select * from supplier", cnn, adOpenKeyset, adLockOptimistic
rstCompany.Open "select * from company where companyname='" & ComboComp.Text & "'", cnn, adOpenKeyset, adLockOptimistic
If addb = True Then
rstSave.AddNew
rstSave.Fields(0).Value = txtSID.Text
'save
saveData
rstSave.Update
'key_id
If rstKey.State = adStateOpen Then rstKey.Close
    rstKey.Open "select supplier_id from key_id ", cnn, adOpenKeyset, adLockOptimistic
rstKey.MoveFirst
rstKey.Fields("supplier_id").Value = rstKey.Fields("supplier_id").Value + 1
rstKey.Update
Else
    If rstSave.State = adStateOpen Then rstSave.Close
        rstSave.Open "select * from supplier where supplierid='" & txtSID.Text & "'", cnn, adOpenKeyset, adLockOptimistic
        saveData
    rstSave.Update
    End If
'End If
flexData
Manupilation.Enabled = True
Confirmation.Enabled = False

bordertext (0)
MsgBox "Data Saved", vbOKOnly + vbInformation, "Medical Store Automation"
MSFlexGrid1.Enabled = True
End Sub

Private Sub showData()
If rstSupplier.State = adStateOpen Then rstSupplier.Close
rstSupplier.Open "select * from querysup", cnn, adOpenKeyset, adLockOptimistic
txtSID.Text = rstSupplier.Fields(0).Value
txtSName.Text = rstSupplier.Fields(1).Value
ComboComp.Clear
ComboComp.additem rstSupplier.Fields(2).Value
ComboComp.ListIndex = 0
txtStreet.Text = rstSupplier.Fields(3).Value
txtColony.Text = rstSupplier.Fields(4).Value
txtCity.Text = rstSupplier.Fields(5).Value
txtContact.Text = rstSupplier.Fields(6).Value
txtEmail.Text = rstSupplier.Fields(7).Value
'MsgBox MSFlexGrid1.Rows



'MSFlexGrid1.ColWidth(7) = 1500

End Sub

'Private Sub Command1_Click()
'If rstsearch.State = adStateOpen Then rstKey.Close
'rstsearch.Open "select * from supplier ", cnn, adOpenKeyset, adLockOptimistic
'rstsearch.Find "contact='" & txtContact.Text & "'", 0, adSearchForward, 1
'If rstsearch.EOF = True Then
'    MsgBox "no"
'End If
'End Sub

Private Sub Form_Load()
txtSID.BorderStyle = 0
txtSID.BackColor = &H8000000F
txtSID.Enabled = False
'key_id
If rstKey.State = adStateOpen Then rstKey.Close
rstKey.Open "Select supplier_id from key_id", cnn, adOpenKeyset, adLockOptimistic
txtSID.Text = rstKey.Fields("supplier_id").Value
showData
flexData
bordertext (0)
Manupilation.Enabled = True
Confirmation.Enabled = False
End Sub
Private Sub combocompanyName()
ComboComp.Clear
If rstCompany.State = adStateOpen Then rstCompany.Close
        rstCompany.Open "select companyname from company", cnn, adOpenKeyset, adLockOptimistic
        rstCompany.MoveFirst
    For i = 0 To rstCompany.RecordCount - 1
        ComboComp.additem rstCompany.Fields(0).Value
        rstCompany.MoveNext
    Next
    ComboComp.ListIndex = 0

End Sub
Private Sub flexData()
 Dim X As Integer, Y As Integer
    If rstSupplier.State = adStateOpen Then rstSupplier.Close

   MSFlexGrid1.Visible = True
    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 8

    rstSupplier.Open "select * from querysup", cnn, adOpenKeyset, adLockOptimistic

    MSFlexGrid1.Cols = rstSupplier.Fields.Count
    
    
    For X = 0 To rstSupplier.Fields.Count - 1
        MSFlexGrid1.TextMatrix(0, X) = rstSupplier.Fields(X).Name
'        MSFlexGrid1.ColWidth(X) = 1000

    Next
     MSFlexGrid1.ColWidth(7) = 1800

    MSFlexGrid1.Height = 3000
'    MSFlexGrid1.Width = 9000

    For X = 1 To rstSupplier.RecordCount
        For Y = 0 To rstSupplier.Fields.Count - 1
            MSFlexGrid1.TextMatrix(X, Y) = rstSupplier.Fields(Y).Value
        Next
        rstSupplier.MoveNext
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        Next
        combocompanyName
End Sub

Private Sub MSFlexGrid1_Click()
 txtSID.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtSName.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)

If rstCompany.State = adStateOpen Then rstCompany.Close

ComboComp.Clear
rstCompany.Open "select * from company where companyname='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) & "'", cnn, adOpenKeyset, adLockOptimistic
ComboComp.additem rstCompany.Fields(1).Value
ComboComp.ListIndex = 0
 txtStreet.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
 txtColony.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
txtCity.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
 txtContact.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
 txtEmail.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)
If addb = True Then
bordertext (0)
Manupilation.Enabled = True
Confirmation.Enabled = False
End If
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)

'        KeyAscii = key(KeyAscii, txtCity.Text)

If Len(txtCity.Text) = 0 Or txtCity.SelStart = 0 Then
    KeyAscii = key(KeyAscii, txtCity.Text)
Else

    If KeyAscii = 8 Or KeyAscii = 32 Then
    
    Else
        KeyAscii = key(KeyAscii, txtCity.Text)
    End If
    If Len(txtCity.Text) > 1 And KeyAscii = 32 Then
        If Asc(Right((Left(txtCity.Text, txtCity.SelStart)), 1)) = KeyAscii Or Asc(Right((Left(txtCity.Text, txtCity.SelStart + 1)), 1)) = KeyAscii Then
        KeyAscii = 0
        End If
    End If
End If

End Sub

Private Sub txtColony_KeyPress(KeyAscii As Integer)
If Len(txtColony.Text) = 0 Or txtColony.SelStart = 0 Then
    KeyAscii = key(KeyAscii, txtColony.Text)
Else

    If KeyAscii = 8 Or KeyAscii = 32 Then
    
    Else
        KeyAscii = key(KeyAscii, txtColony.Text)
    End If
    If Len(txtColony.Text) > 1 And KeyAscii = 32 Then
        If Asc(Right((Left(txtColony.Text, txtColony.SelStart)), 1)) = KeyAscii Or Asc(Right((Left(txtColony.Text, txtColony.SelStart + 1)), 1)) = KeyAscii Then
        KeyAscii = 0
        End If
    End If
End If
'KeyAscii = key(KeyAscii, txtColony.Text)
End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)
'If Len(txtcontact.Text) >= 10 Then
'    KeyAscii = 0
'End If
'If Len(txtcontact.Text) = 1 Then
'    If Not (txtcontact.Text = 8 Or txtcontact.Text = 9 Or txtcontact.Text = 7 Or KeyAscii = 8) Then
'        KeyAscii = 0
'End If
'End If
If Len(txtContact.Text) = 0 Then
    If Not (KeyAscii = 57 Or KeyAscii = 55 Or KeyAscii = 56) Then
        KeyAscii = 0
    End If
Else
        KeyAscii = keynum(KeyAscii, txtContact.Text)
End If


End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
If Len(txtEmail.Text) = 0 Or txtEmail.SelStart = 0 Then
    If KeyAscii = 32 Then
        KeyAscii = 0
    End If
    KeyAscii = key(KeyAscii, txtEmail.Text)
Else
'''
    If KeyAscii = 8 Or KeyAscii = 64 Or KeyAscii = 95 Or KeyAscii = 46 Then
    
    Else
        KeyAscii = keyboth(KeyAscii, txtEmail.Text)
    End If
    If Len(txtEmail.Text) > 1 And (KeyAscii = 64 Or KeyAscii = 95 Or KeyAscii = 46) Then
    Dim aa1%, aa2%
    aa1 = Asc(Right((Left(txtEmail.Text, txtEmail.SelStart)), 1))
    aa2 = Asc(Right((Left(txtEmail.Text, txtEmail.SelStart + 1)), 1))
        If aa1 = KeyAscii Or aa2 = KeyAscii Then
        KeyAscii = 0
        End If
        If ((aa1 = 64 Or aa1 = 95 Or aa1 = 46) And (KeyAscii = 64 Or KeyAscii = 95 Or KeyAscii = 46)) Then
                KeyAscii = 0
        End If
        If ((aa2 = 64 Or aa2 = 95 Or aa2 = 46) And (KeyAscii = 64 Or KeyAscii = 95 Or KeyAscii = 46)) Then
                KeyAscii = 0
        End If
    End If
End If
End Sub

Private Sub txtSName_KeyPress(KeyAscii As Integer)
If Len(txtSName.Text) = 0 Or txtSName.SelStart = 0 Then
    KeyAscii = key(KeyAscii, txtSName.Text)
Else

    If KeyAscii = 8 Or KeyAscii = 32 Then
    
    Else
        KeyAscii = key(KeyAscii, txtSName.Text)
    End If
    If Len(txtSName.Text) > 1 And KeyAscii = 32 Then
        If Asc(Right((Left(txtSName.Text, txtSName.SelStart)), 1)) = KeyAscii Or Asc(Right((Left(txtSName.Text, txtSName.SelStart + 1)), 1)) = KeyAscii Then
        KeyAscii = 0
        End If
    End If
End If
'        KeyAscii = key(KeyAscii, txtSName.Text)
''If Not (KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 120 Or KeyAscii = 8 Or KeyAscii = 32) Or Len(txtSName.Text) > 2 Then
''    KeyAscii = 0
''End If
End Sub

Private Sub txtStreet_KeyPress(KeyAscii As Integer)
'If (KeyAscii = 47 Or KeyAscii = 46 Or KeyAscii = 41 Or KeyAscii = 45 Or KeyAscii = 40 Or KeyAscii = 8) Then
'
'Else
'        KeyAscii = keyboth(KeyAscii, txtStreet.Text)
'End If
If Len(txtStreet.Text) = 0 Or txtStreet.SelStart = 0 Then
    If KeyAscii = 32 Then
        KeyAscii = 0
    End If
    KeyAscii = keyboth(KeyAscii, txtStreet.Text)
Else
'''
   ' If KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 32 Then
    If (KeyAscii = 47 Or KeyAscii = 46 Or KeyAscii = 41 Or KeyAscii = 45 Or KeyAscii = 40 Or KeyAscii = 8) Then
    
    Else
        KeyAscii = keyboth(KeyAscii, txtStreet.Text)
    End If
    If Len(txtStreet.Text) > 1 And (KeyAscii = 47 Or KeyAscii = 46 Or KeyAscii = 41 Or KeyAscii = 45 Or KeyAscii = 40) Then
        If Asc(Right((Left(txtStreet.Text, txtStreet.SelStart)), 1)) = KeyAscii Or Asc(Right((Left(txtStreet.Text, txtStreet.SelStart + 1)), 1)) = KeyAscii Then
        KeyAscii = 0
        End If
    End If
    ''''
    ''''
    If Len(txtEmail.Text) > 1 And (KeyAscii = 47 Or KeyAscii = 46 Or KeyAscii = 41 Or KeyAscii = 45 Or KeyAscii = 40) Then
    Dim aa1%, aa2%
    aa1 = Asc(Right((Left(txtStreet.Text, txtStreet.SelStart)), 1))
    aa2 = Asc(Right((Left(txtStreet.Text, txtStreet.SelStart + 1)), 1))
        If aa1 = KeyAscii Or aa2 = KeyAscii Then
        KeyAscii = 0
        End If
        If ((aa1 = 47 Or aa1 = 41 Or aa1 = 46 Or aa1 = 45 Or aa1 = 40) And (KeyAscii = 47 Or KeyAscii = 46 Or KeyAscii = 41 Or KeyAscii = 45 Or KeyAscii = 40)) Then
                KeyAscii = 0
        End If
        If ((aa2 = 47 Or aa2 = 41 Or aa2 = 46 Or aa1 = 45 Or aa1 = 40) And (KeyAscii = 47 Or KeyAscii = 46 Or KeyAscii = 41 Or KeyAscii = 45 Or KeyAscii = 40)) Then
                KeyAscii = 0
        End If
    End If
    ''''
    ''''
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim a As Integer
a = MsgBox("Do you want to Exit ?", vbYesNo + vbDefaultButton2 + vbCritical, "Medical Store Automation")
If a = 6 Then
    Unload Me
Else
    Cancel = True
End If

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Confirmation.Enabled = True Then
    MsgBox "Please Complete the session", vbCritical, "Medical Store Automation"
    Cancel = True
End If
End Sub
