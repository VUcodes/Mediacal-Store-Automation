VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmcategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Category"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmsave 
      Caption         =   "Confirmation"
      Height          =   1815
      Left            =   3360
      TabIndex        =   11
      Top             =   2040
      Width           =   1215
      Begin VB.CommandButton cmdcanle 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame frmedit 
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtc_id 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtc_name 
         Height          =   405
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblc_id 
         Caption         =   "Category ID"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblc_name 
         Caption         =   "Category Nmae"
         Height          =   255
         Left            =   1560
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame frmadd 
      Caption         =   "Manipulation"
      Height          =   1815
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   6800
      _Version        =   393216
      BackColorBkg    =   -2147483633
   End
End
Attribute VB_Name = "frmcategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstAdd As New ADODB.Recordset ' recordset to add key record in category table
Dim rstcategory As New ADODB.Recordset ' record set for categery table
Dim rstaddnew As New ADODB.Recordset ' recoedset to add new record in category table
Dim rstEdit As New ADODB.Recordset ' recordset to edit record in category table
Dim pk As Integer
Dim i As Integer, j As Integer
Dim ca As String
Dim addb As Boolean 'variable to check the button add or edit

Private Sub bStyle(b As Boolean)
If b = True Then
    'txtc_id.BorderStyle = vbFixedSingle
    'txtc_id.BackColor = &H80000005
    '
    txtc_name.BorderStyle = vbFixedSingle
    txtc_name.BackColor = &H80000005
    '
    
Else

    txtc_id.BorderStyle = 0
    txtc_id.BackColor = &H8000000F
    '
    txtc_name.BorderStyle = 0
    txtc_name.BackColor = &H8000000F
End If    '
End Sub

Private Sub cmdadd_Click()
MSFlexGrid1.Enabled = False
frmedit.Enabled = True
frmSave.Enabled = True
frmadd.Enabled = False
bStyle (True)
additem

End Sub

Private Sub cmdcanle_Click()
frmedit.Enabled = False
frmSave.Enabled = False
frmadd.Enabled = True
MSFlexGrid1.Enabled = True
cmdadd.SetFocus
showData
bStyle (False)
End Sub

Private Sub cmdedit_Click()
MSFlexGrid1.Enabled = False
frmedit.Enabled = True
frmSave.Enabled = True
frmadd.Enabled = False
txtc_name.SetFocus
selected
addb = False
bStyle (True)
End Sub
Private Sub selected()
    txtc_name.SelStart = 0
    txtc_name.SelLength = Len(txtc_name.Text)
End Sub
Private Sub cmdSave_Click()

If rstAdd.State = adStateOpen Then rstAdd.Close
If rstaddnew.State = adStateOpen Then rstaddnew.Close
rstAdd.Open "select * from category", cnn, adOpenKeyset, adLockOptimistic
rstaddnew.Open "select * from key_id", cnn, adOpenKeyset, adLockOptimistic

If Len(txtc_name.Text) < 4 Then
    MsgBox "Category Name Should Be Greater Than 3 Character", vbExclamation
    txtc_name.SetFocus
    selected
    Exit Sub
End If

j = 0

ca = Left(txtc_name.Text, 1)
For i = 1 To Len(txtc_name.Text)
   If (ca = Right(Left(txtc_name.Text, i), 1)) Then
    j = j + 1
   End If
Next
If j = Len(txtc_name.Text) Then
    MsgBox ("Invalid Name"), vbInformation, "Medical Store Automation"
    txtc_name.SetFocus
    selected
    Exit Sub
End If

If rstAdd.RecordCount > 0 Then
    
    
'    If txtc_name.Text = "" Then
'            MsgBox "Category cannot be  empty", vbInformation, "Medical Store Automation"
'            txtc_name.SetFocus
'            selected
'            Exit Sub
'    End If
    
    While rstAdd.EOF <> True
    
        If StrConv(Trim(txtc_name.Text), vbProperCase) = rstAdd.Fields(1).Value Then
            MsgBox "Category Name Already Exist", vbInformation, "Medical Store Automation"
            txtc_name.SetFocus
            selected
            Exit Sub
        End If
        rstAdd.MoveNext
    Wend
End If
If addb = True Then
    
    rstaddnew.MoveFirst
    rstaddnew.Fields(0).Value = rstaddnew.Fields(0).Value + 1
    rstaddnew.Update
    
    rstAdd.AddNew
    rstAdd.Fields(0).Value = Trim(Str(pk))
    rstAdd.Fields(1).Value = StrConv(Trim(txtc_name.Text), vbProperCase)
    rstAdd.Update
Else
    If rstEdit.State = adStateOpen Then rstEdit.Close
    rstEdit.Open "select * from category where categoryid='" & txtc_id.Text & "'", cnn, adOpenKeyset, adLockOptimistic
    rstEdit.Fields(1).Value = StrConv(Trim(txtc_name.Text), vbProperCase)
    rstEdit.Update
End If
cmdcanle_Click
showData
MsgBox "Data Saved", vbOKOnly + vbInformation, "Medical Store Automation"
End Sub

Private Sub Form_Activate()
cmdadd.SetFocus
End Sub
Private Sub Form_Load()
frmedit.Enabled = False
frmSave.Enabled = False
showData
addb = True
bStyle (False)
End Sub

Private Sub additem()
addb = True
clear_all
txtc_name.SetFocus
If rstaddnew.State = adStateOpen Then rstaddnew.Close
rstaddnew.Open "select * from key_id", cnn, adOpenKeyset, adLockOptimistic
pk = rstaddnew.Fields(0).Value + 1
If pk = 100 Then
    MsgBox ("Category limit is 99 ")
    frmedit.Enabled = False
    frmSave.Enabled = False
    frmadd.Enabled = True
    showData
    Exit Sub
End If
If pk > 0 And pk < 10 Then
txtc_id.Text = pk
Else
txtc_id.Text = pk
End If

End Sub
Private Sub showData()

If rstcategory.State = adStateOpen Then rstcategory.Close
    Dim X As Integer, Y As Integer
    MSFlexGrid1.Visible = True
    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 2
    
    rstcategory.Open "select * from Category", cnn, adOpenKeyset, adLockOptimistic
    
    txtc_id.Text = rstcategory.Fields(0).Value
    txtc_name.Text = rstcategory.Fields(1).Value
    
    MSFlexGrid1.Cols = rstcategory.Fields.Count
        For X = 0 To rstcategory.Fields.Count - 1
            MSFlexGrid1.TextMatrix(0, X) = rstcategory.Fields(X).Name
            MSFlexGrid1.ColWidth(X) = 1400
'
        Next
'
'            'MSFlexGrid1.CellHeight(1, 1) = 1500
'            'MSFlexGrid1.CellWidth(1, 1) = 500
    
'    MSFlexGrid1.Height = 5000
'    MSFlexGrid1.Width = 3400
    For X = 1 To rstcategory.RecordCount
        For Y = 0 To rstcategory.Fields.Count - 1
            MSFlexGrid1.TextMatrix(X, Y) = rstcategory.Fields(Y).Value
'            MSFlexGrid1.RowHeight(X) = 500
        Next
        rstcategory.MoveNext
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    Next
End Sub
Private Sub clear_all()
txtc_id.Text = ""
txtc_name.Text = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If frmedit.Enabled = True Then
    MsgBox "Please Complete the Session", vbCritical, "Medical Store Automation"
    Cancel = True
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

Private Sub MSFlexGrid1_Click()
    txtc_id.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
    txtc_name.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
End Sub


Private Sub txtc_name_KeyPress(KeyAscii As Integer)
'If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 8) Then
'    KeyAscii = 0
'End If
''    keyascii = validation1
'If ((Len(txtc_name.Text) = 0) And KeyAscii = 32) Or (Right(txtc_name, 1) = " " And KeyAscii = 32) Then
'    KeyAscii = 0
'End If
If Len(txtc_name.Text) = 0 Or txtc_name.SelStart = 0 Then
    If KeyAscii = 32 Then
        KeyAscii = 0
    End If
    KeyAscii = key(KeyAscii, txtc_name.Text)
Else

    If KeyAscii = 8 Or KeyAscii = 32 Then
    
    Else
        KeyAscii = key(KeyAscii, txtc_name.Text)
    End If
    
    If Len(txtc_name.Text) > 1 And KeyAscii = 32 Then
        If Asc(Right((Left(txtc_name.Text, txtc_name.SelStart)), 1)) = KeyAscii Or Asc(Right((Left(txtc_name.Text, txtc_name.SelStart + 1)), 1)) = KeyAscii Then
        KeyAscii = 0
        End If
    End If
    
End If
End Sub
