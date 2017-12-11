VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmmedicine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medicine"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11625
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11625
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3975
      Left            =   5280
      TabIndex        =   26
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7011
      _Version        =   393216
   End
   Begin VB.Frame fraManupilation 
      Caption         =   "Manupilation"
      Height          =   1695
      Left            =   3840
      TabIndex        =   25
      Top             =   240
      Width           =   1095
      Begin VB.CommandButton cmdedit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraConfirmation 
      Caption         =   "Confirmation"
      Height          =   1695
      Left            =   3840
      TabIndex        =   24
      Top             =   2040
      Width           =   1095
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraunit 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   6375
      Begin VB.CommandButton btnuremove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton btnuadd 
         Caption         =   "Add"
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2055
         Left            =   3840
         TabIndex        =   17
         Top             =   120
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   3625
         _Version        =   393216
      End
      Begin VB.TextBox txtrl 
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtweight 
         Height          =   375
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Reorder Level"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Weight"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame framedicin 
      Caption         =   "Medicine Details"
      Height          =   4695
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   11655
      Begin VB.TextBox txtbestb 
         Height          =   375
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   6
         Top             =   3960
         Width           =   1815
      End
      Begin VB.ComboBox cmbcomn 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ComboBox cmbcatn 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2760
         Width           =   1815
      End
      Begin VB.ComboBox cmbuname 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox txtmname 
         Height          =   375
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtmid 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblbestb 
         Caption         =   "Best Before             (In Months)"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Unit Name"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Category Name"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Company Name"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Medicine Name"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Medicine Id"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmmedicine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstmedicine As New ADODB.Recordset
Dim rstunit As New ADODB.Recordset
Dim rstcomb As New ADODB.Recordset
Dim rstmsave As New ADODB.Recordset
Dim rstC As New ADODB.Recordset
Dim rstaddnew As New ADODB.Recordset
Dim rstrsave As New ADODB.Recordset
Dim rstcheck As New ADODB.Recordset
Dim cc%, a As Boolean, pk%

Dim rstKey As New ADODB.Recordset
Private Sub bStyle(b As Boolean)
If b = True Then
    txtmID.BorderStyle = vbFixedSingle
    txtmID.BackColor = &H80000005
    '
    txtmName.BorderStyle = vbFixedSingle
    txtmName.BackColor = &H80000005
    '
    txtrl.BorderStyle = vbFixedSingle
    txtrl.BackColor = &H80000005
    '
    txtweight.BorderStyle = vbFixedSingle
    txtweight.BackColor = &H80000005
    
    txtbestb.BorderStyle = vbFixedSingle
    txtbestb.BackColor = &H80000005
    
    cmbcatn.BackColor = &H80000005
    '
    cmbcomn.BackColor = &H80000005
    '
    cmbuname.BackColor = &H80000005
    
Else

    txtmID.BorderStyle = 0
    txtmID.BackColor = &H8000000F
    '
    cmbcatn.BackColor = &H8000000F
    '
    cmbcomn.BackColor = &H8000000F
    '
    cmbuname.BackColor = &H8000000F
    '
    txtmName.BorderStyle = 0
    txtmName.BackColor = &H8000000F
    '
    txtrl.BorderStyle = 0
    txtrl.BackColor = &H8000000F
    '
    txtweight.BorderStyle = 0
    txtweight.BackColor = &H8000000F
    
    txtbestb.BorderStyle = 0
    txtbestb.BackColor = &H8000000F
End If

End Sub
Private Sub btnuadd_Click()
If txtweight.Text = "" Then
    MsgBox ("Please Enter Weight"), vbCritical, "Medical Store Automation"
    txtweight.SetFocus
    Exit Sub
End If
If txtrl.Text = "" Then
    MsgBox ("Please Enter reorderlevel"), vbCritical, "Medical Store Automation"
    txtrl.SetFocus
    Exit Sub
End If

uadd
End Sub

Private Sub btnuremove_Click()
uremove
End Sub

Private Sub combo()
'Dim rstC As New ADODB.Recordset
Dim i As Integer
'If rstC.State = adStateOpen Then rstC.Close
'rstC.Open "select companyname from company", cnn, adOpenKeyset, adLockOptimistic
'For i = 0 To rstC.RecordCount - 1
'    cmbcomn.additem rstC.Fields(0).Value
'    rstC.MoveNext
'Next
cmbuname.Clear
cmbcatn.Clear
If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select * from companycat where CompanyName='" & cmbcomn.Text & "'", cnn, adOpenKeyset, adLockOptimistic
For i = 0 To rstC.RecordCount - 1
    cmbcatn.additem rstC.Fields(2).Value
    
    If rstC.Fields(2).Value = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) Then
        cmbcatn.ListIndex = i
    End If
    rstC.MoveNext
    
Next


If cmbcatn.ListCount > 0 And cmbcatn.Text = "" Then
cmbcatn.ListIndex = 0
End If

'Dim rstC As New ADODB.Recordset
'Dim i As Integer
'If rstC.State = adStateOpen Then rstC.Close
'rstC.Open "select companyname from company", cnn, adOpenKeyset, adLockOptimistic
'While (rstC.EOF = False)
'cmbcomn.additem rstC.Fields(0).Value
'rstC.MoveNext
'Wend
'If rstC.State = adStateOpen Then rstC.Close
'rstC.Open "select categoryname from category", cnn, adOpenKeyset, adLockOptimistic
'While (rstC.EOF = False)
'cmbcatn.additem rstC.Fields(0).Value
'rstC.MoveNext
'Wend
'cmbuname.Clear
'If a = True Then
If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select measure_name from measure", cnn, adOpenKeyset, adLockOptimistic
While (rstC.EOF = False)
 cmbuname.additem rstC.Fields(0).Value
    rstC.MoveNext
Wend
cmbuname.ListIndex = 0
'End If
'cmbcomn.ListIndex = 0
'cmbcatn.ListIndex = 0
'cmbuname.ListIndex = 0



End Sub
Private Sub combo1()

Dim i As Integer
cmbcatn.Clear
If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select * from companycat where CompanyName='" & cmbcatn.Text & "'", cnn, adOpenKeyset, adLockOptimistic
For i = 0 To rstC.RecordCount - 1
    cmbuname.additem rstC.Fields(2).Value
    rstC.MoveNext
Next


cmbcatn.ListIndex = 0
End Sub

Private Sub cmbcatn_Click()
'combo1
End Sub

Private Sub cmbcomn_Click()
If txtmName.Enabled = True Then
combo
End If
End Sub

Private Sub cmdadd_Click()

framedicin.Enabled = True
fraunit.Enabled = True
fraConfirmation.Enabled = True
fraManupilation.Enabled = False
a = True
cc = 0
cmbcomn.Clear
cmbuname.Clear
Dim i As Integer
If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select companyname from company", cnn, adOpenKeyset, adLockOptimistic
For i = 0 To rstC.RecordCount - 1
    cmbcomn.additem rstC.Fields(0).Value
    rstC.MoveNext
Next
cmbcomn.ListIndex = 0
combo
bStyle (True)
txtmName.SetFocus

If rstaddnew.State = adStateOpen Then rstaddnew.Close
rstaddnew.Open "select * from key_id", cnn, adOpenKeyset, adLockOptimistic
pk = rstaddnew.Fields(5).Value + 1
If pk = 10000 Then
    MsgBox ("Company Limit Is 9999 ")
    bStyle (False)
    Exit Sub
End If
'''''''
If rstunit.State = adStateOpen Then rstunit.Close
rstunit.Open "select * from medicinedetails where medicineid='" & txtmID.Text & "'", cnn, adOpenKeyset, adLockOptimistic

MSFlexGrid2.Clear
MSFlexGrid2.Cols = rstunit.Fields.Count - 2
MSFlexGrid2.Rows = 1
For i = 1 To rstunit.Fields.Count - 2
    MSFlexGrid2.TextMatrix(0, (i - 1)) = rstunit.Fields(i).Name
Next
txtbestb.Text = ""
txtmName.Text = ""
'''''''
txtmID.Text = pk
MSFlexGrid1.Enabled = False
End Sub

Private Sub cmdCancel_Click()
framedicin.Enabled = False
MSFlexGrid1.Enabled = True
fraunit.Enabled = False
fraConfirmation.Enabled = False
fraManupilation.Enabled = True
bStyle (False)
txtrl.Text = ""
txtweight.Text = ""
MSFlexGrid1.Row = 1
MSFlexGrid1_Click
txtbestb.Enabled = True
End Sub

Private Sub cmdedit_Click()
If txtmID.Text = "" Then
    Exit Sub
End If
bStyle (True)
framedicin.Enabled = True
fraConfirmation.Enabled = True
fraManupilation.Enabled = False
txtmName.SetFocus
cmbcomn.Clear
Dim i As Integer
If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select companyname from company", cnn, adOpenKeyset, adLockOptimistic
For i = 0 To rstC.RecordCount - 1
    cmbcomn.additem rstC.Fields(0).Value
    rstC.MoveNext
Next
cmbcomn.ListIndex = 0

For i = 0 To cmbcomn.ListCount - 1
    If cmbcomn.List(i) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) Then
        cmbcomn.ListIndex = i
    End If
Next
a = False
txtbestb.Enabled = False
MSFlexGrid1.Enabled = False
End Sub


Private Sub Command4_Click()
MSFlexGrid1.Enabled = True
framedicin.Enabled = False
fraunit.Enabled = False
fraConfirmation.Enabled = False
'showm
'showu
End Sub

Private Sub cmdSave_Click()

If txtmName.Text = "" Or Len(txtmName.Text) < 4 Then
    MsgBox "Medicine name should be grater than 3", vbCritical, "Medical Store Automation"
    txtmName.SetFocus
    Exit Sub
End If
Dim j%, i%, ca As String
j = 0

ca = Left(txtmName.Text, 1)
For i = 0 To Len(txtmName.Text)
   If (ca = Right(Left(txtmName.Text, i), 1)) Then
    j = j + 1
   End If
Next
If j = Len(txtmName.Text) Then
    MsgBox ("Invalid Name"), vbInformation, "Medical Store Automation"
    txtmName.SetFocus
    Exit Sub
End If

'''''''''

Dim c%, cn%, k%

        For k = 1 To Len(txtmName.Text)
                If IsNumeric(Right(Left(txtmName.Text, k), 1)) Then
                    c = c + 1
                End If
        Next
    If c = Len(txtmName.Text) Then
        MsgBox ("Invalid name"), vbOKOnly + vbCritical, "Medical Store Automation"
            txtmName.SetFocus
            Exit Sub
    End If
    
'''''''''''''''
If txtbestb.Text = "" Then
 MsgBox "Best before should be  6 to 60 ", vbCritical, "Medical Store Automation"
    txtbestb.SetFocus
    Exit Sub
End If
If Not (CInt(txtbestb.Text) >= 6 And CInt(txtbestb.Text) <= 60) Then
    MsgBox "Best before should be  6 to 60 ", vbCritical, "Medical Store Automation"
    txtbestb.SetFocus
    Exit Sub
End If
'''''''''''
'If txtweight.Text = "" Then
'    MsgBox "Weight cannot be empty or 0", vbCritical, "Medical Store Automation"
'    txtweight.SetFocus
'    Exit Sub
'End If
'If Not CInt(txtweight.Text) > 0 Then
'    MsgBox "Weight cannot be empty or 0", vbCritical, "Medical Store Automation"
'    txtweight.SetFocus
'    Exit Sub
'End If
''''''''''
'If (cmbuname.Text = "Kg" Or cmbuname.Text = "Ltr") And txtweight.Text <> "1" Then
'        MsgBox "This unit only allowed '1' in weight", vbCritical, "Medical Store Automation"
'        txtweight.SetFocus
'        Exit Sub
'End If

'If txtrl.Text = "" Then
'    MsgBox "Reorder level cannot be empty or 0", vbCritical, "Medical Store Automation"
'    txtrl.SetFocus
'    Exit Sub
'End If
'If Not CInt(txtrl.Text) > 0 Then
'    MsgBox "Reorder level cannot be empty or 0", vbCritical, "Medical Store Automation"
'    txtrl.SetFocus
'    Exit Sub
'End If

If rstmsave.State = adStateOpen Then rstmsave.Close
rstmsave.Open "select * from medicine", cnn, adOpenKeyset, adLockOptimistic
If a = True Then
    If cc = 0 Then
        MsgBox "Please enter weight and reorder level", vbCritical, "Medical Store Automation"
        txtweight.SetFocus
        Exit Sub
    End If
    
    If rstrsave.State = adStateOpen Then rstrsave.Close
    rstrsave.Open "select * from Measure where Measure_Name='" & cmbuname.Text & "'", cnn, adOpenKeyset, adLockOptimistic
    
    If Not txtmID = 1 Then
    If rstcheck.State = adStateOpen Then rstcheck.Close
    rstcheck.Open "select * from medicine where medicinename='" & StrConv(Trim(txtmName.Text), vbProperCase) & "' and compcatid='" & rstrsave.Fields(0).Value & "'", cnn, adOpenKeyset, adLockOptimistic
    If Not rstcheck.RecordCount = 0 Then
        MsgBox "Record already exist", vbCritical, "Medical Store Automation"
        txtmName.SetFocus
        Exit Sub
    End If
    End If
    rstmsave.AddNew
    'rstmsave.MoveFirst
    rstmsave.Fields(0).Value = pk
    rstmsave.Fields(1).Value = StrConv(Trim(txtmName.Text), vbProperCase)
    If rstrsave.State = adStateOpen Then rstrsave.Close
    rstrsave.Open "select * from companyCat where CompanyName='" & cmbcomn.Text & "'and CategoryName='" & cmbcatn.Text & "'", cnn, adOpenKeyset, adLockOptimistic
    rstmsave.Fields(2).Value = rstrsave.Fields(3).Value
    If rstrsave.State = adStateOpen Then rstrsave.Close
    rstrsave.Open "select * from Measure where Measure_Name='" & cmbuname.Text & "'", cnn, adOpenKeyset, adLockOptimistic
    rstmsave.Fields(3).Value = rstrsave.Fields(0).Value
    rstmsave.Fields(4).Value = StrConv(Trim(txtbestb.Text), vbProperCase)
    rstmsave.Update
    
    If rstaddnew.State = adStateOpen Then rstaddnew.Close
    rstaddnew.Open "select * from key_id", cnn, adOpenKeyset, adLockOptimistic
    rstaddnew.Fields(5).Value = txtmID.Text
    rstaddnew.Update
    
    ''''''''
    Dim pk1%
    If rstaddnew.State = adStateOpen Then rstaddnew.Close
    rstaddnew.Open "select * from key_id", cnn, adOpenKeyset, adLockOptimistic
    If rstmsave.State = adStateOpen Then rstmsave.Close
    rstmsave.Open "select * from medicinedetails", cnn, adOpenKeyset, adLockOptimistic
    pk1 = rstaddnew.Fields(6).Value
    For i = 1 To MSFlexGrid2.Rows - 1
        If pk1 + 1 = 100000 Then
            MsgBox ("Data base if full ")
            bStyle (False)
            Exit Sub
        End If
        
        rstmsave.AddNew
        rstmsave.Fields(0).Value = pk1 + i
        rstmsave.Fields(1).Value = MSFlexGrid2.TextMatrix(i, 0)
        rstmsave.Fields(2).Value = MSFlexGrid2.TextMatrix(i, 1)
        rstmsave.Fields(3).Value = txtmID.Text
        
        rstaddnew.Fields(6).Value = pk1 + i
        rstaddnew.Update
        rstmsave.Update
    Next
End If
    ''''''''
Dim cx%
If a = False Then
    If rstrsave.State = adStateOpen Then rstrsave.Close
    rstrsave.Open "select * from companyCat where CompanyName='" & cmbcomn.Text & "' and CategoryName='" & cmbcatn.Text & "'", cnn, adOpenKeyset, adLockOptimistic
    If rstcheck.State = adStateOpen Then rstcheck.Close
    rstcheck.Open "select * from medicine where medicinename='" & StrConv(Trim(txtmName.Text), vbProperCase) & "' and compcatid='" & rstrsave.Fields(3).Value & "'", cnn, adOpenKeyset, adLockOptimistic
    If Not rstcheck.RecordCount = 0 Then
        cx = cx + 1
    End If
    
End If
If a = False And cx = 1 Then
    If rstaddnew.State = adStateOpen Then rstaddnew.Close
    rstaddnew.Open "select * from qrymedicine where medicineid='" & txtmID.Text & "' and Measure_Name='" & cmbuname.Text & "' and bestbefore=" & CInt(txtbestb.Text) & ";", cnn, adOpenKeyset, adLockOptimistic
    If rstaddnew.RecordCount = 0 Then
                If rstmsave.State = adStateOpen Then rstmsave.Close
                rstmsave.Open "select * from medicine where medicineid='" & txtmID.Text & "'", cnn, adOpenKeyset, adLockOptimistic
                If rstrsave.State = adStateOpen Then rstrsave.Close
                rstrsave.Open "select * from Measure where Measure_Name='" & cmbuname.Text & "'", cnn, adOpenKeyset, adLockOptimistic
                rstmsave.Fields(3).Value = rstrsave.Fields(0).Value
                rstmsave.Fields(4).Value = StrConv(Trim(txtbestb.Text), vbProperCase)
                rstmsave.Update
                cx = cx + 1
    Else
        MsgBox "NO updation implemented", vbCritical, "Medical Store Automation"
        txtmName.SetFocus
        Exit Sub
    End If
End If
If a = False And cx = 1 Then
    If rstrsave.State = adStateOpen Then rstrsave.Close
    rstrsave.Open "select * from companyCat where CompanyName='" & cmbcomn.Text & "' and CategoryName='" & cmbcatn.Text & "'", cnn, adOpenKeyset, adLockOptimistic
    If rstcheck.State = adStateOpen Then rstcheck.Close
    rstcheck.Open "select * from medicine where medicinename='" & StrConv(Trim(txtmName.Text), vbProperCase) & "' and compcatid='" & rstrsave.Fields(3).Value & "'", cnn, adOpenKeyset, adLockOptimistic
    If Not rstcheck.RecordCount = 0 Then
        MsgBox "Record already present", vbCritical, "Medical Store Automation"
        txtmName.SetFocus
        Exit Sub
    End If
End If
If cx = 0 Then
        If rstmsave.State = adStateOpen Then rstmsave.Close
        rstmsave.Open "select * from medicine where medicineid='" & txtmID.Text & "'", cnn, adOpenKeyset, adLockOptimistic
        rstmsave.Fields(1).Value = StrConv(Trim(txtmName.Text), vbProperCase)
        If rstrsave.State = adStateOpen Then rstrsave.Close
        rstrsave.Open "select * from companyCat where CompanyName='" & cmbcomn.Text & "'and CategoryName='" & cmbcatn.Text & "'", cnn, adOpenKeyset, adLockOptimistic
        rstmsave.Fields(2).Value = rstrsave.Fields(3).Value
        rstmsave.Update
End If
showm
cmdCancel_Click
MSFlexGrid1.Row = 1
MSFlexGrid1_Click

MsgBox "Data Saved", vbOKOnly + vbInformation, "Medical Store Automation"

txtbestb.Enabled = True


End Sub

'Private Sub cmdSave_Click()
'If rstmedicine.State = adStateOpen Then rstmedicine.Close
'rstmedicine.Open "select * from medicinesdetails ", cnn, adOpenKeyset, adLockOptimistic
'rstmedicine.AddNew
'rstmedicine.Fields(0).Value = txtmid.Text
'rstmedicine.Fields(1).Value = txtmname.Text
'rstmedicine.Fields(2).Value = cmbcomn.Text
'rstmedicine.Fields(3).Value = cmbcatn.Text
'rstmedicine.Fields(4).Value = cmbuname.Text
'rstmedicine.Update
'If rstKey.State = adStateOpen Then rstKey.Close
'rstKey.Open "select medicine_id from key_id", cnn, adOpenKeyset, adLockOptimistic
'rstKey.AddNew
'rstKey.Fields(0).Value = txtmid.Text + 1
'rstKey.Update
'End Sub

Private Sub Form_Activate()
framedicin.Enabled = False
fraunit.Enabled = False
fraConfirmation.Enabled = False
a = True
showm
showu
bStyle (False)
If MSFlexGrid1.Rows > 1 Then
MSFlexGrid1.Row = 1
End If
MSFlexGrid1_Click
If MSFlexGrid1.Rows < 2 Then
    txtmID.Text = ""
End If
End Sub

Public Function showm()
If rstmedicine.State = adStateOpen Then rstmedicine.Close
rstmedicine.Open "select * from qrymedicine", cnn, adOpenKeyset, adLockOptimistic

MSFlexGrid1.Clear
Dim i%, j%
MSFlexGrid1.Cols = rstmedicine.Fields.Count - 2
MSFlexGrid1.Rows = rstmedicine.RecordCount + 1
For i = 0 To rstmedicine.Fields.Count - 3
    MSFlexGrid1.TextMatrix(0, i) = rstmedicine.Fields(i).Name
Next
For i = 1 To rstmedicine.RecordCount
        For j = 0 To rstmedicine.Fields.Count - 3
            MSFlexGrid1.TextMatrix(i, j) = rstmedicine.Fields(j).Value
        Next
rstmedicine.MoveNext
Next

End Function

Public Function uadd()
Dim i%
If Not (MSFlexGrid2.Row = 0) Then
        For i = 1 To MSFlexGrid2.Rows - 1
                If txtweight.Text = MSFlexGrid2.TextMatrix(i, 0) Then
                    MsgBox "Weight already added", vbCritical, "Medical Store Automation"
                    txtweight.SetFocus
                    Exit Function
                End If
                
        Next
End If

If rstunit.State = adStateOpen Then rstunit.Close
rstunit.Open "select * from medicinedetails", cnn, adOpenKeyset, adLockOptimistic
Dim j%
cc = cc + 1
If cc = 1 Then
MSFlexGrid2.Clear
MSFlexGrid2.Cols = rstunit.Fields.Count - 2
MSFlexGrid2.Rows = cc + 1
For i = 1 To rstunit.Fields.Count - 2
    MSFlexGrid2.TextMatrix(0, (i) - 1) = rstunit.Fields(i).Name
Next
End If

MSFlexGrid2.Rows = cc + 1
'For i = 1 To rstunit.RecordCount
'        For j = 0 To rstunit.Fields.Count - 2
'            MSFlexGrid2.TextMatrix(i, j) = rstunit.Fields(j).Value
'        Next
'rstunit.MoveNext
'Next

'MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 0) =
MSFlexGrid2.TextMatrix(cc, 0) = txtweight
MSFlexGrid2.TextMatrix(cc, 1) = txtrl

txtweight.Text = ""
txtrl.Text = ""
txtweight.SetFocus
End Function

Public Function showu()
If rstunit.State = adStateOpen Then rstunit.Close
rstunit.Open "select * from medicinedetails where medicineid='" & txtmID.Text & "'", cnn, adOpenKeyset, adLockOptimistic

MSFlexGrid2.Clear
Dim i%, j%
MSFlexGrid2.Cols = rstunit.Fields.Count - 2
MSFlexGrid2.Rows = rstunit.RecordCount + 1
For i = 1 To rstunit.Fields.Count - 2
    MSFlexGrid2.TextMatrix(0, (i - 1)) = rstunit.Fields(i).Name
Next

For i = 0 To rstunit.RecordCount - 1
        
            MSFlexGrid2.TextMatrix(i + 1, 0) = rstunit.Fields(1).Value
            MSFlexGrid2.TextMatrix(i + 1, 1) = rstunit.Fields(2).Value
rstunit.MoveNext
Next

End Function
Public Function uremove()
If rstunit.State = adStateOpen Then rstunit.Close
rstunit.Open "select * from medicinedetails", cnn, adOpenKeyset, adLockOptimistic
Dim i%, j%

For i = MSFlexGrid2.Row To MSFlexGrid2.Rows - 2
    MSFlexGrid2.TextMatrix(i, 0) = MSFlexGrid2.TextMatrix(i + 1, 0)
    MSFlexGrid2.TextMatrix(i, 1) = MSFlexGrid2.TextMatrix(i + 1, 1)
Next

If cc = 0 Then
MsgBox "No Record to Remove", vbCritical, "Medical Store Automation"
Exit Function
'MSFlexGrid2.Clear
'MSFlexGrid2.Cols = rstunit.Fields.Count - 1
'MSFlexGrid2.Rows = cc + 1
'For i = 0 To rstunit.Fields.Count - 2
'    MSFlexGrid2.TextMatrix(0, i) = rstunit.Fields(i).Name
'Next
End If

MSFlexGrid2.Rows = cc
'For i = 1 To rstunit.RecordCount
'        For j = 0 To rstunit.Fields.Count - 2
'            MSFlexGrid2.TextMatrix(i, j) = rstunit.Fields(j).Value
'        Next
'rstunit.MoveNext
'Next

'MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 0) =
'MSFlexGrid2.TextMatrix(cc, 1) = txtweight
'MSFlexGrid2.TextMatrix(cc, 2) = txtrl

cc = cc - 1
End Function

Private Sub Form_Load()

If rstKey.State = adStateOpen Then rstKey.Close
rstKey.Open "select medicine_id from key_id", cnn, adOpenKeyset, adLockOptimistic
txtmID.Text = rstKey.Fields(0).Value


End Sub

Private Sub MSFlexGrid1_Click()
    Dim j%
    j = MSFlexGrid1.Row
    cmbcomn.Clear
    cmbcatn.Clear
    cmbuname.Clear
    txtmID.Text = MSFlexGrid1.TextMatrix(j, 0)
    txtmName = MSFlexGrid1.TextMatrix(j, 1)
    cmbcomn.additem MSFlexGrid1.TextMatrix(j, 2)
    cmbcatn.additem MSFlexGrid1.TextMatrix(j, 3)
    cmbuname.additem MSFlexGrid1.TextMatrix(j, 4)
    txtbestb = MSFlexGrid1.TextMatrix(j, 5)
    cmbcomn.ListIndex = 0
    If cmbcatn.ListCount > 0 Then
       cmbcatn.ListIndex = 0
    End If
    cmbuname.ListIndex = 0
    showu
    'cmdadd.SetFocus
End Sub

Private Sub txtbestb_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then

Else
KeyAscii = keynum(KeyAscii, txtbestb.Text)
End If
End Sub

Private Sub txtmName_KeyPress(KeyAscii As Integer)
If Len(txtmName.Text) = 0 Or txtmName.SelStart = 0 Then
'    If KeyAscii = 32 Then
'        KeyAscii = 0
'    End If
    KeyAscii = keyboth(KeyAscii, txtmName.Text)
Else
'''
    If KeyAscii = 8 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 46 Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 44 Then
    
    Else
        KeyAscii = keyboth(KeyAscii, txtmName.Text)
    End If
    If Len(txtmName.Text) > 1 And (KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 46 Or KeyAscii = 32 Or KeyAscii = 44) Then
    Dim aa1%, aa2%
    aa1 = Asc(Right((Left(txtmName.Text, txtmName.SelStart)), 1))
    aa2 = Asc(Right((Left(txtmName.Text, txtmName.SelStart + 1)), 1))
        If aa1 = KeyAscii Or aa2 = KeyAscii Then
        KeyAscii = 0
        End If
        If ((aa1 = 43 Or aa1 = 45 Or aa2 = 32 Or aa1 = 46 Or aa1 = 44) And (KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 46 Or KeyAscii = 32 Or KeyAscii = 44)) Then
                KeyAscii = 0
        End If
        If ((aa2 = 43 Or aa2 = 45 Or aa2 = 32 Or aa2 = 46 Or aa1 = 44) And (KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 46 Or KeyAscii = 32 Or KeyAscii = 44)) Then
                KeyAscii = 0
        End If
    End If
End If
End Sub

Private Sub txtrl_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then

Else
KeyAscii = keynum(KeyAscii, txtbestb.Text)
End If
If KeyAscii = 48 And txtrl.SelStart = 0 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtrl_LostFocus()

If txtrl.Text = "" Then
    txtrl.Text = "0"
End If
If txtrl.Text = "" Or Not (txtrl.Text >= 4 And txtrl.Text <= 30) Then
    MsgBox "Best before should be  4 to 30 ", vbCritical, "Medical Store Automation"
    txtrl.SetFocus
    txtrl.Text = "4"
End If
End Sub

Private Sub txtweight_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then

Else
KeyAscii = keynum(KeyAscii, txtbestb.Text)
End If
If KeyAscii = 48 And txtweight.SelStart = 0 Then
    KeyAscii = 0
End If
End Sub

