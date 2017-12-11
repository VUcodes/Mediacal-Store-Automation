VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmunit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Measuring Unit"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   3915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSave 
      Caption         =   "Confirmation"
      Height          =   1815
      Left            =   2520
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame frmManupilation 
      Caption         =   "Manipulation"
      Height          =   1935
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   435
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2566
      _Version        =   393216
      BackColorBkg    =   -2147483633
      Appearance      =   0
   End
   Begin VB.Frame frmM 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtmName 
         BackColor       =   &H8000000F&
         Height          =   495
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtmID 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Unit Name"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Unit ID"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmunit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstAdd As New ADODB.Recordset
Dim rstmeasure As New ADODB.Recordset
Dim rstAddID As New ADODB.Recordset
Dim rstEdit As New ADODB.Recordset
Dim rstDuplicate As New ADODB.Recordset
Dim pk As Integer
Dim i As Integer, j As Integer
Dim ca As String
Dim addb As Boolean
Dim d As String

Private Sub selected()
    txtmid.SelStart = 0
    txtmname.SelLength = Len(txtmname.Text)
    
End Sub

Public Sub bordertext(i As Integer)
    If i = 1 Then
        txtmid.BorderStyle = vbFixedSingle
        txtmname.BackColor = &H80000009
    Else
        txtmid.BorderStyle = 0
        txtmname.BackColor = &H8000000F

    End If
End Sub

Private Sub clear_all()
txtmid.Text = ""
txtmname.Text = ""
End Sub

Private Sub additem()
addb = True
clear_all
txtmname.SetFocus
If rstAddID.State = adStateOpen Then rstAddID.Close
rstAddID.Open "select * from key_id", cnn, adOpenKeyset, adLockOptimistic
pk = rstAddID.Fields(4).Value + 1
If pk = 10 Then
    MsgBox "Category limit is 9", vbCritical, "Medical Store Automation"
    MSFlexGrid1.Enabled = True
    frmM.Enabled = False
    frmSave.Enabled = False
    frmManupilation.Enabled = True
    showData
    Exit Sub
End If
    
If pk > 0 And pk < 10 Then
txtmid.Text = pk
Else
txtmid.Text = pk
End If

End Sub
Private Sub cmdadd_Click()
frmM.Enabled = True
frmSave.Enabled = True
MSFlexGrid1.Enabled = False
frmManupilation.Enabled = False
txtmname.Enabled = True
bordertext (1)
additem

End Sub

Private Sub cmdedit_Click()
frmM.Enabled = True
MSFlexGrid1.Enabled = False
frmSave.Enabled = True
frmManupilation.Enabled = False
If txtmname.Text = "" Then
    MsgBox "Select Option from the grid"
    frmManupilation.Enabled = False
Else
    txtmname.Enabled = True
    txtmname.SetFocus
    txtmname.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
    selected
    addb = False
End If
bordertext (1)
End Sub

Private Sub cmdSave_Click()
If rstAdd.State = adStateOpen Then rstAdd.Close
If rstAddID.State = adStateOpen Then rstAddID.Close
If rstDuplicate.State = adStateOpen Then rstDuplicate.Close

rstAdd.Open "select * from measure", cnn, adOpenKeyset, adLockOptimistic
rstAddID.Open "select measure_Id from key_id", cnn, adOpenKeyset, adLockOptimistic

If Len(txtmname.Text) < 2 Then
    MsgBox (" Unit name must be grater then 2"), vbCritical, "Medical Store Automation"
    txtmname.SetFocus
    Exit Sub
End If

j = 0

ca = Left(txtmname.Text, 1)
For i = 1 To Len(txtmname.Text)
   If (ca = Right(Left(txtmname.Text, i), 1)) Then
    j = j + 1
   End If
Next
If j = Len(txtmname.Text) And j > 3 Then
    MsgBox ("Invalid unit name"), vbCritical, "Medical Store Automation"
    txtmname.SetFocus
    
    Exit Sub
End If
  
    
    If txtmname.Text = "" Then
            MsgBox ("Measuring name cannot be  empty"), vbCritical, "Medical Store Automation"
            txtmname.SetFocus
            Exit Sub
    End If
    rstDuplicate.Open "select * from Measure  where measure_Name= '" + txtmname.Text + "'", cnn, adOpenKeyset, adLockOptimistic
    If rstDuplicate.RecordCount > 0 Then
        MsgBox ("Unit name already exist"), vbCritical, "Medical Store Automation"
            txtmname.SetFocus
            txtmname = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
            selected
            Exit Sub
    End If
     
    rstDuplicate.Close
    If addb = True Then
    
    rstAddID.MoveFirst
    rstAddID.Fields(0).Value = rstAddID.Fields(0).Value + 1
    rstAddID.Update
    
    rstAdd.AddNew
    rstAdd.Fields(0).Value = txtmid.Text
    rstAdd.Fields(1).Value = StrConv(Trim(txtmname.Text), vbProperCase)
    rstAdd.Update
Else
    If rstEdit.State = adStateOpen Then rstEdit.Close
    rstEdit.Open "select * from measure where measure_ID='" & txtmid.Text & "'", cnn, adOpenKeyset, adLockOptimistic
    rstEdit.Fields(1).Value = StrConv(Trim(txtmname.Text), vbProperCase)
    rstEdit.Update
End If

cmdCancel_Click
showData
MsgBox "Data Saved", vbOKOnly + vbInformation, "Medical Store Automation"
End Sub

Private Sub cmdCancel_Click()
frmM.Enabled = False
MSFlexGrid1.Enabled = True
frmSave.Enabled = False
frmManupilation.Enabled = True
txtmname.Enabled = False
cmdadd.SetFocus
bordertext (0)
showData

End Sub

Private Sub Form_Load()

Dim rstmeasure As New ADODB.Recordset
    Dim X As Integer, Y As Integer
     
    If rstmeasure.State = adStateOpen Then rstmeasure.Close
    
    MSFlexGrid1.Visible = True
    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 2
    
    rstmeasure.Open "select * from Measure", cnn, adOpenKeyset, adLockOptimistic
   
    MSFlexGrid1.Cols = rstmeasure.Fields.Count
    For X = 0 To rstmeasure.Fields.Count - 1
        MSFlexGrid1.TextMatrix(0, X) = rstmeasure.Fields(X).Name
    
        
    Next
    
   
    If rstmeasure.RecordCount = 0 Then
    For X = 1 To rstmeasure.RecordCount
        For Y = 0 To rstmeasure.Fields.Count - 1
            MSFlexGrid1.TextMatrix(X, Y) = rstmeasure.Fields(Y).Value
        Next
        rstmeasure.MoveNext
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    Next
    If Not pk = 0 Then
    rstmeasure.MoveFirst
    End If
    txtmid.Text = rstmeasure.Fields(0).Value
    txtmname.Text = rstmeasure.Fields(1).Value
   Else
    
   End If
   showData
txtmid.Enabled = False
txtmname.Enabled = False

End Sub

Private Sub showData()

If rstmeasure.State = adStateOpen Then rstmeasure.Close
    Dim X As Integer, Y As Integer
    MSFlexGrid1.Visible = True
    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 2
    
    rstmeasure.Open "select * from measure", cnn, adOpenKeyset, adLockOptimistic
    
    txtmid.Text = rstmeasure.Fields(0).Value
    txtmname.Text = rstmeasure.Fields(1).Value
    
    MSFlexGrid1.Cols = rstmeasure.Fields.Count
        For X = 0 To rstmeasure.Fields.Count - 1
            MSFlexGrid1.TextMatrix(0, X) = rstmeasure.Fields(X).Name
           
        Next
         MSFlexGrid1.ColWidth(0) = 950
          MSFlexGrid1.ColWidth(1) = 1400
'    MSFlexGrid1.Height = rstmeasure.Fields.Count * 2000 + 350
'    MSFlexGrid1.Width = 4300
    For X = 1 To rstmeasure.RecordCount
        For Y = 0 To rstmeasure.Fields.Count - 1
            MSFlexGrid1.TextMatrix(X, Y) = rstmeasure.Fields(Y).Value
        Next
        rstmeasure.MoveNext
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If frmManupilation.Enabled = False Then
    MsgBox "Please Complete the session", vbCritical, "Medical Store Automation"
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
    txtmid.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
    d = txtmid.Text
    txtmname.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
End Sub

Private Sub txtmName_KeyPress(KeyAscii As Integer)
KeyAscii = key(KeyAscii, txtmname.Text)
If KeyAscii = 32 Then
KeyAscii = 0
End If
End Sub

