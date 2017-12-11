VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmmedicindetail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medicine Detail"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraManupilation 
      Caption         =   "Manupilation"
      Height          =   1695
      Left            =   4200
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add"
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1695
      Left            =   2880
      TabIndex        =   7
      Top             =   240
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2990
      _Version        =   393216
   End
   Begin VB.Frame fraConfirmation 
      Caption         =   "Confirmation"
      Height          =   1695
      Left            =   5400
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame fraunit 
      Height          =   4215
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtweight 
         Height          =   375
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtrl 
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3625
         _Version        =   393216
      End
      Begin VB.Label Label6 
         Caption         =   "Weight"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Reorder Level"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmmedicindetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstmedicine As New ADODB.Recordset
Dim rstunit As New ADODB.Recordset
Dim rstrsave As New ADODB.Recordset
Dim rstmsave As New ADODB.Recordset
Dim rstaddnew As New ADODB.Recordset
Dim rstcheck As New ADODB.Recordset, add As Boolean

Private Sub cmdadd_Click()
fraManupilation.Enabled = False
MSFlexGrid1.Enabled = False
fraunit.Enabled = True
MSFlexGrid2_Click
MSFlexGrid2.Enabled = False
fraConfirmation.Enabled = True
add = True
txtweight.Text = ""
txtrl.Text = ""
txtweight.SetFocus
End Sub

Private Sub cmdCancel_Click()
fraManupilation.Enabled = True
MSFlexGrid1.Enabled = True
fraunit.Enabled = False
fraConfirmation.Enabled = False
txtrl.Text = ""
txtweight.Text = ""
cmdadd.SetFocus
End Sub

Private Sub cmdedit_Click()
fraManupilation.Enabled = False
MSFlexGrid1.Enabled = False
fraunit.Enabled = True
MSFlexGrid2_Click
fraConfirmation.Enabled = True
MSFlexGrid2.Enabled = True
add = False
End Sub

Private Sub cmdSave_Click()
If txtrl.Text = "" Then
    MsgBox "Reorder level cannot be empty or 0", vbCritical, "Medical Store Automation"
    txtrl.SetFocus
    Exit Sub
End If
If Not CInt(txtrl.Text) > 0 Then
    MsgBox "Reorder level cannot be empty or 0", vbCritical, "Medical Store Automation"
    txtrl.SetFocus
    Exit Sub
End If
Dim i%

For i = 1 To MSFlexGrid2.Rows - 1
    If MSFlexGrid2.TextMatrix(i, 0) = txtweight.Text Then
        MsgBox "Record already present", vbCritical, "Medical Store Automation"
        txtweight.SetFocus
        Exit Sub
    End If
Next

Dim pk1%
    If rstaddnew.State = adStateOpen Then rstaddnew.Close
    rstaddnew.Open "select * from key_id", cnn, adOpenKeyset, adLockOptimistic
    If add = True Then
    If rstmsave.State = adStateOpen Then rstmsave.Close
    rstmsave.Open "select * from medicinedetails", cnn, adOpenKeyset, adLockOptimistic
    pk1 = rstaddnew.Fields(6).Value + 1
    
        If pk1 = 100000 Then
            MsgBox ("Data base if full ")
            
            Exit Sub
        End If
        
        rstmsave.AddNew
        rstmsave.Fields(0).Value = pk1
        rstmsave.Fields(1).Value = txtweight.Text
        rstmsave.Fields(2).Value = txtrl.Text
        rstmsave.Fields(3).Value = MSFlexGrid2.TextMatrix(1, 2)
        
        rstaddnew.Fields(6).Value = pk1
        rstaddnew.Update
        rstmsave.Update
    End If
    If add = False Then
        If rstmsave.State = adStateOpen Then rstmsave.Close
        rstmsave.Open "select * from medicinedetails where medicineid='" & MSFlexGrid2.TextMatrix(1, 2) & "' and Weight='" & (MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 0)) & "'", cnn, adOpenKeyset, adLockOptimistic
    
        rstmsave.Fields(0).Value = pk1
        rstmsave.Fields(1).Value = txtweight.Text
        rstmsave.Fields(2).Value = txtrl.Text
        rstmsave.Fields(3).Value = MSFlexGrid2.TextMatrix(1, 2)
        
        rstaddnew.Fields(6).Value = pk1
        rstaddnew.Update
        rstmsave.Update
        
    End If
cmdCancel_Click
showu
MsgBox "Data Saved", vbOKOnly + vbInformation, "Medical Store Automation"
End Sub

Private Sub Form_Activate()
showm
showu
fraunit.Enabled = False
fraConfirmation.Enabled = False
End Sub
Public Function showm()
If rstmedicine.State = adStateOpen Then rstmedicine.Close
rstmedicine.Open "select * from qrymedicine", cnn, adOpenKeyset, adLockOptimistic

MSFlexGrid1.Clear
Dim i%, j%
MSFlexGrid1.Cols = rstmedicine.Fields.Count - 4
MSFlexGrid1.Rows = rstmedicine.RecordCount + 1
For i = 0 To rstmedicine.Fields.Count - 5
    MSFlexGrid1.TextMatrix(0, i) = rstmedicine.Fields(i).Name
Next
For i = 1 To rstmedicine.RecordCount
        For j = 0 To rstmedicine.Fields.Count - 5
            MSFlexGrid1.TextMatrix(i, j) = rstmedicine.Fields(j).Value
        Next
rstmedicine.MoveNext
Next
End Function
Public Function showu()
If rstunit.State = adStateOpen Then rstunit.Close
rstunit.Open "select * from medicinedetails where medicineid='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) & "'", cnn, adOpenKeyset, adLockOptimistic

MSFlexGrid2.Clear
Dim i%, j%
MSFlexGrid2.Cols = rstunit.Fields.Count - 1
MSFlexGrid2.Rows = rstunit.RecordCount + 1
For i = 1 To rstunit.Fields.Count - 1
    MSFlexGrid2.TextMatrix(0, (i - 1)) = rstunit.Fields(i).Name
Next

For i = 0 To rstunit.RecordCount - 1
        
            MSFlexGrid2.TextMatrix(i + 1, 0) = rstunit.Fields(1).Value
            MSFlexGrid2.TextMatrix(i + 1, 1) = rstunit.Fields(2).Value
            MSFlexGrid2.TextMatrix(i + 1, 2) = rstunit.Fields(3).Value
rstunit.MoveNext
Next

End Function

Private Sub MSFlexGrid1_Click()
showu
End Sub

Private Sub MSFlexGrid2_Click()
txtweight.Text = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 0)
txtrl.Text = MSFlexGrid2.TextMatrix(MSFlexGrid2.Row, 1)
End Sub

Private Sub txtrl_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then

Else
KeyAscii = keynum(KeyAscii, txtrl.Text)
End If
End Sub

Private Sub txtweight_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then

Else
KeyAscii = keynum(KeyAscii, txtweight.Text)
End If
End Sub
