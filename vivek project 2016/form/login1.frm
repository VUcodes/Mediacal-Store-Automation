VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5715
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5715
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   5040
      Top             =   360
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3360
      MaxLength       =   25
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtUser 
      Height          =   495
      Left            =   3360
      MaxLength       =   15
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "User Name"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstLogin As New ADODB.Recordset
Dim X As Integer
Private Sub clearE()
 If txtPassword.Text = "" And txtUser.Text = "" Then
     cmdClear.Enabled = False
    Else
        cmdClear.Enabled = True
    End If
 If Not (txtPassword.Text = "" Or txtUser.Text = "") Then
    cmdLogin.Enabled = True
    Else
        cmdLogin.Enabled = False
End If
End Sub
Private Sub cmdCancel_Click()
Dim i As VbMsgBoxResult
i = MsgBox("Do You Want To Exit", vbYesNo + vbExclamation)
If i = vbYes Then
    Unload Me
Else
        txtUser.SetFocus
    cmdLogin.Enabled = False
    cmdClear_Click
    Exit Sub
End If


End Sub

Private Sub cmdLogin_Click()
If Len(txtUser.Text) <= 0 Then
    MsgBox "UserName cannot be Left Blank", vbCritical
    txtUser.SetFocus
    Exit Sub
End If
If Len(txtPassword.Text) <= 0 Then
    MsgBox "Password cannot be Left Blank", vbCritical
    txtPassword.SetFocus
    Exit Sub
End If

If rstLogin.State = adStateOpen Then rstLogin.Close
rstLogin.Open "select * from login where username='" & txtUser.Text & "' and pass='" & txtPassword.Text & "'", cnn, adOpenKeyset, adLockOptimistic

If rstLogin.RecordCount = 1 Then
    MsgBox "Login Successful", vbExclamation
    user = txtUser.Text
    Timer1.Interval = 500
    Timer1.Enabled = True
    
    'Exist
    'If rstLogin.State = adStateOpen Then rstLogin.Close
'rstLogin.Open "select * from login", cnn, adOpenKeyset, adLockOptimistic

'While rstLogin.EOF = False
 '   If StrConv(rstLogin.Fields(0).Value, vbProperCase) = StrConv(txtUser.Text, vbProperCase) Then
  '      rstLogin.Fields(0).Value = StrConv(rstLogin.Fields(0).Value, vbUpperCase)
   ' Else
    '    rstLogin.Fields(0).Value = StrConv(rstLogin.Fields(0).Value, vbProperCase)
    'End If
    'rstLogin.MoveNext
'Wend
    'exist
    Else
    MsgBox "Invalid UserName or Password"
    txtUser.Text = ""
    txtPassword.Text = ""
    txtUser.SetFocus
End If

End Sub

Private Sub cmdClear_Click()
txtUser.Text = ""
txtPassword.Text = ""
txtUser.SetFocus
cmdLogin.Enabled = False
End Sub



Private Sub Command1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    cmdCancel_Click
End If
End Sub

Private Sub Form_Load()
txtUser.Text = "Administrator"
If rstLogin.State = adStateOpen Then rstLogin.Close
rstLogin.Open "select * from login", cnn, adOpenKeyset, adLockOptimistic

'Dim i As Integer
'For i = 0 To rstLogin.RecordCount - 1
    'If rstLogin.Fields(0).Value = StrConv(rstLogin.Fields(0).Value, vbUpperCase) Then
        'txtUser.Text = StrConv(rstLogin.Fields(0).Value, vbProperCase)
        txtUser.SelStart = 0
        txtUser.SelLength = Len(txtUser.Text)
   '     Exit For
   ' End If

    'rstLogin.MoveNext
'Next
'If txtUser.Text = Null Then
'rstLogin.MoveFirst
'txtUser.Text = rstLogin.Fields(0).Value
'End If
cmdLogin.Enabled = False

End Sub




Private Sub Timer1_Timer()
If ProgressBar1.Value >= ProgressBar1.Max Then
    MDIForm1.Show
Else
    ProgressBar1.Value = ProgressBar1.Value + 20
End If
    
End Sub

Private Sub txtPassword_Change()
clearE
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdLogin_Click
    End If
End Sub



Private Sub txtUser_Change()
clearE
End Sub

