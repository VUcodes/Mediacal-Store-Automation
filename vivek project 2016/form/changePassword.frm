VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CHANGE"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   360
      TabIndex        =   9
      Top             =   3960
      Width           =   4815
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Reset"
         Height          =   495
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin VB.TextBox txtUser 
         Height          =   495
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtOldPass 
         Height          =   495
         Left            =   2520
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtNewPass 
         Height          =   495
         Left            =   2520
         TabIndex        =   2
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtConPass 
         Height          =   495
         Left            =   2520
         TabIndex        =   1
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "UserName"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Old Password"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "New Password"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Confirm Password"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstChange As New ADODB.Recordset
Private Sub cmdUpdate_Click()
If Len(txtOldPass.Text) < 0 Or Len(txtNewPass.Text) Or Len(txtConPass.Text) Then
    MsgBox "Fields cannot be Left Blank", vbCritical
End If
If rstChange.State = adStateOpen Then rstChange.Close
rstChange.Open "select * from login where username='" & txtUser.Text & "'", cnn, adOpenKeyset, adLockOptimistic
If Len(txtNewPass.Text) >= 8 Or Len(txtConPass.Text) Then
    MsgBox "Password To Short", vbCritical
    Exit Sub
End If
If rstChange.Fields(1).Value = txtOldPass.Text Then
    If txtNewPass.Text <> txtConPass.Text Then
        MsgBox "New and Confirm Does not Match", vbCritical
        txtNewPass.Text = ""
        txtConPass.Text = ""
        txtNewPass.SetFocus
        Exit Sub
    End If
    rstChange.Fields(1) = txtNewPass.Text
    rstChange.Update
    MsgBox "Password Changed"
    txtOldPass.Text = ""
    txtNewPass.Text = ""
    txtConPass.Text = ""
Else
    MsgBox "Existing Password does not Match", vbCritical
    txtOldPass.Text = ""
    txtNewPass.Text = ""
    txtConPass.Text = ""
    txtOldPass.SetFocus
End If
End Sub

Private Sub Command1_Click()
MsgBox "Want to Exit", vbCritical
Dim m As VbMsgBoxResult
If m = vbYes Then
    
End If
End Sub


Private Sub Form_Load()
cmdUpdate.Enabled = False
End Sub

Private Sub txtUser_Change()
If Len(txtUser.Text) = 0 Then
    cmdUpdate.Enabled = False
Else
    cmdUpdate.Enabled = True
End If
End Sub
