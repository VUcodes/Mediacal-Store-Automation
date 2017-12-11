VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8955
   LinkTopic       =   "Form4"
   ScaleHeight     =   6090
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Manupilation"
      Height          =   1215
      Left            =   3240
      TabIndex        =   12
      Top             =   4560
      Width           =   5655
      Begin VB.CommandButton cmdD 
         Caption         =   "Delete User"
         Height          =   615
         Left            =   3960
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Change Passowrd"
         Height          =   615
         Left            =   2040
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add User"
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Delete User"
      Height          =   5295
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   3015
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "Delete"
         Height          =   735
         Left            =   0
         TabIndex        =   11
         Top             =   4560
         Width           =   3015
      End
      Begin VB.ListBox lstDelete 
         Height          =   3885
         ItemData        =   "loginManupilation.frx":0000
         Left            =   240
         List            =   "loginManupilation.frx":0002
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "New User"
      Height          =   3855
      Left            =   3240
      TabIndex        =   0
      Top             =   600
      Width           =   5655
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   615
         Left            =   3120
         TabIndex        =   8
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   615
         Left            =   600
         TabIndex        =   7
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtCPass 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtPass 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtNewUser 
         Height          =   495
         Left            =   3120
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Confirm Password"
         Height          =   615
         Left            =   480
         TabIndex        =   3
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   615
         Left            =   480
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "New User Name"
         Height          =   615
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rstL As New ADODB.Recordset
Dim p As Boolean
Private Sub listD()
If rstL.State = adStateOpen Then rstL.Close
rstL.Open "select * from login", cnn, adOpenKeyset, adLockOptimistic
lstDelete.Clear
While (rstL.EOF = False)
    If Not (rstL.Fields(0).Value = StrConv("administrator", vbProperCase)) Then
        lstDelete.additem rstL.Fields(0).Value
    End If
    rstL.MoveNext
Wend

End Sub
Private Sub cmdadd_Click()
p = True
Frame1.Enabled = True
txtNewUser.SetFocus
Frame1.Caption = "New User"
Frame2.Enabled = False
Label1.Caption = "New User Name"
Label2.Caption = "Password"
Label3.Caption = "Confirm Password"
End Sub

Private Sub cmdClear_Click()
txtNewUser.Text = ""
txtPass.Text = ""
txtCPass.Text = ""
End Sub

Private Sub cmdD_Click()
Frame1.Enabled = False
Frame2.Enabled = True
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer
i = 0
While (i < lstDelete.ListCount)
    If lstDelete.selected(i) = True Then
        If rstL.State = adStateOpen Then rstL.Close
         rstL.Open "delete from login where username='" & lstDelete.List(i) & "'", cnn, adOpenKeyset, adLockOptimistic
    End If
i = i + 1
Wend
listD
End Sub

Private Sub cmdedit_Click()
p = False
Frame1.Enabled = True
Frame2.Enabled = True
txtNewUser.SetFocus
Frame1.Caption = "Edit"
Label1.Caption = "User Name"
Label3.Caption = "New Password"
txtCPass.Enabled = False
cmdDelete.Enabled = True
End Sub

Private Sub cmdEdit_LostFocus()
txtNewUser.SetFocus
End Sub

Private Sub cmdSave_Click()

'
If p = True Then
        If txtPass.Text = txtCPass.Text Then
        If rstL.State = adStateOpen Then rstL.Close
            rstL.Open "select * from login  where username='" & txtNewUser.Text & "' and pass='" & txtPass.Text & "'", cnn, adOpenKeyset, adLockOptimistic
                If rstL.RecordCount > 0 Then
                 MsgBox "User already exists"
                 cmdClear_Click
             Else
                    If rstL.State = adStateOpen Then rstL.Close
                     rstL.Open "select * from login ", cnn, adOpenKeyset, adLockOptimistic
                 rstL.MoveLast
                 rstL.AddNew
                 rstL.Fields(0).Value = txtNewUser.Text
                 rstL.Fields(1).Value = txtPass.Text
                 rstL.Fields(2).Value = 0
                 rstL.Update
                 cmdClear_Click
                 MsgBox "Record Updated", vbInformation
                End If
            Else
             MsgBox "Password does not Match"
             txtCPass.Text = ""
            
        End If
Else
    
    txtPass_LostFocus
    If rstL.RecordCount > 0 Then
        rstL.Fields(0).Value = txtNewUser.Text
                 rstL.Fields(1).Value = txtCPass.Text
                 rstL.Fields(2).Value = 0
                 
    End If
    rstL.Update
End If
'
listD
Frame1.Enabled = False
Frame2.Enabled = False
End Sub

Private Sub Form_Load()
Frame1.Enabled = False
Frame2.Enabled = False
listD
cmdClear.Enabled = False
cmdSave.Enabled = False
End Sub


Private Sub lstDelete_Click()
Dim i As Integer
For i = 0 To lstDelete.ListCount
    If lstDelete.selected(i) = True Then
        txtNewUser.Text = lstDelete.List(i)
        Exit For
    End If
Next
End Sub

Private Sub txtCPass_KeyPress(KeyAscii As Integer)
cmdSave.Enabled = True

End Sub

Private Sub txtNewUser_KeyPress(KeyAscii As Integer)

keyboth KeyAscii, txtNewUser.Text
If KeyAscii > 0 Then
    
    cmdClear.Enabled = True
End If
End Sub


Private Sub txtPass_LostFocus()
If rstL.State = adStateOpen Then rstL.Close
    rstL.Open "select * from login  where username='" & txtNewUser.Text & "' and pass='" & txtPass.Text & "'", cnn, adOpenKeyset, adLockOptimistic
    If rstL.RecordCount > 0 Then
        txtCPass.Enabled = True
    End If
End Sub
