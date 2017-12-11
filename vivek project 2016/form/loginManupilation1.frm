VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Manipulation"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Manipulation"
      Height          =   1215
      Left            =   3240
      TabIndex        =   12
      Top             =   4560
      Width           =   5655
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   615
         Left            =   4320
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdD 
         Caption         =   "Delete User"
         Height          =   615
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Change Password"
         Height          =   615
         Left            =   1560
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add User"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Delete User"
      Height          =   5295
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   3015
      Begin VB.CommandButton cmdDCancel 
         Appearance      =   0  'Flat
         Caption         =   "Cancel"
         Height          =   615
         Left            =   1560
         TabIndex        =   18
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "Delete"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   4320
         Width           =   1095
      End
      Begin VB.ListBox lstDelete 
         Height          =   3885
         ItemData        =   "loginManupilation1.frx":0000
         Left            =   240
         List            =   "loginManupilation1.frx":0002
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
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   615
         Left            =   3960
         TabIndex        =   16
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   615
         Left            =   2160
         TabIndex        =   8
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtCPass 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3120
         MaxLength       =   24
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtPass 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3120
         MaxLength       =   24
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtNewUser 
         Height          =   495
         Left            =   3120
         MaxLength       =   14
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
         Caption         =   "User Name"
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
Dim del As Boolean
Private Sub clearE()

If txtNewUser.Text = "" And txtPass.Text = "" And txtCPass.Text = "" Then
    cmdClear.Enabled = False
Else
    cmdClear.Enabled = True
End If
If txtNewUser.Text = "" Or txtPass.Text = "" Or txtCPass.Text = "" Then
    cmdSave.Enabled = False
Else
    cmdSave.Enabled = True
End If

End Sub
Private Sub listD()
If rstL.State = adStateOpen Then rstL.Close
rstL.Open "select * from login", cnn, adOpenKeyset, adLockOptimistic
lstDelete.Clear
While (rstL.EOF = False)
    If p = False Then
            lstDelete.additem StrConv(rstL.Fields(0).Value, vbProperCase)
    Else
        If Not (StrConv(rstL.Fields(0).Value, vbProperCase) = "Administrator") Then
            lstDelete.additem StrConv(rstL.Fields(0).Value, vbProperCase)
        End If
    End If
    rstL.MoveNext
Wend

End Sub
Private Sub cmdadd_Click()

p = True
listD
Frame1.Enabled = True
txtNewUser.Enabled = True
txtNewUser.Text = ""
txtNewUser.SetFocus
Frame1.Caption = "New User"
Frame2.Enabled = False
Label1.Caption = "User Name"
Label2.Caption = "Password"
Label3.Caption = "Confirm Password"
cmdDelete.Enabled = False
cmdSave.Enabled = False
cmdCancel.Enabled = True
Frame3.Enabled = False
End Sub

Private Sub cmdCancel_Click()
Dim i As VbMsgBoxResult

i = MsgBox("All Changes Will Be Undo", vbCritical + vbYesNo)
If i = vbYes Then
    cmdClear_Click
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = True
    cmdClear.Enabled = False
    
Else
    If Frame2.Enabled = True Then
        txtPass.SetFocus
        Exit Sub
    Else
        If Frame1.Enabled = True Then
            cmdClear_Click
            txtNewUser.SetFocus
            End If
    End If
End If

End Sub

Private Sub cmdClear_Click()
If Frame2.Enabled = False Then
    txtNewUser.Text = ""
    txtPass.Text = ""
    txtCPass.Text = ""
    txtNewUser.SetFocus
    Else
    txtPass.Text = ""
    txtCPass.Text = ""
    txtPass.SetFocus
End If
End Sub

Private Sub cmdD_Click()
Frame1.Enabled = False
Frame2.Enabled = True
del = True
txtNewUser.Text = ""
txtPass.Text = ""
txtCPass.Text = ""
cmdDelete.Enabled = True
cmdCancel.Enabled = False
p = True
listD
'cmdD.Enabled = False
cmdDCancel.Enabled = True
Frame3.Enabled = False
End Sub

Private Sub cmdDCancel_Click()
Dim i As VbMsgBoxResult

i = MsgBox("All Changes Will Be Undo", vbCritical + vbYesNo)
If i = vbYes Then
Frame2.Enabled = False
cmdDCancel.Enabled = False
cmdDelete.Enabled = False
Frame3.Enabled = True
Else
    Exit Sub
End If
End Sub

Private Sub cmdDelete_Click()

Dim i As Integer, j As Integer
j = 0
i = 0
While (i < lstDelete.ListCount)
    If lstDelete.selected(i) = True Then
        j = j + 1
    End If
   i = i + 1
Wend
If j Then

    Dim con As VbMsgBoxResult
    con = MsgBox("Are You Sure You Want to Delete", vbCritical + vbYesNo)
    If con = vbYes Then
    i = 0
        While (i < lstDelete.ListCount)
                 If lstDelete.selected(i) = True Then
                     If rstL.State = adStateOpen Then rstL.Close
                    rstL.Open "delete from login where username='" & lstDelete.List(i) & "'", cnn, adOpenKeyset, adLockOptimistic
                End If
    
            i = i + 1
         Wend
        MsgBox "Selected Users Deleted", vbCritical
        
        Frame3.Enabled = True
        cmdDelete.Enabled = False
        cmdCancel.Enabled = False
        cmdDCancel.Enabled = False
    End If
Else
    MsgBox "Select User First", vbInformation
End If
p = True
listD
End Sub

Private Sub cmdedit_Click()
p = False
listD
cmdCancel.Enabled = True
Frame1.Enabled = True
Frame2.Enabled = True
txtNewUser.Enabled = False
Frame1.Caption = "Edit"
Label1.Caption = "User Name"
Label2.Caption = "Password"
Label3.Caption = "Confirm Password"
txtPass.SetFocus
'txtCPass.Enabled = False
cmdDelete.Enabled = False
del = False
txtNewUser.Text = ""
txtCPass.Text = ""
txtPass.Text = ""
txtNewUser.Text = lstDelete.List(0)
cmdClear.Enabled = False
cmdDCancel.Enabled = False
Frame3.Enabled = False
End Sub

Private Sub cmdSave_Click()
'
If Len(txtNewUser.Text) < 2 Then
        MsgBox "To short Username", vbCritical
        txtNewUser.SetFocus
        Exit Sub
    End If
    If Len(txtPass.Text) < 8 Then
        MsgBox "Password is to short", vbCritical
        txtCPass.Text = ""
        txtPass.Text = ""
        txtPass.SetFocus
        Exit Sub
    End If
If p = True Then
    
    
        If txtPass.Text = txtCPass.Text Then
        If rstL.State = adStateOpen Then rstL.Close
            rstL.Open "select * from login  where username='" & StrConv(txtNewUser.Text, vbProperCase) & "'", cnn, adOpenKeyset, adLockOptimistic
            ' a nd pass='" & txtPass.Text & "'"
            If rstL.RecordCount > 0 Then
                MsgBox "User already exists", vbCritical
                cmdClear_Click
                Frame3.Enabled = True
            Else
'               If rstL.State = adStateOpen Then rstL.Close
'                     rstL.Open "select * from login where username<>'Proper(Administrator)'", cnn, adOpenKeyset, adLockOptimistic
'                        If rstL.RecordCount < 0 Then
'                             rstL.MoveLast
                rstL.AddNew
                rstL.Fields(0).Value = StrConv(txtNewUser.Text, vbProperCase)
                rstL.Fields(1).Value = txtPass.Text
                 
                rstL.Update
                cmdClear_Click
                MsgBox "New User Added", vbInformation
                        'Else
                         '   cmdClear_Click
                          '  MsgBox "Administrator Cannot Be Created", vbCritical
                Frame3.Enabled = True
                            
            End If
                
                    
        Else
             MsgBox "Password does not Match", vbCritical
             Frame1.Enabled = True
             txtPass.Text = ""
             txtCPass.Text = ""
             txtPass.SetFocus
            Exit Sub
        End If
Else
    Dim Rp As String
    If rstL.State = adStateOpen Then rstL.Close
            rstL.Open "select * from login  where username='" & txtNewUser.Text & "' ", cnn, adOpenKeyset, adLockOptimistic
        'and pass='" & txtPass.Text & "'"
    'If rstL.RecordCount > 0 Then
    If txtPass.Text = txtCPass.Text Then
        'If rstL.Fields(0).Value = StrConv(rstL.Fields(0).Value, vbUpperCase) Then
                'rstL.Fields(0).Value = StrConv(txtNewUser.Text, vbUpperCase)
        'Else
                rstL.Fields(0).Value = StrConv(txtNewUser.Text, vbProperCase)
        'End If
        Rp = txtPass.Text
        If Rp = rstL.Fields(1).Value Then
            MsgBox "Password is Same", vbInformation
            txtPass.Text = ""
            txtCPass.Text = ""
            txtPass.SetFocus
            Exit Sub
        End If
        rstL.Fields(1).Value = txtPass.Text
        rstL.Update
        MsgBox "Record Updated", vbInformation
        Frame3.Enabled = True
        cmdClear_Click
    Else
        MsgBox "Invalid Password", vbCritical
        Frame1.Enabled = True
        txtPass.SetFocus
        txtPass.Text = ""
        txtCPass.Text = ""
        Exit Sub
    End If
   
End If
'
listD
cmdSave.Enabled = False
cmdClear.Enabled = False
cmdCancel.Enabled = False
Frame1.Enabled = False
Frame2.Enabled = False
End Sub



Private Sub Command1_Click()
Dim i As VbMsgBoxResult

i = MsgBox("Do You Want To Exit", vbCritical + vbYesNo)
If i = vbYes Then
    If Frame1.Enabled = True Or Frame2.Enabled = True Then
        MsgBox "Session is Still Open,Please Cancel It", vbCritical
        Exit Sub
    End If
    Unload Me
Else
    Exit Sub
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
If Frame1.Enabled = True Then
    If KeyCode = 13 Then
        clearE
        cmdSave_Click
        Exit Sub
    End If
   
End If
End Sub

Private Sub Form_Load()
Frame1.Enabled = False
Frame2.Enabled = False
p = True
listD
cmdClear.Enabled = False
cmdSave.Enabled = False
cmdDelete.Enabled = False
cmdCancel.Enabled = False
cmdDCancel.Enabled = False
End Sub




Private Sub lstDelete_Click()
If del = False Then
Dim i As Integer
    For i = 0 To lstDelete.ListCount - 1
        If lstDelete.selected(i) = True Then
               txtNewUser.Text = lstDelete.List(i)
               txtPass.SetFocus
               lstDelete.selected(i) = False
               Exit For
        End If
    Next

End If
End Sub

Private Sub txtCPass_Change()
clearE
End Sub

Private Sub txtCPass_KeyPress(KeyAscii As Integer)

cmdSave.Enabled = True
End Sub

Private Sub txtNewUser_Change()
clearE
End Sub

Private Sub txtNewUser_KeyPress(KeyAscii As Integer)
If Len(txtNewUser.Text) <= 0 Then
    If KeyAscii = 95 Then
        KeyAscii = 0
    End If
End If
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 8 Or KeyAscii = 95) Then
    KeyAscii = 0
End If
End Sub


'Private Sub txtPass_LostFocus()
'If rstL.State = adStateOpen Then rstL.Close
 '   rstL.Open "select * from login  where username='" & txtNewUser.Text & "' and pass='" & txtPass.Text & "'", cnn, adOpenKeyset, adLockOptimistic
  '  If rstL.RecordCount < 1 Then
   '     txtCPass.Enabled = True
   ' End If
'End Sub
Private Sub txtPass_Change()
clearE
End Sub
