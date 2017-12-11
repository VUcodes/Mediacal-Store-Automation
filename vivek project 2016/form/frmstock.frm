VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmstock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Receive"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdor 
      Caption         =   "Order Received"
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   6960
      Width           =   1335
   End
   Begin VB.ComboBox cmborder 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   120
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   3975
      Left            =   0
      TabIndex        =   2
      Top             =   2880
      Width           =   7095
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   42427
      End
      Begin VB.CommandButton cmdrem 
         Caption         =   "Remove"
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton btnadd 
         Caption         =   "Add"
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3625
         _Version        =   393216
      End
      Begin VB.TextBox txtbno 
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtprice 
         Height          =   375
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Batch No."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Mfg. Date"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Price"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2055
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3625
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Order No."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstmedicine As New ADODB.Recordset
Dim rstC As New ADODB.Recordset
Dim rstk As New ADODB.Recordset
Private Function border1(b As Integer)
        txtbno.BorderStyle = b
        txtprice.BorderStyle = b
        If b = 1 Then
             txtbno.BackColor = vbWhite
             txtprice.BackColor = vbWhite
             
        Else
            txtbno.BackColor = Me.BackColor
             txtprice.BackColor = Me.BackColor
        End If
End Function

Private Sub btnadd_Click()
Dim i%
If txtprice.Text = "" Then
    MsgBox "Price cannot be empty", vbCritical, "Medical Store Automation"
    txtprice.SetFocus
    Exit Sub
End If

If Not (txtprice.Text >= 1 And txtprice.Text <= 1000000) Then
    MsgBox "Price must exist between 1 to 10,00,000", vbCritical, "Medical Store Automation"
    txtprice.SetFocus
    Exit Sub
End If

If DateDiff("d", Format(DateAdd("d", -10, Date), "dd-mmm-yyyy"), DTPicker1.Value) > 1 Then
    MsgBox "Invalid Mfg date", vbCritical, "Medical Store Automation"
    DTPicker1.SetFocus
    Exit Sub
End If

If txtbno.Text = "" Then
    MsgBox "Batch no cannot be empty", vbCritical, "Medical Store Automation"
    txtbno.SetFocus
    Exit Sub
End If

If Not IsNumeric(Right(txtbno, 1)) Then
    MsgBox "Batch no is incorrect", vbCritical, "Medical Store Automation"
    txtbno.SetFocus
    Exit Sub
End If

If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select * from qryexpc where Mdetail_ID='" & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) & "'", cnn, adOpenKeyset, adLockBatchOptimistic

If (DateDiff("m", Date, DateAdd("m", rstC.Fields(1).Value, DTPicker1.Value)) > 6) Then
    MsgBox "Medicine Will expire Before 6 Months in stock", vbCritical, "Medicl Stoere Automation"
    Exit Sub
End If

If MSFlexGrid2.Rows > 1 Then
    For i = 1 To MSFlexGrid2.Rows - 1
        If MSFlexGrid2.TextMatrix(i, 0) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0) Then
            MsgBox "Record already exist of this medicine", vbCritical, "Medical Store Automation"
            Exit Sub
        End If
        If txtbno.Text = MSFlexGrid2.TextMatrix(i, 6) Then
            MsgBox "Batch No. already exist", vbCritical, "Medical Store Automation"
            Exit Sub
        End If
    Next
End If

If rstk.State = adStateOpen Then rstk.Close
rstk.Open "select * from stock", cnn, adOpenKeyset, adLockOptimistic
Dim j%
For j = 0 To rstk.RecordCount - 1
    If txtbno.Text = rstk.Fields(4).Value Then
    MsgBox "Batch No. already exist in data base", vbCritical, "Medical Store Automation"
    Exit Sub
    End If
    rstk.MoveNext
Next

MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1

For i = 0 To MSFlexGrid1.Cols - 1
    
    MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, i) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, i)
Next
MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, 4) = txtprice.Text
MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, 5) = Format(DTPicker1.Value, "dd-mmm-yyyy")
MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, 6) = txtbno.Text

txtbno.Text = ""
txtprice.Text = ""
txtprice.SetFocus

If MSFlexGrid2.Rows > 1 Then
    cmdor.Enabled = True
    cmdrem.Enabled = True
Else
    cmdor.Enabled = False
    cmdrem.Enabled = False
End If

End Sub

Private Sub cmborder_Click()
showm
showmn
'MSFlexGrid2.Rows = 1
End Sub

Private Sub cmdor_Click()
If Not MSFlexGrid2.Rows = MSFlexGrid1.Rows Then
    MsgBox "Place all the orders", vbCritical, "Medical Store Automation"
    Exit Sub
End If
Dim i%
If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select * from orders where orderno='" & Left(cmborder.Text, InStr(cmborder.Text, ",") - 1) & "'", cnn, adOpenKeyset, adLockOptimistic

    rstC.Fields(3).Value = True
    rstC.Update
    rstC.MoveNext


If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select * from stock", cnn, adOpenKeyset, adLockOptimistic
For i = 1 To MSFlexGrid2.Rows - 1
    If rstk.State = adStateOpen Then rstk.Close
    rstk.Open "select * from key_id", cnn, adOpenKeyset, adLockOptimistic
    
    rstk.Fields(8).Value = rstk.Fields(8).Value + 1
    rstk.Update
    rstC.AddNew
    rstC.Fields(0).Value = rstk.Fields(8).Value
    rstC.Fields(1).Value = MSFlexGrid2.TextMatrix(i, 3)
    rstC.Fields(2).Value = MSFlexGrid2.TextMatrix(i, 4)
    rstC.Fields(3).Value = MSFlexGrid2.TextMatrix(i, 5)
    rstC.Fields(4).Value = MSFlexGrid2.TextMatrix(i, 6)
    rstC.Fields(5).Value = MSFlexGrid2.TextMatrix(i, 0)
    
    rstC.Update
Next
MsgBox "Order received", vbInformation, "Medical Store Automation"
cmborder.Clear
Form_Activate
txtbno.Text = ""
txtprice.Text = ""
cmborder.Enabled = True
Frame1.Enabled = False
border1 0
End Sub

Private Sub cmdrem_Click()
Dim i%
For i = MSFlexGrid2.Row To MSFlexGrid2.Rows - 2
    MSFlexGrid2.TextMatrix(i, 0) = MSFlexGrid2.TextMatrix(i + 1, 0)
    MSFlexGrid2.TextMatrix(i, 1) = MSFlexGrid2.TextMatrix(i + 1, 1)
    MSFlexGrid2.TextMatrix(i, 2) = MSFlexGrid2.TextMatrix(i + 1, 2)
    MSFlexGrid2.TextMatrix(i, 3) = MSFlexGrid2.TextMatrix(i + 1, 3)
    MSFlexGrid2.TextMatrix(i, 4) = MSFlexGrid2.TextMatrix(i + 1, 4)
Next

If MSFlexGrid2.Rows = 1 Then
MsgBox "No Record to Remove", vbCritical, "Medical Store Automation"
Exit Sub
End If
MSFlexGrid2.Rows = MSFlexGrid2.Rows - 1
If MSFlexGrid2.Rows > 1 Then
    cmdor.Enabled = True
    cmdrem.Enabled = True
Else
    cmdor.Enabled = False
    cmdrem.Enabled = False
End If
End Sub

Private Sub Form_Activate()
DTPicker1.MinDate = Format(DateAdd("yyyy", -2, Date), "dd-mmm-yyyy")
DTPicker1.Value = Format(DateAdd("yyyy", -2, Date), "dd-mmm-yyyy")
'
If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select * from qryordersup", cnn, adOpenKeyset, adLockOptimistic

Dim i%
For i = 0 To rstC.RecordCount - 1
cmborder.additem rstC.Fields(0).Value & "," & Format(rstC.Fields(1).Value, "dd-mmm-yyyy") & "," & rstC.Fields(2).Value
rstC.MoveNext
Next
cmborder.ListIndex = 0
showm
showmn
border1 0
End Sub

Public Function showm()
If rstmedicine.State = adStateOpen Then rstmedicine.Close
rstmedicine.Open "select * from qryor where orderno='" & Left(cmborder.Text, InStr(cmborder.Text, ",") - 1) & "'", cnn, adOpenKeyset, adLockOptimistic

MSFlexGrid1.Clear
Dim i%, j%
MSFlexGrid1.Cols = rstmedicine.Fields.Count - 1
MSFlexGrid1.Rows = rstmedicine.RecordCount + 1
For i = 1 To rstmedicine.Fields.Count - 1
    MSFlexGrid1.TextMatrix(0, i - 1) = rstmedicine.Fields(i).Name
    If i = 3 Then
        MSFlexGrid1.TextMatrix(0, i - 1) = "Weight"
    End If
Next
For i = 1 To rstmedicine.RecordCount
        For j = 1 To rstmedicine.Fields.Count - 1
            MSFlexGrid1.TextMatrix(i, j - 1) = rstmedicine.Fields(j).Value
        Next
rstmedicine.MoveNext
Next
If MSFlexGrid2.Rows > 1 Then
    cmdor.Enabled = True
    cmdrem.Enabled = True
Else
    cmdor.Enabled = False
    cmdrem.Enabled = False
End If
End Function
Public Function showmn()
If rstmedicine.State = adStateOpen Then rstmedicine.Close
rstmedicine.Open "select * from qryor", cnn, adOpenKeyset, adLockOptimistic

MSFlexGrid2.Clear
Dim i%, j%
MSFlexGrid2.Cols = rstmedicine.Fields.Count + 2
MSFlexGrid2.Rows = 1
For i = 1 To rstmedicine.Fields.Count - 1
    
    MSFlexGrid2.TextMatrix(0, i - 1) = rstmedicine.Fields(i).Name
    If i = 3 Then
        MSFlexGrid2.TextMatrix(0, i - 1) = "Weight"
    End If
Next
MSFlexGrid2.TextMatrix(0, 4) = "Price"
MSFlexGrid2.TextMatrix(0, 5) = "MFG Date"
MSFlexGrid2.TextMatrix(0, 6) = "Batch No."
'For i = 1 To rstmedicine.RecordCount
'        For j = 1 To rstmedicine.Fields.Count - 1
'            MSFlexGrid2.TextMatrix(i, j - 1) = rstmedicine.Fields(j).Value
'        Next
'rstmedicine.MoveNext
'Next

End Function
'
'Private Sub Form_Paint()
'If MSFlexGrid2.Rows > 1 Then
'    cmdor.Enabled = True
'    cmdrem.Enabled = True
'Else
'    cmdor.Enabled = False
'    cmdrem.Enabled = False
'End If
'End Sub

Private Sub MSFlexGrid1_Click()
If MSFlexGrid1.Rows > 2 Then

Frame1.Enabled = True
txtprice.SetFocus
border1 1
MSFlexGrid1.BackColorSel = vbCyan
'MSFlexGrid1.CellBackColor = "blue"
End If
End Sub

Private Sub txtbno_KeyPress(KeyAscii As Integer)
If Len(txtbno.Text) = 0 Or txtbno.SelStart = 0 Then
    If KeyAscii = 32 Then
        KeyAscii = 0
    End If
    KeyAscii = key(KeyAscii, txtbno.Text)
Else
    If KeyAscii = 32 Then
        KeyAscii = 0
    End If
    
'    If InStr(1, Text1.Text, "-") = 1 And KeyAscii = 45 Then
'    KeyAscii = 0
'    End If
     
    
    If Len(txtbno.Text) < 2 Or txtbno.SelStart < 2 Then
        KeyAscii = key(KeyAscii, txtbno.Text)
    End If
    
        
   ' If KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 32 Then
    If (KeyAscii = 45 Or KeyAscii = 8) Then
    
    Else
        KeyAscii = keyboth(KeyAscii, txtbno.Text)
    End If
    If Len(txtbno.Text) > 1 And (KeyAscii = 45) Then
        If Asc(Right((Left(txtbno.Text, txtbno.SelStart)), 1)) = KeyAscii Or Asc(Right((Left(txtbno.Text, txtbno.SelStart + 1)), 1)) = KeyAscii Then
        KeyAscii = 0
        End If
    End If
    ''''
    ''''
'    If Len(txtEmail.Text) > 1 And (KeyAscii = 47 Or KeyAscii = 46 Or KeyAscii = 41 Or KeyAscii = 45 Or KeyAscii = 40) Then
'    Dim aa1%, aa2%
'    aa1 = Asc(Right((Left(txtbno.Text, txtbno.SelStart)), 1))
'    aa2 = Asc(Right((Left(txtbno.Text, txtbno.SelStart + 1)), 1))
'        If aa1 = KeyAscii Or aa2 = KeyAscii Then
'        KeyAscii = 0
'        End If
'        If ((aa1 = 47 Or aa1 = 41 Or aa1 = 46 Or aa1 = 45 Or aa1 = 40) And (KeyAscii = 47 Or KeyAscii = 46 Or KeyAscii = 41 Or KeyAscii = 45 Or KeyAscii = 40)) Then
'                KeyAscii = 0
'        End If
'        If ((aa2 = 47 Or aa2 = 41 Or aa2 = 46 Or aa1 = 45 Or aa1 = 40) And (KeyAscii = 47 Or KeyAscii = 46 Or KeyAscii = 41 Or KeyAscii = 45 Or KeyAscii = 40)) Then
'                KeyAscii = 0
'        End If
'    End If
'    ''''
'    ''''
End If
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
If txtprice.SelStart = 0 And KeyAscii = 48 Then
    KeyAscii = 0
End If
If KeyAscii = 8 Then

Else
KeyAscii = keynum(KeyAscii, txtprice.Text)
End If
End Sub
