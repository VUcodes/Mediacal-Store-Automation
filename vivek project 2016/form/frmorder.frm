VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmorder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnsubmit 
      Caption         =   "Place Order"
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton btnadd 
      Caption         =   "Add"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton btndelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   2040
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2990
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3201
      _Version        =   393216
   End
   Begin VB.ComboBox cmbsn 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.ComboBox cmbcn 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtsearch 
      Height          =   405
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtorderdate 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   405
      Left            =   1440
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtorderno 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   405
      Left            =   1440
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Search"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Supplier Name"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Company Name"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Order Date"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Order No."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstmedicine As New ADODB.Recordset
Dim rstC As New ADODB.Recordset
Dim pk%

Private Sub btnadd_Click()
Dim j%, i%
For i = 1 To MSFlexGrid1.Rows - 1
            If MSFlexGrid1.TextMatrix(i, 7) = "" Then
                    j = j + 1
            Else
                
                If CInt(MSFlexGrid1.TextMatrix(i, 7)) + CInt(MSFlexGrid1.TextMatrix(i, 6)) < CInt(MSFlexGrid1.TextMatrix(i, 5)) Then
                    MsgBox "Reorder level must be smaller than sum of quantity and stock", vbCritical, "Medical Store Automation"
                    Exit Sub
                End If
               
            End If
            
Next
    If j = MSFlexGrid1.Rows - 1 Then
            MsgBox "Please insert quantity", vbCritical, "Medical Store Automation"
            Exit Sub
    End If

    
    
For i = 1 To MSFlexGrid1.Rows - 1
            If Not MSFlexGrid1.TextMatrix(i, 7) = "" Then
            
            For j = 1 To MSFlexGrid2.Rows - 1
                If MSFlexGrid2.TextMatrix(j, 0) = MSFlexGrid1.TextMatrix(i, 0) Then
                MsgBox "Quantity of this medicine already exist", vbCritical, "Medical Store Automation"
                MSFlexGrid1.TextMatrix(i, 7) = ""
                Exit Sub
                End If
            Next
            
                  MSFlexGrid2.Rows = MSFlexGrid2.Rows + 1
                  MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, 0) = MSFlexGrid1.TextMatrix(i, 0)
                  MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, 1) = MSFlexGrid1.TextMatrix(i, 1)
                  MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, 2) = MSFlexGrid1.TextMatrix(i, 2)
                  MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, 3) = MSFlexGrid1.TextMatrix(i, 4)
                  MSFlexGrid2.TextMatrix(MSFlexGrid2.Rows - 1, 4) = MSFlexGrid1.TextMatrix(i, 7)
            End If
Next
For i = 1 To MSFlexGrid1.Rows - 1
             MSFlexGrid1.TextMatrix(i, 7) = ""
Next

End Sub

Private Sub btndelete_Click()
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
End Sub

Private Sub btnsubmit_Click()
If MSFlexGrid2.Row < 1 Then
    MsgBox "Atleast one order is required", vbCritical, "Medical Store Automation"
    Exit Sub
End If
If cmbsn.Text = "Supplier Id-Name-Contact" Then
    MsgBox "Please select the supplire", vbCritical, "Medical Store Automation"
    cmbsn.SetFocus
    Exit Sub
End If


If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select * from orders", cnn, adOpenKeyset, adLockOptimistic
rstC.AddNew
rstC.Fields(0) = txtorderno
rstC.Fields(1) = txtorderdate
rstC.Fields(2) = Left(cmbsn.Text, InStr(1, cmbsn.Text, "-") - 1)
rstC.Update

If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select * from key_id", cnn, adOpenKeyset, adLockOptimistic
rstC.Fields(7).Value = txtorderno.Text
rstC.Update

Dim i%
If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select * from orderdetail", cnn, adOpenKeyset, adLockOptimistic
For i = 1 To MSFlexGrid2.Rows - 1
rstC.AddNew
rstC.Fields(0) = txtorderno
rstC.Fields(1) = MSFlexGrid2.TextMatrix(i, 0)
rstC.Fields(2) = Left(cmbsn.Text, InStr(1, cmbsn.Text, "-") - 1)
rstC.Update
Next
cmbcn.Clear

Form_Activate
cmbsn.ListIndex = 0
txtsearch.Text = ""

MsgBox "Data Saved", vbOKOnly + vbInformation, "Medical Store Automation"
End Sub

Private Sub cmbcn_Click()
cmbsn.Clear
If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select * from qrysupplire where CompanyName='" & cmbcn.Text & "'", cnn, adOpenKeyset, adLockOptimistic
cmbsn.additem "Supplier Id-Name-Contact"
cmbsn.ListIndex = 0
'If rstC.RecordCount = 0 Then
'    MsgBox "No Supplier found", vbInformation, "Medical Store Automation"
'    Exit Sub
'End If
Dim i%
For i = 0 To rstC.RecordCount - 1
    cmbsn.additem rstC.Fields(0).Value & "-" & rstC.Fields(1).Value & "-" & rstC.Fields(2).Value
    rstC.MoveNext
Next
showm
showmn
End Sub

Private Sub Form_Activate()

Dim i As Integer
If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select companyname from company", cnn, adOpenKeyset, adLockOptimistic
For i = 0 To rstC.RecordCount - 1
    cmbcn.additem rstC.Fields(0).Value
    rstC.MoveNext
Next
cmbcn.ListIndex = 0

If rstC.State = adStateOpen Then rstC.Close
rstC.Open "select * from key_id", cnn, adOpenKeyset, adLockOptimistic
pk = rstC.Fields(7).Value + 1
If pk = 10000000 Then
    MsgBox ("Company Limit Is 9999999 ")
    Exit Sub
End If
txtorderno.Text = pk

txtorderdate.Text = Format(Date, "dd-mmm-yyyy")
txtsearch.SetFocus
showm
showmn
End Sub

Public Function showm()
If rstmedicine.State = adStateOpen Then rstmedicine.Close
rstmedicine.Open "select * from qryordernew where CompanyName='" & cmbcn.Text & "'", cnn, adOpenKeyset, adLockOptimistic

MSFlexGrid1.Clear
Dim i%, j%
MSFlexGrid1.Cols = rstmedicine.Fields.Count + 1
MSFlexGrid1.Rows = rstmedicine.RecordCount + 1
For i = 0 To rstmedicine.Fields.Count - 1
    MSFlexGrid1.TextMatrix(0, i) = rstmedicine.Fields(i).Name
Next
    MSFlexGrid1.TextMatrix(0, i) = "Quantity"
For i = 1 To rstmedicine.RecordCount
        For j = 0 To rstmedicine.Fields.Count - 1
            MSFlexGrid1.TextMatrix(i, j) = rstmedicine.Fields(j).Value
        Next
rstmedicine.MoveNext
Next
End Function
Public Function showmn()
If rstmedicine.State = adStateOpen Then rstmedicine.Close
rstmedicine.Open "select * from qryaddqty", cnn, adOpenKeyset, adLockOptimistic

MSFlexGrid2.Clear
Dim i%, j%
MSFlexGrid2.Cols = rstmedicine.Fields.Count + 1
MSFlexGrid2.Rows = 1
For i = 0 To rstmedicine.Fields.Count - 1
    MSFlexGrid2.TextMatrix(0, i) = rstmedicine.Fields(i).Name
Next
    MSFlexGrid2.TextMatrix(0, i) = "Quantity"
'For i = 1 To rstmedicine.RecordCount
'        For j = 0 To rstmedicine.Fields.Count - 1
'            MSFlexGrid1.TextMatrix(i, j) = rstmedicine.Fields(j).Value
'        Next
'rstmedicine.MoveNext
'Next
End Function

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
Dim m As String
m = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)

If m = "" And KeyAscii = 48 Then
    KeyAscii = 0
    Exit Sub
End If

If KeyAscii = 8 Then
    If m = "" Then
        KeyAscii = 0
        Exit Sub
    Else
        KeyAscii = 0
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) = Left(m, Len(m) - 1)
        
    End If
Else
    KeyAscii = keynum(KeyAscii, MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7))
End If
If Not KeyAscii = 0 Then
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) + Chr(KeyAscii)
End If

m = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)
If Len(m) > 0 Then
    If Not (CInt(m) > 0 And CInt(m) <= 255) Then
        MsgBox ("Quantity should 1 to 255"), vbCritical, "Medical Store Automation"
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) = Left(m, Len(m) - 1)
    End If
End If
End Sub

Private Sub txtsearch_Change()
If Len(txtsearch.Text) > 0 Then

    If rstmedicine.State = adStateOpen Then rstmedicine.Close
    rstmedicine.Open "select * from qryorder where CompanyName='" & cmbcn.Text & "'and medicinename like '%" & txtsearch.Text & "%'", cnn, adOpenKeyset, adLockOptimistic
    
    MSFlexGrid1.Clear
    Dim i%, j%
    MSFlexGrid1.Cols = rstmedicine.Fields.Count
    MSFlexGrid1.Rows = rstmedicine.RecordCount + 1
    For i = 0 To rstmedicine.Fields.Count - 1
        MSFlexGrid1.TextMatrix(0, i) = rstmedicine.Fields(i).Name
    Next
    For i = 1 To rstmedicine.RecordCount
            For j = 0 To rstmedicine.Fields.Count - 1
                MSFlexGrid1.TextMatrix(i, j) = rstmedicine.Fields(j).Value
            Next
    rstmedicine.MoveNext
    Next
Else
    showm
End If
  

End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
If Len(txtsearch.Text) = 0 Or txtsearch.SelStart = 0 Then
'    If KeyAscii = 32 Then
'        KeyAscii = 0
'    End If
    KeyAscii = keyboth(KeyAscii, txtsearch.Text)
Else
'''
    If KeyAscii = 8 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 46 Or KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 44 Then
    
    Else
        KeyAscii = keyboth(KeyAscii, txtsearch.Text)
    End If
    If Len(txtsearch.Text) > 1 And (KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 46 Or KeyAscii = 32 Or KeyAscii = 44) Then
    Dim aa1%, aa2%
    aa1 = Asc(Right((Left(txtsearch.Text, txtsearch.SelStart)), 1))
    aa2 = Asc(Right((Left(txtsearch.Text, txtsearch.SelStart + 1)), 1))
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
