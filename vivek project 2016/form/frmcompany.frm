VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmcompany 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8985
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmadd 
      Caption         =   "Manipulation"
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   2175
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame frmsave 
      Caption         =   "Confirmation"
      Height          =   1455
      Left            =   2280
      TabIndex        =   0
      Top             =   3600
      Width           =   2175
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdcanle 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5175
      Left            =   4560
      TabIndex        =   12
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9128
      _Version        =   393216
      BackColorBkg    =   -2147483633
   End
   Begin VB.Frame frmedit 
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4335
      Begin VB.ListBox lstcompany 
         Height          =   2085
         ItemData        =   "frmcompany.frx":0000
         Left            =   1560
         List            =   "frmcompany.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtc_name 
         Height          =   375
         Left            =   1560
         MaxLength       =   35
         TabIndex        =   1
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtc_id 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Category"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblc_name 
         Caption         =   "Company Name"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblc_id 
         Caption         =   "Company ID"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmcompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstAdd As New ADODB.Recordset ' recordset to add key record in company table
Dim rstCompany As New ADODB.Recordset ' record set for categery table
Dim rstaddnew As New ADODB.Recordset ' recoedset to add new record in company table
Dim rstEdit As New ADODB.Recordset ' recordset to edit record in company table
Dim rstcomcat As New ADODB.Recordset
Dim rstsearch As New ADODB.Recordset
Dim pk As Integer
Dim i As Integer, j As Integer
Dim ca As String, ccc() As Integer
Dim addb As Boolean, check1 As Boolean, check2 As Boolean  'variable to check the button add or edit

'
'
'

Private Sub bStyle(b As Boolean)
If b = True Then
    txtc_name.BorderStyle = vbFixedSingle
    txtc_name.BackColor = &H80000005
    '
    'lstcompany.BackColor = &H80000005
    '
    
Else

    txtc_name.BorderStyle = 0
    txtc_name.BackColor = &H8000000F
    '
    'lstcompany.BorderStyle = 0
    lstcompany.BackColor = &H8000000F
    '
End If
End Sub




Private Sub cmdadd_Click()
MSFlexGrid1.Enabled = False
frmedit.Enabled = True
frmSave.Enabled = True
frmadd.Enabled = False
additem
showgrid
addb = True
bStyle (True)
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
listshow
listshowedit
addb = False
check1 = False
check2 = False
bStyle (True)
ReDim ccc(0)
For i = 0 To lstcompany.ListCount - 1
    If lstcompany.selected(i) = True Then
        ReDim Preserve ccc(i)
        ccc(i) = i
    End If
Next
End Sub
Private Sub selected()
    txtc_name.SelStart = 0
    txtc_name.SelLength = Len(txtc_name.Text)
End Sub
Private Sub cmdSave_Click()
'If rstcomcat.State = adStateOpen Then rstcomcat.Close
If rstAdd.State = adStateOpen Then rstAdd.Close
If rstaddnew.State = adStateOpen Then rstaddnew.Close
rstAdd.Open "select * from company", cnn, adOpenKeyset, adLockOptimistic
rstaddnew.Open "select * from key_id", cnn, adOpenKeyset, adLockOptimistic
'rstcomcat.Open "select * from companycategory", cnn, adOpenKeyset, adLockOptimistic
If Len(txtc_name.Text) < 4 Then
    MsgBox ("Text Length Should Be Greater Than 4")
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
    
    
    If txtc_name.Text = "" Then
            MsgBox "Company Name Cannot Be  Empty", vbInformation, "Medical Store Automation"
           ' MsgBox ("company cannot be  empty")
            txtc_name.SetFocus
            selected
            Exit Sub
    End If
    
''''    If addb = False Then
''''    Dim X As Integer
''''       If rstedit.State = adStateOpen Then rstedit.Close
''''       rstedit.Open "select * from companycategory where companyid='" & txtc_id.Text & "'", cnn, adOpenKeyset, adLockOptimistic
''''       For i = 0 To lstcompany.ListCount - 1
''''            If lstcompany.selected(i) = True Then
''''                X = X + 1
''''                If rstadd.State = adStateOpen Then rstadd.Close
''''                rstadd.Open "select * from category where categoryid='" & rstedit.Fields("CategoryID").Value & "';", cnn, adOpenKeyset, adLockOptimistic
''''                If lstcompany.List(i) = rstadd.Fields("CategoryName").Value Then
''''                rstadd.MoveNext
''''                Else
''''                MsgBox "not equal"
''''                End If
''''
''''            End If
''''        If rstedit.EOF = False Then
''''            rstedit.MoveNext
''''        End If
''''       Next
''''       If X = rstedit.RecordCount Then
''''        MsgBox "You have made no change in list"
''''        txtc_name.SetFocus
''''            selected
''''        Exit Sub
''''       End If
''''    End If
If addb = False Then
    For i = 0 To UBound(ccc)
    If lstcompany.selected(ccc(i)) = False Then
        MsgBox "You deselect the selected category", vbCritical, "Medical Store Automation"
        lstcompany.SetFocus
        cmdedit_Click
        Exit Sub
    End If
    Next
End If
    
   
        If checklistvalidone = True Then
            check2 = False
        If rstAdd.State = adStateOpen Then rstAdd.Close
        rstAdd.Open "select * from company", cnn, adOpenKeyset, adLockOptimistic
        End If

        
    If (addb = True) Or (check2 = False And addb = False) Then
        rstAdd.MoveFirst
        While rstAdd.EOF <> True
            If StrConv(Trim(txtc_name.Text), vbProperCase) = rstAdd.Fields(1).Value Then
                MsgBox ("Company Already Exist"), vbInformation, "Medical Store Automation"
                txtc_name.SetFocus
                selected
                Exit Sub
            End If
            rstAdd.MoveNext
        Wend
        check1 = True
    Else
        If addb = False Then
            If checklistvalidone = True And check1 = False Then
                MsgBox "No Changes Made", vbInformation, "Medical Store Automation"
                txtc_name.SetFocus
                    selected
                Exit Sub
            End If
        End If
    End If
    
    
        
   
       
End If

Dim temp As Integer
    temp = 0
    For i = 0 To lstcompany.ListCount - 1
        If lstcompany.selected(i) = True Then
            temp = temp + 1
        End If
        
    Next
    
    
    
    If temp < 1 Then
        MsgBox ("Please Select Atleast One Category"), vbInformation, "Medical Store Automation"
            lstcompany.SetFocus
                        
            Exit Sub
    End If

If addb = True Then
    
    rstaddnew.MoveFirst
    rstaddnew.Fields(1).Value = rstaddnew.Fields(1).Value + 1
    rstaddnew.Update
    
    rstAdd.AddNew
    rstAdd.Fields(0).Value = Trim(Trim(Str(pk)))
    rstAdd.Fields(1).Value = StrConv(Trim(txtc_name.Text), vbProperCase)
    rstAdd.Update
    
    listaddfunc
    
'    If rstcomcat.State = adStateOpen Then rstcomcat.Close
'    rstcomcat.Open "select * from companycategory", cnn, adOpenKeyset, adLockOptimistic
'
'    For i = 0 To lstcompany.ListCount - 1
'        If lstcompany.selected(i) = True Then
'            If Not rstaddnew.Fields(2).Value = 0 Then
'                rstcomcat.MoveFirst
'            End If
'            rstcomcat.AddNew
'
'            rstcomcat.Fields(0).Value = rstaddnew.Fields(2).Value + 1
'            rstcomcat.Fields(1).Value = Trim(Trim(Str(pk)))
'            If rstedit.State = adStateOpen Then rstedit.Close
'            rstedit.Open "select * from category where categoryname='" & lstcompany.List(i) & "';", cnn, adOpenKeyset, adLockOptimistic
'
'            rstcomcat.Fields(2).Value = rstedit.Fields("CategoryID").Value
'            'MsgBox (rstedit.Fields("CategoryID").Value)
'            'rstcomcat.Fields(2).Value = i + 1
'
'            rstaddnew.Fields(2).Value = rstaddnew.Fields(2).Value + 1
'
'            rstedit.MoveNext
'            rstcomcat.Update
'            rstaddnew.Update
'        End If
'
'    Next
    
Else

    
    


    If rstEdit.State = adStateOpen Then rstEdit.Close
    rstEdit.Open "select * from company where companyid='" & txtc_id.Text & "'", cnn, adOpenKeyset, adLockOptimistic
    rstEdit.Fields(1).Value = StrConv(Trim(txtc_name.Text), vbProperCase)
    rstEdit.Update
    
 '   If rstEdit.State = adStateOpen Then rstEdit.Close
'    rstEdit.Open "delete * from companycategory where companyid='" & txtc_id.Text & "'", cnn, adOpenKeyset, adLockOptimistic
'    rstedit.Update
'    For i = 0 To rstedit.RecordCount - 1
'    rstedit.Delete
'    rstedit.MoveNext
'    Next
    
    If rstcomcat.State = adStateOpen Then rstcomcat.Close
    rstcomcat.Open "select * from companycategory", cnn, adOpenKeyset, adLockOptimistic
    
    
    If rstcomcat.State = adStateOpen Then rstcomcat.Close
    rstcomcat.Open "select * from companycategory", cnn, adOpenKeyset, adLockOptimistic
    Dim m%
   
    For i = (UBound(ccc) + 1) To lstcompany.ListCount - 1
        If lstcompany.selected(i) = True Then
            If Not rstaddnew.Fields(2).Value = 0 Then
                rstcomcat.MoveFirst
            End If
            rstcomcat.AddNew
            rstcomcat.Fields(0).Value = rstaddnew.Fields(2).Value + 1
            
                rstcomcat.Fields(1).Value = txtc_id.Text
           
            If rstEdit.State = adStateOpen Then rstEdit.Close
            rstEdit.Open "select * from category where categoryname='" & lstcompany.List(i) & "';", cnn, adOpenKeyset, adLockOptimistic
            
            rstcomcat.Fields(2).Value = rstEdit.Fields("CategoryID").Value
           
            
            rstaddnew.Fields(2).Value = rstaddnew.Fields(2).Value + 1
            
            rstEdit.MoveNext
            rstcomcat.Update
            rstaddnew.Update
        End If
    Next
    
    
    
End If
cmdcanle_Click
showData
MsgBox "Data Saved", vbOKOnly + vbInformation, "Medical Store Automation"
ReDim ccc(0)
End Sub
Private Function checklistvalidone() As Boolean
If addb = False Then
    Dim X As Integer
       If rstEdit.State = adStateOpen Then rstEdit.Close
       rstEdit.Open "select * from companycategory where companyid='" & txtc_id.Text & "'", cnn, adOpenKeyset, adLockOptimistic
       For i = 0 To lstcompany.ListCount - 1
            If lstcompany.selected(i) = True Then
                If rstEdit.EOF Then Exit For
                X = X + 1
                If rstAdd.State = adStateOpen Then rstAdd.Close
                rstAdd.Open "select * from category where categoryid='" & rstEdit.Fields("CategoryID").Value & "';", cnn, adOpenKeyset, adLockOptimistic
                If lstcompany.List(i) = rstAdd.Fields("CategoryName").Value Then
                    rstAdd.MoveNext
'                Else
'                    MsgBox "not equal"
                End If
                
            End If
'            If rstedit.EOF = False Then
'                rstedit.MoveNext
'            End If
       Next
            If X = rstEdit.RecordCount Then
                checklistvalidone = True
                Exit Function
            End If
End If
    checklistvalidone = False
    
End Function

Private Sub listaddfunc()
If rstcomcat.State = adStateOpen Then rstcomcat.Close
    rstcomcat.Open "select * from companycategory", cnn, adOpenKeyset, adLockOptimistic
    Dim m%
   
    For i = 0 To lstcompany.ListCount - 1
      
        
        If lstcompany.selected(i) = True Then
            If Not rstaddnew.Fields(2).Value = 0 Then
                rstcomcat.MoveFirst
            End If
            rstcomcat.AddNew
            
            rstcomcat.Fields(0).Value = rstaddnew.Fields(2).Value + 1
            If addb = True Then
                rstcomcat.Fields(1).Value = Trim(Trim(Str(pk)))
            Else
                rstcomcat.Fields(1).Value = txtc_id.Text
            End If
            If rstEdit.State = adStateOpen Then rstEdit.Close
            rstEdit.Open "select * from category where categoryname='" & lstcompany.List(i) & "';", cnn, adOpenKeyset, adLockOptimistic
            
            rstcomcat.Fields(2).Value = rstEdit.Fields("CategoryID").Value
            'MsgBox (rstedit.Fields("CategoryID").Value)
            'rstcomcat.Fields(2).Value = i + 1
            
            rstaddnew.Fields(2).Value = rstaddnew.Fields(2).Value + 1
            
            rstEdit.MoveNext
            rstcomcat.Update
            rstaddnew.Update
           End If
        
        
    Next
End Sub

Private Sub Form_Activate()
cmdadd.SetFocus
End Sub
Private Sub Form_Load()
frmedit.Enabled = False
frmSave.Enabled = False
check1 = False
check2 = False
showData
'lstcompany.Visible = False
addb = True
bStyle (False)
End Sub

Private Sub additem()
addb = True
clear_all
txtc_name.SetFocus
If rstaddnew.State = adStateOpen Then rstaddnew.Close
rstaddnew.Open "select * from key_id", cnn, adOpenKeyset, adLockOptimistic
pk = rstaddnew.Fields(1).Value + 1
If pk = 10000 Then
    MsgBox ("Company Limit Is 9999 ")
    frmedit.Enabled = False
    frmSave.Enabled = False
    frmadd.Enabled = True
    showData
    Exit Sub
End If

txtc_id.Text = pk

End Sub
Private Sub showData()
If rstcomcat.State = adStateOpen Then rstcomcat.Close
If rstCompany.State = adStateOpen Then rstCompany.Close
    Dim X As Integer, Y As Integer
    MSFlexGrid1.Visible = True
    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Cols = 2
    
    rstCompany.Open "select * from company", cnn, adOpenKeyset, adLockOptimistic
    
    txtc_id.Text = rstCompany.Fields(0).Value
    txtc_name.Text = rstCompany.Fields(1).Value
    
    MSFlexGrid1.Cols = rstCompany.Fields.Count
        For X = 0 To rstCompany.Fields.Count - 1
            MSFlexGrid1.TextMatrix(0, X) = rstCompany.Fields(X).Name
            MSFlexGrid1.ColWidth(X) = 1500
'
        Next
'
'            'MSFlexGrid1.CellHeight(1, 1) = 1500
            'MSFlexGrid1.CellWidth(1,1) = 500
    
    MSFlexGrid1.Height = 5000
    MSFlexGrid1.Width = 3400
    For X = 1 To rstCompany.RecordCount
        For Y = 0 To rstCompany.Fields.Count - 1
            MSFlexGrid1.TextMatrix(X, Y) = rstCompany.Fields(Y).Value
'            MSFlexGrid1.RowHeight(X) = 500
        Next
        rstCompany.MoveNext
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    Next
    
    listshow
    
'   If rstsearch.State = adStateOpen Then rstsearch.Close
'   lstcompany.Clear
'   rstsearch.Open "select categoryname from category", cnn, adOpenKeyset, adLockOptimistic
'   For i = 0 To rstsearch.RecordCount - 1
'        lstcompany.additem rstsearch.Fields("categoryname").Value
'        rstsearch.MoveNext
'   Next
   
'   If Not txtc_id.Text = "" Then
'   lstcompany.Clear
'        If rstcomcat.State = adStateOpen Then rstcomcat.Close
'        If rstedit.State = adStateOpen Then rstedit.Close
'        rstcomcat.Open "select * from companycategory where companyid='" & txtc_id.Text & "';", cnn, adOpenKeyset, adLockOptimistic
''        rstedit.Open "select * from category", cnn, adOpenKeyset, adLockOptimistic
'        For i = 0 To rstcomcat.RecordCount - 1
'            lstcompany.additem rstcomcat.Fields("categoryid").Value
'            lstcompany.selected(i) = True
'            rstcomcat.MoveNext
'        Next
'   End If
   
'   rstcomcat.Open "select * from key_id", cnn, adOpenKeyset, adLockOptimistic
'   For i = 0 To lstcompany.ListCount
'        If lstcompany.selected(i) = True Then
'            rstcomcat.MoveFirst
'            rstcomcat.AddNew
'
'            rstcomcat.Fields(0).Value = rstadd.Fields(2).Value + 1
'            rstcomcat.Fields(1).Value = Trim(Str(pk))
'            rstcomcat.Fields(0).Value = i
'
'            rstadd.Fields(2).Value 1 = rstcomcat.Fields(0).Value
'
'            rstcomcat.Update
'        End If
'
'    Next
   
End Sub
Private Sub clear_all()
txtc_id.Text = ""
txtc_name.Text = ""
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If frmedit.Enabled = True Then
'    MsgBox ("Please Complete the session")
'    Cancel = True
'End If
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'Dim a As Integer
'a = MsgBox("Do you want to Exit ?", vbYesNo, "EXIT")
'If a = 6 Then
'    Unload Me
'Else
'    Cancel = True
'End If
'
'End Sub
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
    listshow
End Sub


Private Sub txtc_name_KeyPress(KeyAscii As Integer)
'If Not ((KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 8) Then
'    KeyAscii = 0
'End If
''    keyascii = validation1
'If ((Len(txtc_name.Text) = 0) And KeyAscii = 32) Or (Right(txtc_name, 1) = " " And KeyAscii = 32) Then
'    KeyAscii = 0
'End If
'If Len(txtc_name.Text) = 0 Then
'    KeyAscii = key(KeyAscii, txtc_name.Text)
'Else

If Len(txtc_name.Text) = 0 Or txtc_name.SelStart = 0 Then
    If KeyAscii = 32 Then
        KeyAscii = 0
    End If
    KeyAscii = key(KeyAscii, txtc_name.Text)
Else
'''
    If KeyAscii = 8 Or KeyAscii = 45 Or KeyAscii = 32 Or KeyAscii = 46 Then
    
    Else
        KeyAscii = keyboth(KeyAscii, txtc_name.Text)
    End If
    If Len(txtc_name.Text) > 1 And (KeyAscii = 32 Or KeyAscii = 45 Or KeyAscii = 46) Then
        If Asc(Right((Left(txtc_name.Text, txtc_name.SelStart)), 1)) = KeyAscii Or Asc(Right((Left(txtc_name.Text, txtc_name.SelStart + 1)), 1)) = KeyAscii Then
        KeyAscii = 0
        End If
    End If
'    If Len(txtc_name.Text) > 1 Then
'        If Asc(Right((Left(txtc_name.Text, txtc_name.SelStart)), 1)) = KeyAscii Or Asc(Right((Left(txtc_name.Text, txtc_name.SelStart + 1)), 1)) = KeyAscii Then
'        KeyAscii = 0
'        End If
'    End If
End If
'    If KeyAscii = 46 Then
'        KeyAscii = 0
'    End If
End Sub

Private Sub txtcategory_Change()
If Len(txtcompany.Text) > 0 Then
    lstcompany.Visible = True
Else
    lstcompany.Visible = False
End If

'Dim fn As String
'fn = txtcategory
'If rstsearch.State = adStateOpen Then rstsearch.Close
'lstcompany.Clear
'rstsearch.Open "select categoryname from category where categoryname like'%" & fn & "%';", cnn, adOpenKeyset, adLockOptimistic


End Sub
Private Sub showgrid()
   If rstsearch.State = adStateOpen Then rstsearch.Close
   lstcompany.Clear
   rstsearch.Open "select categoryname from category", cnn, adOpenKeyset, adLockOptimistic
   For i = 0 To rstsearch.RecordCount - 1
        lstcompany.additem rstsearch.Fields("categoryname").Value
        rstsearch.MoveNext
   Next
End Sub

Private Sub listshow()
If Not txtc_id.Text = "" Then
   lstcompany.Clear
        If rstcomcat.State = adStateOpen Then rstcomcat.Close
        rstcomcat.Open "select * from companycategory where companyid='" & txtc_id.Text & "';", cnn, adOpenKeyset, adLockOptimistic
        For i = 0 To rstcomcat.RecordCount - 1
            If rstEdit.State = adStateOpen Then rstEdit.Close
            rstEdit.Open "select CategoryName from category where categoryid='" & rstcomcat.Fields("categoryid").Value & "';", cnn, adOpenKeyset, adLockOptimistic
            lstcompany.additem rstEdit.Fields("CategoryName").Value
            lstcompany.selected(i) = True
            rstcomcat.MoveNext
        Next
   End If
End Sub
Private Sub listshowedit()

'If rstcomcat.State = adStateOpen Then rstcomcat.Close
'rstcomcat.Open "select categoryid from companycategory where companyid='" & txtc_id.Text & "'", cnn, adOpenKeyset, adLockOptimistic
If rstEdit.State = adStateOpen Then rstEdit.Close
rstEdit.Open "select * from category where categoryid not in (select categoryid from companycategory where companyid='" & txtc_id.Text & "')", cnn, adOpenKeyset, adLockOptimistic
For i = 0 To rstEdit.RecordCount - 1
 lstcompany.additem rstEdit.Fields("CategoryName").Value
 rstEdit.MoveNext
Next
'Dim j As Integer
'For i = 0 To rstedit.RecordCount - 1
''    MsgBox (rstcomcat.Fields("CategoryID").Value)
''    MsgBox (rstedit.Fields("CategoryID").Value)
'    Dim x%
'    For j = i To rstedit.RecordCount - 1
'        If lstcompany.List(j) = rstedit.Fields("CategoryName").Value Then
'        rstedit.MoveNext
'        Exit For
'        Else
'            lstcompany.additem rstedit.Fields("CategoryName").Value
'        End If
'    Next
'
'Next
''rstedit.Move (rstcomcat.RecordCount - 1)
''For i = (rstcomcat.RecordCount) To (rstedit.RecordCount - rstcomcat.RecordCount) - 2
''lstcompany.additem rstedit.Fields("CategoryName").Value
''    rstedit.MoveNext
''Next

   
End Sub
