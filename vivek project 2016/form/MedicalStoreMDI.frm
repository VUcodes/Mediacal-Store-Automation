VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8085
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9465
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu frm 
      Caption         =   "Form"
      Begin VB.Menu measure 
         Caption         =   "Measure"
      End
      Begin VB.Menu category 
         Caption         =   "Category"
      End
      Begin VB.Menu company 
         Caption         =   "Company"
      End
      Begin VB.Menu supplier 
         Caption         =   "Supplier"
      End
      Begin VB.Menu medicnine 
         Caption         =   "Medicine"
      End
      Begin VB.Menu md 
         Caption         =   "Medicine Details"
      End
      Begin VB.Menu order 
         Caption         =   "Order"
      End
      Begin VB.Menu orderR 
         Caption         =   "Order Received"
      End
      Begin VB.Menu sale 
         Caption         =   "Sale"
      End
   End
   Begin VB.Menu rpt 
      Caption         =   "Reports"
   End
   Begin VB.Menu um 
      Caption         =   "User Manipulation"
      Begin VB.Menu add 
         Caption         =   "Add"
      End
      Begin VB.Menu chgPassword 
         Caption         =   "ChangePassword"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub category_Click()
frmcategory.Show
End Sub

Private Sub company_Click()
frmcompany.Show
End Sub

Private Sub md_Click()
frmmedicindetail.Show
End Sub

Private Sub MDIForm_Load()
MDIForm1.Caption = "WELCOME" & StrConv(user, vbUpperCase)
End Sub

Private Sub measure_Click()
frmunit.Show
End Sub

Private Sub medicnine_Click()
frmmedicine.Show
End Sub

Private Sub order_Click()
frmorder.Show
End Sub

Private Sub orderR_Click()
frmstock.Show
End Sub

Private Sub sale_Click()
frmsale.Show
End Sub

Private Sub supplier_Click()
frmsupplier.Show
End Sub
