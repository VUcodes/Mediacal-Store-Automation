VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmsale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnsale 
      Caption         =   "Sale"
      Height          =   375
      Left            =   9240
      TabIndex        =   29
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton btnr 
      Caption         =   "Remove"
      Height          =   435
      Left            =   5040
      TabIndex        =   28
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton btnadd 
      Caption         =   "Add"
      Height          =   435
      Left            =   3960
      TabIndex        =   27
      Top             =   6360
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6135
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   10821
      _Version        =   393216
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1560
      TabIndex        =   25
      Top             =   3480
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1560
      TabIndex        =   24
      Top             =   2880
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1560
      TabIndex        =   23
      Top             =   2400
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   22
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   1560
      TabIndex        =   26
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1560
      TabIndex        =   21
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblgt 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label lblt 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblm 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblp 
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lbln 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lbld 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lbli 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Grand Total"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Total"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Mfg. Date"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Price"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Batch No."
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Weight"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Category Name"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Company Name"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Medicine Nmae"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Customer Name"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Sale Date"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Sale Id"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmsale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsta As New ADODB.Recordset
Dim rstb As New ADODB.Recordset
Private Sub Form_Activate()
If rsta.State = adStateOpen Then rsta.Close
rsta.Open "Select * from key_id", cnn, adOpenKeyset, adLockOptimistic
lbli = rsta.Fields(9).Value + 1
lbld = Format(Date, "dd-mmm-yyyy")
lblt.Caption = 0
lblgt.Caption = 0
End Sub

