VERSION 5.00
Begin VB.Form frmedit 
   Caption         =   "Form2"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   2520
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CANCLE"
      Height          =   735
      Left            =   2640
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      Height          =   735
      Left            =   720
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "CATEGORY_NAME"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "CATEGORY_ID"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
