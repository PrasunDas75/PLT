VERSION 5.00
Begin VB.Form frmEvenOdd 
   Caption         =   "Form2"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12270
   LinkTopic       =   "Form2"
   ScaleHeight     =   7920
   ScaleWidth      =   12270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtNumber 
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Check if given No. is Odd or Even"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label lbEnter 
      Caption         =   "Enter The Number"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "frmEvenOdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
