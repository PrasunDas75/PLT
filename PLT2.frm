VERSION 5.00
Begin VB.Form frmSwapNumber 
   Caption         =   "SwapNumber"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwap 
      Caption         =   "Swap"
      Height          =   855
      Left            =   2640
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtRes2 
      Height          =   735
      Left            =   4680
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtRes1 
      Height          =   735
      Left            =   4680
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtN2 
      Height          =   735
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtxN1 
      Height          =   735
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "B:"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "A:"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmSwapNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSwap_Click()
Dim a, b As Integer
a = Val(txtxN1.Text)
b = Val(txtN2.Text)

Dim t As Integer

t = a
a = b
b = t
txtRes1.Text = a
txtRes2.Text = b

End Sub
