VERSION 5.00
Begin VB.Form frmReverse 
   Caption         =   "ReverseOfaNumber"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5250
   LinkTopic       =   "Form2"
   ScaleHeight     =   3465
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Reverse"
      Height          =   2895
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      Begin VB.CommandButton cmdRev1 
         Caption         =   "Reverse"
         Height          =   495
         Left            =   1200
         TabIndex        =   3
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtnum 
         Height          =   615
         Left            =   2400
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Enter the number:"
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmReverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdRev1_Click()
Dim num As Integer
Dim re As Integer
Dim rev As Integer
num = Val(txtnum.Text)

re = 0
rev = 0

 While num > 0
    re = num Mod 10
    rev = (rev * 10) + re
    num = num \ 10
    
 Wend
 
MsgBox rev
End Sub
