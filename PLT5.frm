VERSION 5.00
Begin VB.Form frmDecimalSeparator 
   Caption         =   "Form2"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8940
   LinkTopic       =   "Form2"
   ScaleHeight     =   5940
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes2 
      Height          =   615
      Left            =   4200
      TabIndex        =   3
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtRes1 
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton cmdSeparate 
      Caption         =   "Display"
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtNum1 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "frmDecimalSeparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text2_Change()

End Sub

Private Sub cmdSeparate_Click()
Dim num As Double
num = Val(txtNum1.Text)
txtRes1.Text = Int(num)
txtRes2.Text = num - Int(num)

End Sub
