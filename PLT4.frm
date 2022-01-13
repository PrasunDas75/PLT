VERSION 5.00
Begin VB.Form frmDecimalSeparator 
   Caption         =   "DecimalSeparator"
   ClientHeight    =   3525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4500
   LinkTopic       =   "Form2"
   ScaleHeight     =   3525
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes2 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtRes1 
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdSeparate 
      Caption         =   "Separate"
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtNum1 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Double Value:"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   600
      Width           =   1935
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
