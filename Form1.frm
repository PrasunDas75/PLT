VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      Height          =   1335
      Left            =   3240
      TabIndex        =   4
      Text            =   "Result"
      Top             =   4800
      Width           =   3855
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   975
      Left            =   3360
      TabIndex        =   3
      Top             =   3000
      Width           =   3255
   End
   Begin VB.TextBox txtRate 
      Height          =   975
      Left            =   6960
      TabIndex        =   2
      Text            =   "Rate"
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtTime 
      Height          =   975
      Left            =   3240
      TabIndex        =   1
      Text            =   "Time"
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtPrinciple 
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Text            =   "Principle"
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text4_Change()

End Sub

Private Sub txtCalculate_Click()

End Sub

Private Sub cmdResult_Click()


End Sub

Private Sub cmdCalculate_Click()
Dim P As Integer
Dim time As Integer
Dim rate As Integer
Dim si As Integer
    P = Val(txtPrinciple.Text)
    time = Val(txtTime.Text)
    rate = Val(txtRate.Text)
    si = (P * time * rate) / 100
    txtResult.Text = si
End Sub

