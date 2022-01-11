VERSION 5.00
Begin VB.Form frmFactorial 
   Caption         =   "Form2"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7035
   LinkTopic       =   "Form2"
   ScaleHeight     =   4740
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes1 
      Height          =   735
      Left            =   1920
      TabIndex        =   3
      Text            =   "Factorial"
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton cmdFact 
      Caption         =   "Factorial"
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtNum1 
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the number"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmFactorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFact_Click()
Dim i, f, n As Integer

n = Val(txtNum1.Text)
i = 1
f = 1

If n = 0 Then
    txtRes1.Text = 1
ElseIf n < 0 Then
    txtRes1.Text = "Impossible!"
Else
    While i <= n
        f = f * i
        i = i + 1
        
    Wend
    txtRes1.Text = f
End If

End Sub
