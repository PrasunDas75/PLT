VERSION 5.00
Begin VB.Form frmSumOdd 
   Caption         =   "SumOfOddNumbers"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7380
   LinkTopic       =   "Form2"
   ScaleHeight     =   3555
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSum 
      Height          =   735
      Left            =   1920
      TabIndex        =   2
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtN 
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Enter N:"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmSumOdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalc_Click()
Dim N, sum, i As Double
N = Val(txtN.Text)
sum = 0
For i = 0 To N
    If i Mod 2 <> 0 Then
    sum = sum + i
    End If
Next
txtSum.Text = sum
End Sub
