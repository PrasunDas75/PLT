VERSION 5.00
Begin VB.Form frmSumOdd 
   Caption         =   "Form2"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8220
   LinkTopic       =   "Form2"
   ScaleHeight     =   5595
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSum 
      Height          =   855
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3000
      Width           =   3735
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtN 
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   3615
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
