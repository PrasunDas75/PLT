VERSION 5.00
Begin VB.Form frmPattern18 
   Caption         =   "Form2"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7425
   LinkTopic       =   "Form2"
   ScaleHeight     =   6825
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes1 
      Height          =   615
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2520
      Width           =   4695
   End
   Begin VB.CommandButton cmdGen1 
      Caption         =   "Generate"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtN 
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "frmPattern18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGen1_Click()
Dim N1 As Integer
Dim N2 As Integer
Dim N3 As Integer

Dim i As Integer
Dim j As Integer
Dim r As Integer
Dim s As String

r = Val(txtN.Text)

N1 = 0
N2 = 1


txtRes1.Text = Str(N2)

For i = 2 To r
    N3 = N1 + N2
    If N3 <= r Then
     s = s & Str(N3)
    End If
    N1 = N2
    N2 = N3
Next

txtRes1.Text = txtRes1.Text & s

End Sub
