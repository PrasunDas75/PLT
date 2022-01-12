VERSION 5.00
Begin VB.Form frmPattern11 
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes1 
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   6735
   End
   Begin VB.CommandButton cmdStart1 
      Caption         =   "Start"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtN 
      Height          =   975
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "N :"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmPattern11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart1_Click()
Dim i, n As Integer
Dim s As String

n = Val(txtN.Text)

For i = 1 To n
    If i Mod 2 = 0 Then
        s = s & Str(i * i)
    End If
Next

txtRes1.Text = s

End Sub
