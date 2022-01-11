VERSION 5.00
Begin VB.Form frmRevString 
   Caption         =   "Form2"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5670
   LinkTopic       =   "Form2"
   ScaleHeight     =   5370
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes1 
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   3720
      Width           =   4935
   End
   Begin VB.CommandButton cmdRev 
      Caption         =   "Reverse"
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtStr1 
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label enter 
      Caption         =   "Enter the string"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmRevString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRev_Click()

Dim len1 As Integer, i As Integer
Dim s As String
Dim res As String
 
len1 = Len(txtStr1.Text)

For i = len1 To 1 Step -1
    s = Mid(txtStr1.Text, i, 1)
    res = res & s
Next
 
txtRes1.Text = res


End Sub
