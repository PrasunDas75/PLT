VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes4 
      Height          =   1335
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdGen4 
      Caption         =   "Generate4"
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtRes3 
      Height          =   1335
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdGen3 
      Caption         =   "Generate3"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtRes2 
      Height          =   1335
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdGen2 
      Caption         =   "Generate2"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtRes1 
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdGen1 
      Caption         =   "Generate1"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtN 
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "N:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j, n As Integer
Dim s As String

Private Sub cmdGen1_Click()
n = Val(txtN.Text)

For i = 1 To n
    For j = 0 To 4
        s = s & "*"
    Next
    s = s & vbCrLf
Next
txtRes1.Text = s
End Sub

Private Sub cmdGen2_Click()
s = ""

For i = 1 To n
    For j = 0 To 4
        s = s & Str(i)
    Next
    s = s & vbCrLf
Next
txtRes2.Text = s
End Sub

Private Sub cmdGen3_Click()
s = ""

For i = 1 To n
    For j = 0 To 4
        s = s & Str(j + 1)
    Next
    s = s & vbCrLf
Next
txtRes3.Text = s
End Sub

Private Sub cmdGen4_Click()
s = ""

Dim k As Integer

For i = 1 To n
    For j = 1 To i
        s = s & "*" & " "
    Next
    s = s & vbCrLf
Next
txtRes4.Text = s
End Sub
