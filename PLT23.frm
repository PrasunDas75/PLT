VERSION 5.00
Begin VB.Form frmPattern23 
   Caption         =   "GeneratePattern23"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10650
   LinkTopic       =   "Form2"
   ScaleHeight     =   5670
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes4 
      Height          =   2655
      Left            =   7920
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdGen4 
      Caption         =   "Generate4"
      Height          =   495
      Left            =   8280
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdGen3 
      Caption         =   "Generate3"
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtRes3 
      Height          =   2655
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdGen2 
      Caption         =   "Generate2"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtRes2 
      Height          =   2655
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtRes1 
      Height          =   2655
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdGen1 
      Caption         =   "Generate1"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtN 
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "N:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmPattern23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim j As Integer
Dim r As Integer
Dim s As String

Private Sub cmdGen1_Click()
r = Val(txtN.Text)

For i = 1 To r
    For j = 1 To i
        s = s & Str(j) & " "
    Next
    s = s & vbCrLf
Next
txtRes1.Text = s
End Sub

Private Sub cmdGen2_Click()
s = ""

For i = 1 To r
    For j = 1 To i
        s = s & Str(i) & " "
    Next
    s = s & vbCrLf
Next
txtRes2.Text = s
End Sub

Private Sub cmdGen3_Click()
s = ""

Dim k As Integer
k = 1

For i = 1 To r
    For j = 1 To i
        s = s & Str(k) & " "
        k = k + 1
    Next
    s = s & vbCrLf
Next
txtRes3.Text = s
End Sub

Private Sub cmdGen4_Click()
s = ""

Dim N1 As Integer
Dim N2 As Integer
Dim N3 As Integer

N1 = 0
N2 = 1

txtRes4.Text = Str(N2) & vbCrLf


For i = 2 To r
    For j = 1 To i
        N3 = N1 + N2
            s = s & Str(N3)
            N1 = N2
            N2 = N3
    Next
    txtRes4.Text = txtRes4.Text & vbCrLf & s & vbCrLf
    s = ""
Next

End Sub
