VERSION 5.00
Begin VB.Form frmPattern24 
   Caption         =   "GeneratePattern24"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   LinkTopic       =   "Form2"
   ScaleHeight     =   8685
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes4 
      Height          =   2535
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   5760
      Width           =   3015
   End
   Begin VB.CommandButton cmdGen4 
      Caption         =   "Generate4"
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdGen3 
      Caption         =   "Generate3"
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtRes3 
      Height          =   2535
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   5760
      Width           =   2775
   End
   Begin VB.CommandButton cmdGen2 
      Caption         =   "Generate2"
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtRes2 
      Height          =   2775
      Left            =   3840
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   5655
   End
   Begin VB.TextBox txtRes1 
      Height          =   2775
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   3015
   End
   Begin VB.CommandButton cmdGen1 
      Caption         =   "Generate1"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtN 
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "N:"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "frmPattern24"
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

Dim k As Integer
k = 1

For i = 1 To r
    For j = 1 To i
        If k Mod 2 = 0 Then
            s = s & Str(-(k ^ 2)) & " "
        Else
            s = s & Str(k ^ 2) & " "
        End If
        k = k + 1
    Next
    s = s & vbCrLf
Next
txtRes1.Text = s
End Sub

Private Sub cmdGen2_Click()
s = ""

Dim k, f, c As Integer
k = 0
c = 1
f = 1

For i = 1 To r
    For j = 1 To i
        If k = 0 Then
            s = s & Str(1) & " "
        Else
            While c <= k
                f = f * c
                c = c + 1
            Wend
            s = s & Str(f) & " "
        End If
        k = k + 1
    Next
    s = s & vbCrLf
Next
txtRes2.Text = s
End Sub

Private Sub cmdGen3_Click()
s = ""

Dim k As Integer

For i = 1 To r
    For j = i To r
        s = s & " "
    Next
    For k = 1 To i
        s = s & "*"
    Next
    s = s & vbCrLf
Next
txtRes3.Text = s
End Sub

Private Sub cmdGen4_Click()
s = ""

Dim k, sp, num As Integer

sp = r - 1
num = 1

For i = 1 To r
    For j = 1 To sp
        s = s & " "
    Next
    For k = 1 To num
        s = s & "*"
    Next
    If sp > i Then
        sp = sp - 1
        num = num + 2
    End If
    If sp < i Then
        sp = sp + 1
        num = num - 2
    End If
    s = s & vbCrLf
Next
txtRes4.Text = s
End Sub
