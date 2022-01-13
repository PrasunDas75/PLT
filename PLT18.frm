VERSION 5.00
Begin VB.Form frmPattern18 
   Caption         =   "GenerateSeries18"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7170
   LinkTopic       =   "Form2"
   ScaleHeight     =   6480
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes4 
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   5640
      Width           =   6615
   End
   Begin VB.CommandButton cmdGen4 
      Caption         =   "Generate4"
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtRes3 
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   6615
   End
   Begin VB.CommandButton cmdGen3 
      Caption         =   "Generate3"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdGen2 
      Caption         =   "Generate1"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtRes2 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   6615
   End
   Begin VB.TextBox txtRes1 
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3240
      Width           =   6615
   End
   Begin VB.CommandButton cmdGen1 
      Caption         =   "Generate2"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtN 
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "N:"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "frmPattern18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim j As Integer
Dim r As Integer
Dim s As String

Private Sub cmdGen1_Click()
s = ""

Dim n1 As Double
Dim n2 As Double
Dim n3 As Double

n1 = 0
n2 = 1

txtRes1.Text = Str(n2)

For i = 2 To r
    n3 = n1 + n2
    If n3 <= r Then
     s = s & Str(n3)
    End If
    n1 = n2
    n2 = n3
Next

txtRes1.Text = txtRes1.Text & s

End Sub

Private Sub cmdGen2_Click()

r = Val(txtN.Text)

s = ""

Dim n As Double
Dim count As Integer
n = 1

For i = 0 To r
    count = count + 1
    n = n + (i ^ 2)
    If n < r Then
       If count Mod 2 = 0 Then
            s = s & " " & Str(-n)
       Else
            s = s & " " & Str(n)
       End If
    End If
Next

txtRes2.Text = s
End Sub

Private Sub cmdGen3_Click()
s = ""

Dim n1, n2 As Integer
Dim count As Integer

n1 = 1
n2 = 2

txtRes3.Text = Str(n1) & " " & Str(-n2)

For i = 1 To r
    If n1 < r - 2 Then
        n1 = n1 + 3
        s = s & " " & Str(n1)
        n2 = n2 + 4
        s = s & " " & Str(-n2)
    End If
Next

txtRes3.Text = txtRes3.Text & s
End Sub

Private Sub cmdGen4_Click()
s = ""

Dim n1, n2, n3, n4 As Integer
Dim count As Integer

n1 = 1
n2 = 5
n3 = 8

txtRes4.Text = Str(n1) & " " & Str(n2) & " " & Str(n3)

For i = 1 To r
        
        n4 = n1 + n2 + n3
        If n4 <= r Then
        s = s & " " & Str(n4)
        
        n1 = n2
        n2 = n3
        n3 = n4
        End If
    
Next

txtRes4.Text = txtRes4.Text & s
End Sub
