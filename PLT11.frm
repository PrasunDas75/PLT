VERSION 5.00
Begin VB.Form frmPattern11 
   Caption         =   "GenerateSeries11"
   ClientHeight    =   9990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   9990
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes6 
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   9120
      Width           =   6735
   End
   Begin VB.CommandButton cmdStart6 
      Caption         =   "Start6"
      Height          =   495
      Left            =   2160
      TabIndex        =   12
      Top             =   8520
      Width           =   1815
   End
   Begin VB.TextBox txtRes5 
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   6480
      Width           =   6735
   End
   Begin VB.CommandButton cmdStart5 
      Caption         =   "Start4"
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      Top             =   5760
      Width           =   1815
   End
   Begin VB.TextBox txtRes4 
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   7800
      Width           =   6735
   End
   Begin VB.CommandButton cmdStart4 
      Caption         =   "Start5"
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   7200
      Width           =   1815
   End
   Begin VB.TextBox txtRes3 
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   5040
      Width           =   6735
   End
   Begin VB.CommandButton cmdStart3 
      Caption         =   "Start3"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtRes2 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   6735
   End
   Begin VB.CommandButton cmdStart2 
      Caption         =   "Start2"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtRes1 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   6735
   End
   Begin VB.CommandButton cmdStart1 
      Caption         =   "Start1"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtN 
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   855
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
Dim i, n As Integer
Dim s As String

Private Sub cmdStart1_Click()

n = Val(txtN.Text)

For i = 1 To n
    If i Mod 2 = 0 Then
    If i * i <= n Then
        s = s & Str(i * i)
    End If
    End If
Next

txtRes1.Text = s

End Sub

Private Sub cmdStart2_Click()
s = ""

For i = 1 To n
    If i <= n Then
    If i Mod 2 = 0 Then
        s = s & " " & Str(i * -1)
        
    Else
        s = s & " " & Str(i)
    End If
    End If
Next

txtRes2.Text = s
End Sub

Private Sub cmdStart3_Click()
s = ""

For i = 1 To n
    If i ^ i <= n Then
        s = s & " " & Str(i ^ i)
    End If
Next

txtRes3.Text = s
End Sub

Private Sub cmdStart4_Click()
s = ""

Dim count As Integer
count = 0

For i = 1 To n
    count = count + 1
    If count = 4 Then
        s = s
        count = 0
    Else
        If (i ^ 2) <= n Then
        s = s & " " & Str(i ^ 2)
        End If
    End If
Next

txtRes4.Text = s
End Sub

Private Sub cmdStart5_Click()
s = " "

Dim count As Integer
Dim b As Integer
count = 0
b = 1

For i = 1 To n
    count = count + 1
    
    If count = 3 Then
        b = b
        s = s
        count = 0
    Else
        b = b + 4 * i
        
        
    End If
    
Next

txtRes5.Text = s
End Sub

Private Sub cmdStart6_Click()
s = "1"

Dim count As Integer
Dim b As Integer
count = 0
b = 1

For i = 1 To n
    count = count + 1
    If count = 3 Then
        b = b
        s = s
        count = 0
    Else
        b = b + 4 * i
        If b <= n Then
        s = s & " " & Str(b)
        End If
    End If
Next

txtRes6.Text = s
End Sub
