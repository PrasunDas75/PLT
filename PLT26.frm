VERSION 5.00
Begin VB.Form frmArrayBinary 
   Caption         =   "Form2"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6690
   LinkTopic       =   "Form2"
   ScaleHeight     =   5760
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResSort 
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   3720
      Width           =   6135
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort"
      Height          =   615
      Left            =   2160
      TabIndex        =   11
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdEnterN 
      Caption         =   "Enter"
      Height          =   615
      Left            =   4080
      TabIndex        =   10
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter Elements"
      Height          =   615
      Left            =   2160
      TabIndex        =   9
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtEnter 
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSrch 
      Caption         =   "Search"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtRes2 
      Height          =   615
      Left            =   3840
      TabIndex        =   6
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox txtSrch 
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtRes1 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   6135
   End
   Begin VB.TextBox txtN 
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Search element:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Array elements:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Enter N:"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmArrayBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i, j As Integer

Dim n As Integer
Dim a() As Integer


Private Sub cmdEnterN_Click()
n = Val(txtN.Text) - 1
ReDim a(n)
End Sub

Private Sub cmdEnter_Click()
a(i) = Val(txtEnter.Text)
txtRes1.Text = txtRes1.Text & " " & a(i)
txtEnter.Text = ""
i = i + 1
End Sub



Private Sub cmdSort_Click()
Dim t As Integer

For i = 0 To n
    For j = i + 1 To n
        If a(i) > a(j) Then
            t = a(i)
            a(i) = a(j)
            a(j) = t
        End If
    Next
Next
For i = 0 To n
    txtResSort.Text = txtResSort.Text & " " & a(i)
Next
End Sub

Private Sub cmdSrch_Click()
Dim first, last, midl, srch As Integer

Dim s As String

srch = Val(txtSrch.Text)

first = 0
last = n

While first <= last
    midl = first + (last - first) \ 2
    If (a(midl) = srch) Then
        s = "found in " & Str(midl + 1)
    End If
    If a(midl) < srch Then
        first = midl + 1
    Else
        last = midl - 1
    End If
    txtRes2.Text = s
Wend
End Sub
