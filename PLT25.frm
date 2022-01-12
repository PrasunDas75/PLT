VERSION 5.00
Begin VB.Form frmArrayLinear 
   Caption         =   "Form2"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6675
   LinkTopic       =   "Form2"
   ScaleHeight     =   4590
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
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
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtRes2 
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtSrch 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3480
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
      Top             =   3000
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
Attribute VB_Name = "frmArrayLinear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim i As Integer
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



Private Sub cmdSrch_Click()
Dim srch As Integer
Dim c As Integer
srch = Val(txtSrch.Text)

For c = 0 To n
    If a(c) = srch Then
        txtRes2.Text = "Presnt at " & (c + 1)
        Exit For
    End If
Next
If c = n + 1 Then
    txtRes2.Text = "Not Present"
End If
End Sub
