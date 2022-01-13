VERSION 5.00
Begin VB.Form frmMatDisp 
   Caption         =   "DisplayMatrixAndTranspose"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes2 
      Height          =   1815
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   6120
      Width           =   3015
   End
   Begin VB.CommandButton cmdTranspose 
      Caption         =   "Transpose"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox txtRes1 
      Height          =   1815
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3000
      Width           =   3015
   End
   Begin VB.CommandButton cmdDisp 
      Caption         =   "Display"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtElem 
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtCollumns1 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtRows1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Elements"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Collumn"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Row"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmMatDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(10, 10) As Integer
Dim i, j As Integer
Dim x, y As Integer
Dim s As String

Private Sub cmdAdd_Click()

Dim v1 As Integer

i = Val(txtRows1.Text)
j = Val(txtCollumns1.Text)

v1 = Val(txtElem.Text)

a(i, j) = v1

End Sub

Private Sub cmdDisp_Click()

For x = 0 To i
    For y = 0 To j
        s = s & " " & a(x, y)
    Next
    txtRes1.Text = txtRes1.Text & s & vbCrLf
    s = ""
Next

End Sub

Private Sub cmdTranspose_Click()

Dim t(10, 10) As Integer

For x = 0 To i
    For y = 0 To j
        t(y, x) = a(x, y)
    Next
Next

For x = 0 To i
    For y = 0 To j
        s = s & " " & t(x, y)
    Next
    txtRes2.Text = txtRes2.Text & s & vbCrLf
    s = ""
Next

End Sub
