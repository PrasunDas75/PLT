VERSION 5.00
Begin VB.Form frmIdentityMat 
   Caption         =   "CheckIfIdentityMatrix"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd1 
      Caption         =   "Add"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtElements1 
      Height          =   735
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtCollumns1 
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtRows1 
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Element"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Column"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Row"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmIdentityMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(10, 10) As Integer
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim y As Integer

Private Sub cmdAdd1_Click()

Dim v1 As Integer

i = Val(txtCollumns1.Text)
j = Val(txtRows1.Text)

v1 = Val(txtElements1.Text)

a(i, j) = v1

End Sub

Private Sub cmdCheck_Click()

For x = 0 To i
    For y = 0 To j
        If a(x, x) = 1 And a(y, y) = 1 Then
            MsgBox "Identity Matrix"
            Exit For
        Else
            MsgBox "Not Identity Matrix"
            Exit For
        End If
    Next
Next

End Sub
