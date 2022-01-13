VERSION 5.00
Begin VB.Form frmSymMat 
   Caption         =   "CheckIfSymmetricMatrix"
   ClientHeight    =   2715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtElems1 
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtCollumns1 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtRows1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Element"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Column"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Row"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmSymMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(10, 10) As Integer
Dim i, j As Integer
Dim x, y As Integer
Dim m, n As Integer

Private Sub cmdAdd_Click()
Dim v1 As Integer

i = Val(txtRows1.Text)
j = Val(txtCollumns1.Text)

v1 = Val(txtElems1.Text)

a(i, j) = v1

m = i

n = j

End Sub

Private Sub cmdCheck_Click()

Dim t(10, 10) As Integer

For x = 0 To m
    For y = 0 To n
        t(y, x) = a(x, y)
    Next
Next

Dim f As Integer
f = 1

For x = 0 To m
    For y = 0 To n
        If a(x, y) <> t(x, y) Then
            f = 0
            Exit For
        End If
    Next
    If f = 0 Then
        MsgBox "Not Symmetric"
        Exit For
    Else
        MsgBox "Symmetric"
        Exit For
    End If
Next


End Sub
