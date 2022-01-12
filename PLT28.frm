VERSION 5.00
Begin VB.Form frmIdentityMat 
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd1 
      Caption         =   "Add"
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtElements1 
      Height          =   735
      Left            =   5400
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtCollumns1 
      Height          =   735
      Left            =   3000
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
        End If
    Next
Next

End Sub
