VERSION 5.00
Begin VB.Form frmMof7 
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes 
      Height          =   1335
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   6495
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmMof7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
Dim n, i, j, c As Integer
c = 0
For i = 1 To 10
    n = 7 * i
    For j = 2 To 7
        If (n Mod j = 1) Then
        c = c + 1
        If c = 1 Or c = 2 Or c = 4 Then
            txtRes.Text = txtRes.Text & " " & n
        End If
        Exit For
        End If
    Next
Next

End Sub
