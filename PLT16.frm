VERSION 5.00
Begin VB.Form frmMof7 
   Caption         =   "MultiplesOfSeven"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes5 
      Height          =   495
      Left            =   840
      TabIndex        =   9
      Top             =   7080
      Width           =   2895
   End
   Begin VB.CommandButton cmdStart5 
      Caption         =   "Div by 6"
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox txtRes4 
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   5640
      Width           =   2895
   End
   Begin VB.CommandButton cmdStart4 
      Caption         =   "Div by 5"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtRes3 
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton cmdStart3 
      Caption         =   "Div by 4"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtRes2 
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton cmdStart2 
      Caption         =   "Div by 3"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtRes 
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton cmdStart1 
      Caption         =   "Div by 2"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "1st,2nd ,and 4th multiple of 7 which gives remainder 1"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "frmMof7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n, i, j, c As Long


Private Sub cmdStart_Click()

'For i = 1 To 10000
'    n = 7 * i
'        If (n Mod 2 = 1) And (n Mod 3 = 1) And (n Mod 4 = 1) And (n Mod 5 = 1) And (n Mod 6 = 1) Then
'        c = c + 1
'        If c = 1 Or c = 2 Or c = 4 Then
'            txtRes.Text = txtRes.Text & " " & n
'        End If
'        Exit For
'        End If
'Next

End Sub

Private Sub cmdStart1_Click()
c = 0

For i = 1 To 10
    n = 7 * i
        If (n Mod 2 = 1) Then
        c = c + 1
        If c = 1 Or c = 2 Or c = 4 Then
            txtRes.Text = txtRes.Text & " " & n
        End If
        End If
Next
End Sub

Private Sub cmdStart2_Click()
c = 0

For i = 1 To 20
    n = 7 * i
        If (n Mod 3 = 1) Then
        c = c + 1
        If c = 1 Or c = 2 Or c = 4 Then
            txtRes2.Text = txtRes2.Text & " " & n
        End If
        End If
Next
End Sub

Private Sub cmdStart3_Click()
c = 0

For i = 1 To 20
    n = 7 * i
        If (n Mod 4 = 1) Then
        c = c + 1
        If c = 1 Or c = 2 Or c = 4 Then
            txtRes3.Text = txtRes3.Text & " " & n
        End If
        End If
Next
End Sub

Private Sub cmdStart4_Click()
c = 0

For i = 1 To 20
    n = 7 * i
        If (n Mod 5 = 1) Then
        c = c + 1
        If c = 1 Or c = 2 Or c = 4 Then
            txtRes4.Text = txtRes4.Text & " " & n
        End If
        End If
Next
End Sub

Private Sub cmdStart5_Click()
c = 0

For i = 1 To 20
    n = 7 * i
        If (n Mod 6 = 1) Then
        c = c + 1
        If c = 1 Or c = 2 Or c = 4 Then
            txtRes5.Text = txtRes5.Text & " " & n
        End If
        End If
Next
End Sub
