VERSION 5.00
Begin VB.Form frmPrime 
   Caption         =   "Form2"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6975
   LinkTopic       =   "Form2"
   ScaleHeight     =   4875
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSum 
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtRes 
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   6495
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtN 
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtM 
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Sum:"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Prime Numbers:"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "N:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "M:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "frmPrime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFind_Click()

Dim m, n, i, f, j, sum As Integer
m = Val(txtM.Text)
n = Val(txtN.Text)
sum = 0
f = 1

For i = m To n
    
    For j = 2 To i \ 2
    
        If i Mod j = 0 Then
            f = 0
            Exit For
        Else
            f = 1
        End If
    Next
    
    If f = 1 Then
       txtRes.Text = txtRes.Text & ", " & i
       sum = sum + i
    End If
    
Next

txtSum.Text = sum

End Sub

