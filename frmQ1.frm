VERSION 5.00
Begin VB.Form frmQ1 
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes 
      Height          =   1095
      Left            =   840
      TabIndex        =   5
      Top             =   3960
      Width           =   4575
   End
   Begin VB.CommandButton cmdDisp 
      Caption         =   "Display"
      Height          =   735
      Left            =   1680
      TabIndex        =   4
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtHigh 
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtLow 
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "HIgh"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Low"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmQ1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDisp_Click()
Dim n, m, sum As Integer
n = Val(txtLow.Text)
m = Val(txtHigh.Text)

sum = 0

If (n > m) Then
    MsgBox "low cannot be greter than high"
End If

For i = n To m
    If (i Mod 4 = 0 And i Mod 5 <> 0) Then
        sum = sum + i
    End If
Next

txtRes.Text = sum

End Sub
