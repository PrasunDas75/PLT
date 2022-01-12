VERSION 5.00
Begin VB.Form frmToString 
   Caption         =   "Form2"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8550
   LinkTopic       =   "Form2"
   ScaleHeight     =   5910
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes 
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   3240
      Width           =   6135
   End
   Begin VB.CommandButton cmdDisp 
      Caption         =   "Display"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtNum1 
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frmToString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDisp_Click()
Dim num As Double
Dim digit As String
Dim r As Double

num = Val(txtNum1.Text)

While num > 0
    r = num Mod 10
    Select Case r
        Case 1
        digit = "One"
        Case 2
        digit = "Two"
        Case 3
        digit = "Three"
        Case 4
        digit = "Four"
        Case 5
        digit = "Five"
        Case 6
        digit = "Six"
        Case 7
        digit = "Seven"
        Case 8
        digit = "Eight"
        Case 9
        digit = "Nine"
        Case 0
        digit = "Zero"
    End Select
    
    txtRes.Text = digit & " " & txtRes.Text
    num = num \ 10
    
    
Wend

End Sub
