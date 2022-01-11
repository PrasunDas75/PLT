VERSION 5.00
Begin VB.Form frmCheckPalindrome 
   Caption         =   "Form2"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4905
   LinkTopic       =   "Form2"
   ScaleHeight     =   4890
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes1 
      Height          =   735
      Left            =   1680
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox txtStr1 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "is palindrome?"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the String"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmCheckPalindrome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheck_Click()

Dim len1 As Integer, i As Integer
Dim s As String
Dim res As String
 
len1 = Len(txtStr1.Text)

For i = len1 To 1 Step -1
    s = Mid(txtStr1.Text, i, 1)
    res = res & s
Next
 
If txtStr1.Text = res Then
    txtRes1.Text = "yes"
Else
    txtRes1.Text = "no"
End If


End Sub
