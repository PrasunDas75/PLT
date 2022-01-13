VERSION 5.00
Begin VB.Form frmEvenOdd 
   Caption         =   "OddOrEven"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6810
   LinkTopic       =   "Form2"
   ScaleHeight     =   3750
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txtNumber 
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Check if given No. is Odd or Even"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label lbEnter 
      Caption         =   "Enter The Number:"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
End
Attribute VB_Name = "frmEvenOdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCheck_Click()
Dim n As Integer
n = Val(txtNumber.Text)
If n Mod 2 = 0 Then
    MsgBox "even"
Else
    MsgBox "odd"
End If
End Sub
