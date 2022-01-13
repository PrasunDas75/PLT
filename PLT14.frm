VERSION 5.00
Begin VB.Form frmDecimaltoBinary 
   Caption         =   "ConvertDecimalToBinary"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5115
   LinkTopic       =   "Form2"
   ScaleHeight     =   3360
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes1 
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtNum1 
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Binary"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter decimal number"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmDecimaltoBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConvert_Click()

Dim a(10), n, i, j As Integer
n = Val(txtNum1.Text)

i = 0
While n > 0
    a(i) = n Mod 2
    n = n \ 2
    i = i + 1
    
Wend

j = 0
While j <= i
    
    txtRes1 = a(j) & " " & txtRes1
    
    j = j + 1
    
Wend
End Sub
