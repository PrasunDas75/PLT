VERSION 5.00
Begin VB.Form frmDecimaltoBinary 
   Caption         =   "Form2"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5385
   LinkTopic       =   "Form2"
   ScaleHeight     =   5535
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes1 
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtNum1 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Binary"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter decimal number"
      Height          =   375
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

j = i - 1
While j >= 0
    
    txtRes1 = a(j) & " " & txtRes1
    
    j = j - 1
    
Wend
End Sub
