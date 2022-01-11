VERSION 5.00
Begin VB.Form frmBinarytoDecimal 
   Caption         =   "Form2"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5415
   LinkTopic       =   "Form2"
   ScaleHeight     =   4845
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes1 
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CommandButton cmdConv 
      Caption         =   "Convert"
      Height          =   735
      Left            =   1800
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtNum1 
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Decimal"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Enter binary number"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmBinarytoDecimal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConv_Click()

Dim n, dec, base, re As Integer
n = Val(txtNum1.Text)
dec = 0
base = 1

While n > 0
    re = n Mod 10
    dec = dec + re * base
    n = n \ 10
    base = base * 2
Wend

txtRes1.Text = dec
End Sub
