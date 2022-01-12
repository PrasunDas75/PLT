VERSION 5.00
Begin VB.Form frmToThePower 
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes 
      Height          =   735
      Left            =   2160
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox txtExponent 
      Height          =   975
      Left            =   4320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtBase 
      Height          =   975
      Left            =   840
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Height          =   735
      Left            =   2400
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
End
Attribute VB_Name = "frmToThePower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalc_Click()
Dim Base, Expo As Integer
Dim v1 As Double
v1 = 1
Base = Val(txtBase.Text)
Expo = Val(txtExponent.Text)

While Expo <> 0
    v1 = v1 * Base
    Expo = Expo - 1
Wend
txtRes.Text = v1
End Sub
