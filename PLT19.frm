VERSION 5.00
Begin VB.Form frmToThePower 
   Caption         =   "CalculateToThePower"
   ClientHeight    =   3690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRes 
      Height          =   615
      Left            =   1440
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtExponent 
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtBase 
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Exponent:"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Base:"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   600
      Width           =   735
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
