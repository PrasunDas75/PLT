VERSION 5.00
Begin VB.Form frmSimpleInterest 
   Caption         =   "SimpleInterest"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   3840
      Width           =   4935
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtRate 
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtTime 
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtPrinciple 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Simple Interest:"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Time:"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Rate:"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Principle:"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frmSimpleInterest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text4_Change()

End Sub

Private Sub txtCalculate_Click()

End Sub

Private Sub cmdResult_Click()


End Sub

Private Sub cmdCalculate_Click()
Dim P As Integer
Dim time As Integer
Dim rate As Integer
Dim si As Integer
    P = Val(txtPrinciple.Text)
    time = Val(txtTime.Text)
    rate = Val(txtRate.Text)
    si = (P * time * rate) / 100
    txtResult.Text = si
End Sub

