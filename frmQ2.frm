VERSION 5.00
Begin VB.Form frmQ2 
   Caption         =   "Form1"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Height          =   735
      Left            =   2760
      TabIndex        =   5
      Top             =   4920
      Width           =   3255
   End
   Begin VB.TextBox txtMinBal 
      Height          =   1095
      Left            =   2160
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   6360
      Width           =   4575
   End
   Begin VB.TextBox txtCat 
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox txtType 
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtYrs 
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtAcNo 
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "frmQ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalc_Click()
Dim acno, yrs, min, exemp As Integer
Dim actype, cat As String

acno = Val(txtAcNo.Text)
yrs = Val(txtYrs.Text)

actype = (txtType.Text)
cat = (txtCat.Text)

If (actype = "SAVINGS" And cat = "REGULAR") Then
    min = 5000
ElseIf (actype = "SAVINGS" And cat = "GOLD") Then
    min = 25000
ElseIf (actype = "SAVINGS" And cat = "PREMIUM") Then
    min = 100000
ElseIf (actype = "CURRENT" And cat = "REGULAR") Then
    min = 25000
ElseIf (actype = "CURRENT" And cat = "GOLD") Then
    min = 100000
ElseIf (actype = "CURRENT" And cat = "PREMIUM") Then
    min = 300000
End If


If (yrs > 5 And yrs <= 15) Then
    min = min - (min * 5 * (yrs - 5) \ 100)
ElseIf (yrs > 15) Then
    min = min - (min * 5 * 10 \ 100)
End If


txtMinBal.Text = min

End Sub
