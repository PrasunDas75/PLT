VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "PLT List"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command29 
      Caption         =   "PLT29"
      Height          =   615
      Left            =   5160
      TabIndex        =   28
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command28 
      Caption         =   "PLT28"
      Height          =   615
      Left            =   3960
      TabIndex        =   27
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command27 
      Caption         =   "PLT27"
      Height          =   615
      Left            =   2760
      TabIndex        =   26
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command26 
      Caption         =   "PLT26"
      Height          =   615
      Left            =   1560
      TabIndex        =   25
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command25 
      Caption         =   "PLT25"
      Height          =   615
      Left            =   360
      TabIndex        =   24
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command24 
      Caption         =   "PLT24"
      Height          =   615
      Left            =   8760
      TabIndex        =   23
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command23 
      Caption         =   "PLT23"
      Height          =   615
      Left            =   7560
      TabIndex        =   22
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command22 
      Caption         =   "PLT22"
      Height          =   615
      Left            =   6360
      TabIndex        =   21
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command21 
      Caption         =   "PLT21"
      Height          =   615
      Left            =   5160
      TabIndex        =   20
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command20 
      Caption         =   "PLT20"
      Height          =   615
      Left            =   3960
      TabIndex        =   19
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command19 
      Caption         =   "PLT19"
      Height          =   615
      Left            =   2760
      TabIndex        =   18
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command18 
      Caption         =   "PLT18"
      Height          =   615
      Left            =   1560
      TabIndex        =   17
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      Caption         =   "PLT17"
      Height          =   615
      Left            =   360
      TabIndex        =   16
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "PLT16"
      Height          =   615
      Left            =   8760
      TabIndex        =   15
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "PLT15"
      Height          =   615
      Left            =   7560
      TabIndex        =   14
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command14 
      Caption         =   "PLT14"
      Height          =   615
      Left            =   6360
      TabIndex        =   13
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      Caption         =   "PLT13"
      Height          =   615
      Left            =   5160
      TabIndex        =   12
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "PLT12"
      Height          =   615
      Left            =   3960
      TabIndex        =   11
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "PLT11"
      Height          =   615
      Left            =   2760
      TabIndex        =   10
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "PLT10"
      Height          =   615
      Left            =   1560
      TabIndex        =   9
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "PLT9"
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "PLT8"
      Height          =   615
      Left            =   8760
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "PLT7"
      Height          =   615
      Left            =   7560
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PLT6"
      Height          =   615
      Left            =   6360
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PLT5"
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PLT4"
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PLT3"
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PLT2"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PLT1"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
frmSimpleInterest.Show

End Sub

Private Sub Command10_Click()
frmToString.Show

End Sub

Private Sub Command11_Click()
frmPattern11.Show

End Sub

Private Sub Command12_Click()
frmPrime.Show

End Sub

Private Sub Command13_Click()
frmFactorial.Show

End Sub

Private Sub Command14_Click()
frmDecimaltoBinary.Show

End Sub

Private Sub Command15_Click()
frmBinarytoDecimal.Show

End Sub

Private Sub Command16_Click()
frmMof7.Show

End Sub

Private Sub Command17_Click()
frmPurchase.Show

End Sub

Private Sub Command18_Click()
frmPattern18.Show

End Sub

Private Sub Command19_Click()
frmToThePower.Show

End Sub

Private Sub Command2_Click()
frmSwapNumber.Show

End Sub

Private Sub Command20_Click()
frmRevString.Show

End Sub

Private Sub Command21_Click()
frmCheckPalindrome.Show

End Sub

Private Sub Command22_Click()
frmPattern22.Show

End Sub

Private Sub Command23_Click()
frmPattern23.Show

End Sub

Private Sub Command24_Click()
frmPattern24.Show

End Sub

Private Sub Command25_Click()
frmArrayLinear.Show

End Sub

Private Sub Command26_Click()
frmArrayBinary.Show

End Sub

Private Sub Command27_Click()
frmMatDisp.Show

End Sub

Private Sub Command28_Click()
frmIdentityMat.Show

End Sub

Private Sub Command29_Click()
frmSymMat.Show

End Sub

Private Sub Command3_Click()
frmEvenOdd.Show

End Sub

Private Sub Command4_Click()
frmDecimalSeparator.Show

End Sub

Private Sub Command5_Click()
frmStdentDB.Show

End Sub

Private Sub Command6_Click()
frmLargest.Show

End Sub

Private Sub Command7_Click()
frmEmp.Show

End Sub

Private Sub Command8_Click()
frmSumOdd.Show

End Sub

Private Sub Command9_Click()
frmReverse.Show

End Sub
