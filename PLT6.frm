VERSION 5.00
Begin VB.Form frmLargest 
   Caption         =   "FindLrgestAndSecLargest"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8955
   LinkTopic       =   "Form2"
   ScaleHeight     =   5490
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox txtSecondLargest 
      Height          =   615
      Left            =   5160
      TabIndex        =   4
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox txtLargest 
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox txtNum3 
      Height          =   615
      Left            =   6000
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtNum2 
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtNum1 
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "2nd Largest:"
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Largest:"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter 3 Numbers:"
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
      Left            =   720
      TabIndex        =   6
      Top             =   600
      Width           =   7215
   End
End
Attribute VB_Name = "frmLargest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFind_Click()
Dim num1 As Integer
Dim num2 As Integer
Dim num3 As Integer

Dim largest As Integer
Dim secondLargest As Integer

num1 = Val(txtNum1.Text)
num2 = Val(txtNum2.Text)
num3 = Val(txtNum3.Text)

If (num1 > num2 And num1 > num3) Then
        largest = num1
        
        If (num2 > num3) Then
            secondLargest = num2
        Else
            secondLargest = num3
        End If
        
        End If


If (num2 > num1 And num2 > num3) Then
        largest = num2
        
        If (num1 > num3) Then
            secondLargest = num1
        Else
            secondLargest = num3
        End If
        
        End If

If (num3 > num1 And num3 > num2) Then
        largest = num3
        
        If (num1 > num2) Then
            secondLargest = num1
        Else
            secondLargest = num2
        End If
        
        End If

txtLargest.Text = largest
txtSecondLargest.Text = secondLargest
End Sub


