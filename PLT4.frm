VERSION 5.00
Begin VB.Form frmLargest 
   Caption         =   "Form2"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11550
   LinkTopic       =   "Form2"
   ScaleHeight     =   7275
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   615
      Left            =   4560
      TabIndex        =   5
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox txtSecondLargest 
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Text            =   "2nd Largest"
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox txtLargest 
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Text            =   "Largest"
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox txtNum3 
      Height          =   615
      Left            =   7560
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtNum2 
      Height          =   615
      Left            =   4680
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtNum1 
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
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
