VERSION 5.00
Begin VB.Form frmQ3 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNoOf 
      Caption         =   "Enter Number"
      Height          =   735
      Left            =   2880
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtN 
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton txtDisp 
      Caption         =   "Display"
      Height          =   855
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton txtEnt 
      Caption         =   "Enter Elements"
      Height          =   855
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtEle 
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "frmQ3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a() As Integer
Dim i As Integer
Dim n As Integer

Private Sub cmdNoOf_Click()
n = Val(txtN.Text)

ReDim a(n)
End Sub

Private Sub txtDisp_Click()
Dim c1, c2, c3 As Integer
c1 = 0
c2 = 0
c3 = 0

For j = 0 To n
    If a(j) \ 100 <> IsNull Then
        c3 = c3 + 1
    ElseIf a(j) \ 10 <> IsNull Then
        c2 = c2 + 1
    Else
        c1 = c1 + 1
    End If
Next

MsgBox "Count of 1: " & c1 & " Count of 2: " & c2 & " Count of 3: " & c3
End Sub

Private Sub txtEnt_Click()
i = 0

a(i) = Val(txtEle.Text)

i = i + 1


End Sub
