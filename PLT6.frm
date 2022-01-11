VERSION 5.00
Begin VB.Form frmStdentDB 
   Caption         =   "Form2"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10365
   LinkTopic       =   "Form2"
   ScaleHeight     =   7980
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRight 
      Caption         =   ">>"
      Height          =   375
      Left            =   8760
      TabIndex        =   17
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<<"
      Height          =   375
      Left            =   5760
      TabIndex        =   16
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtTotal 
      Height          =   615
      Left            =   7320
      TabIndex        =   11
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtAvg 
      Height          =   615
      Left            =   7320
      TabIndex        =   10
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Student Database"
      Height          =   3135
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4455
      Begin VB.TextBox txtSub3 
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtSub2 
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtSub1 
         Height          =   405
         Left            =   1920
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Subject3"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Subject2"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Subject1"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lbName 
         Caption         =   "Name"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label lblRes 
      Caption         =   "Result"
      Height          =   255
      Left            =   5400
      TabIndex        =   14
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Total"
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Average"
      Height          =   495
      Left            =   5280
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmStdentDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type student
    Name As String
    sub1 As Integer
    sub2 As Integer
    sub3 As Integer
    total As Integer
    avg As Double
    result As String
End Type

Dim S(20) As student
Dim index As Integer
Dim ci As Integer

Private Sub cmdClear_Click()
txtName.Text = ""
txtSub1.Text = ""
txtSub2.Text = ""
txtSub3.Text = ""
txtAvg.Text = ""
txtTotal.Text = ""
lblRes.Caption = ""
End Sub

Private Sub cmdLeft_Click()
If ci > 0 Then
ci = ci - 1
getrecord ci
End If
End Sub

Private Sub cmdRight_Click()
If ci > 0 Then
ci = ci + 1
getrecord ci
End If
End Sub

Private Sub cmdSave_Click()
index = index + 1
ci = ci + 1
update (index)
End Sub

Private Sub update(index As Integer)
S(index).Name = txtName.Text

With S(index)
.sub1 = txtSub1.Text
.sub2 = txtSub2.Text
.sub3 = txtSub3.Text
.total = .sub1 + .sub2 + .sub3
.avg = .total / 3
txtAvg.Text = .avg
txtTotal.Text = .total

If (.avg) > 60 Then
lblRes.Caption = "First Class"
ElseIf (.avg) > 50 Then
lblRes.Caption = "Second Class"
ElseIf (.avg) > 35 Then
lblRes.Caption = "Pass"
Else
lblRes.Caption = "Fail"
End If
End With
End Sub

Private Sub getrecord(index As Integer)
With S(index)
txtName.Text = .Name
txtSub1.Text = .sub1
txtSub2.Text = .sub2
txtSub3.Text = .sub3
txtAvg.Text = .avg
txtTotal.Text = .total
End With
End Sub


