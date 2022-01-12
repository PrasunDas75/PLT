VERSION 5.00
Begin VB.Form frmEmp 
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Controls"
      Height          =   3615
      Left            =   5160
      TabIndex        =   20
      Top             =   3960
      Width           =   5295
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   615
         Left            =   1800
         TabIndex        =   24
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   735
         Left            =   1800
         TabIndex        =   23
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdRight 
         Caption         =   ">>"
         Height          =   735
         Left            =   3960
         TabIndex        =   22
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdLeft 
         Caption         =   "<<"
         Height          =   735
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblIndex 
         Height          =   495
         Left            =   480
         TabIndex        =   25
         Top             =   2880
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Salary"
      Height          =   3015
      Left            =   5160
      TabIndex        =   13
      Top             =   600
      Width           =   5295
      Begin VB.TextBox txtNet 
         Height          =   495
         Left            =   1680
         TabIndex        =   19
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txtAnnual 
         Height          =   495
         Left            =   1680
         TabIndex        =   18
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtGross 
         Height          =   495
         Left            =   1680
         TabIndex        =   17
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Annual Net"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Annual"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Gross"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employee"
      Height          =   7095
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4455
      Begin VB.TextBox txtTaxInv 
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   4200
         Width           =   2295
      End
      Begin VB.TextBox txtBonus 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox txtAllo 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtBSal 
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txtID 
         Height          =   405
         Left            =   1800
         TabIndex        =   8
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Tax Saving"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "% of Bonus"
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Special Allowances"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Basic Salary"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Emp ID"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type emp
    Name As String
    Id As String
    BasicS As Double
    SpecialA As Double
    Bonus As Double
    TaxSI As Double
    Gross As Double
    Annual As Double
    AnnualNet As Double
    ATax As Double
End Type

Dim E(20) As emp
Dim index As Integer
Dim ci As Integer

Private Sub cmdClear_Click()
cmdSave.Enabled = True
txtName.Text = ""
txtID.Text = ""
txtBSal.Text = ""
txtAllo.Text = ""
txtBonus.Text = ""
txtTaxInv.Text = ""
End Sub

Private Sub cmdLeft_Click()
cmdSave.Enabled = False
If ci > 0 Then
ci = ci - 1
getrecord ci
End If
End Sub

Private Sub cmdRight_Click()
cmdSave.Enabled = False
If ci > 0 Then
ci = ci + 1
getrecord ci
End If
End Sub

Private Sub cmdSave_Click()
cmdLeft.Enabled = True
cmdRight.Enabled = True
index = index + 1
ci = ci + 1
update (index)
End Sub

Private Sub update(index As Integer)
With E(index)
.Name = txtName.Text
.Id = txtID.Text
.BasicS = Val(txtBSal.Text)
.SpecialA = Val(txtAllo.Text)
.Bonus = Val(txtBonus.Text) * .BasicS \ 100
.TaxSI = Val(txtTaxInv.Text)
.Gross = .BasicS + .SpecialA
.Annual = .BasicS + .SpecialA + .Bonus

If (.TaxSI <= 100000) Then
    .ATax = .Annual - .TaxSI
ElseIf (.TaxSI > 100000) Then
    .ATax = .Annual - 100000
End If

If (.ATax <= 100000) Then
    .AnnualNet = .ATax
ElseIf (.ATax > 100000 And .ATax <= 150000) Then
    .AnnualNet = .ATax - (.ATax * 20 \ 100)
ElseIf (.ATax > 150000) Then
    .AnnualNet = .ATax - (.ATax * 30 \ 100)
End If

txtGross.Text = .Gross
txtAnnual.Text = .Annual
txtNet.Text = .AnnualNet
lblIndex.Caption = "The Index is: " & index

End With
End Sub

Private Sub getrecord(index As Integer)
With E(index)
txtName.Text = .Name
txtID.Text = .Id
txtBSal.Text = .BasicS
txtAllo.Text = .SpecialA
txtBonus.Text = .Bonus * 100 \ .BasicS
txtTaxInv.Text = .TaxSI
txtGross.Text = .Gross
txtAnnual.Text = .Annual
txtNet.Text = .AnnualNet
lblIndex.Caption = "The Index is: " & index
End With
End Sub

Private Sub Form_Load()
cmdSave.Enabled = False
cmdClear.Enabled = False
cmdLeft.Enabled = False
cmdRight.Enabled = False
End Sub

Private Sub txtTaxInv_Change()
cmdSave.Enabled = True
cmdClear.Enabled = True
End Sub
