VERSION 5.00
Begin VB.Form frmPurchase 
   Caption         =   "BuyingItems"
   ClientHeight    =   10050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   10050
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCash 
      Caption         =   "Cash"
      Height          =   615
      Left            =   3840
      TabIndex        =   18
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton cmdCard 
      Caption         =   "Card"
      Height          =   615
      Left            =   1080
      TabIndex        =   17
      Top             =   8160
      Width           =   1575
   End
   Begin VB.TextBox txtGrandTotal 
      Height          =   615
      Left            =   1200
      TabIndex        =   16
      Top             =   6720
      Width           =   4215
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Calculate Total"
      Height          =   735
      Left            =   2040
      TabIndex        =   15
      Top             =   5760
      Width           =   2655
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   495
      Left            =   3600
      TabIndex        =   13
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtIsNext 
      Height          =   735
      Left            =   2400
      TabIndex        =   12
      Top             =   4200
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Items"
      Height          =   3495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   6255
      Begin VB.CommandButton cmdTotPrice 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtTotPrice 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtPrice 
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtQty 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtDesc 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtCode 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Total Price"
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Price"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Qty."
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label Label6 
      Caption         =   "OK?"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   4440
      Width           =   735
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type item
    code As String
    desc As String
    qty As Double
    price As Double
    total As Double
    
End Type

Dim i(20) As item
Dim index As Integer
Dim ci As Integer
Dim gTotal As Double

Private Sub cmdCard_Click()
If gTotal > 10000 Then
    MsgBox "you paid " & Str(gTotal - (gTotal * 10 \ 100))
ElseIf gTotal < 1000 Then
    MsgBox "you paid " & Str(gTotal + (gTotal * 2.5 \ 100))
Else
    MsgBox "you paid " & Str(gTotal)
End If
End Sub

Private Sub cmdCash_Click()
If gTotal > 10000 Then
    MsgBox "you paid " & Str(gTotal - (gTotal * 10 \ 100))
Else
    MsgBox "you paid " & Str(gTotal)
End If
End Sub

Private Sub cmdNext_Click()
index = index + 1
If txtIsNext.Text = "y" Then
update (index)
txtCode.Text = ""
txtDesc.Text = ""
txtQty.Text = ""
txtPrice.Text = ""
txtTotPrice.Text = ""
End If
End Sub

Private Sub update(index As Integer)
With i(index)
.code = txtCode.Text
.desc = txtDesc.Text
.qty = Val(txtQty.Text)
.price = Val(txtPrice.Text)
.total = Val(txtTotPrice.Text)
End With
End Sub

Private Sub cmdTotal_Click()
For c = 0 To index
    gTotal = gTotal + i(c).total
Next
txtGrandTotal.Text = gTotal
End Sub

Private Sub cmdTotPrice_Click()
txtTotPrice.Text = Val(txtQty.Text) * Val(txtPrice.Text)
End Sub
