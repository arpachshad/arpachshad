VERSION 5.00
Begin VB.Form frmSMenu 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   3135
   ClientTop       =   4620
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSignOut 
      BackColor       =   &H00FFFF80&
      Caption         =   "Sign Out"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "ORDER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8640
      TabIndex        =   3
      Top             =   4440
      Width           =   3255
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00000080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   19920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdGrandTotal 
      Caption         =   "GRAND TOTAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   13560
      TabIndex        =   1
      Top             =   4560
      Width           =   3255
   End
   Begin VB.CommandButton cmdFood 
      Caption         =   "ALL ITEM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3960
      TabIndex        =   0
      Top             =   4440
      Width           =   3255
   End
End
Attribute VB_Name = "frmSMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDessert_Click()
frmSMenu.Hide
frmSListOfDessert.Show
End Sub

Private Sub cmdDrinks_Click()
frmSMenu.Hide
frmSListOfDrinks.Show
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdFood_Click()
frmSMenu.Hide
frmSListAllMenu.Show
End Sub

Private Sub cmdGrandTotal_Click()

'Insert all data item'
Open "C:\Users\habib\Desktop\Project NewVb\item.txt" For Input As #1

Do Until EOF(1)

Input #1, itemId, itemName, itemPrice, totalPrice

If (itemId = 1) Then
frmSTotalSalesFood.lblBurger = itemName & " - " & itemPrice


ElseIf (itemId = 2) Then
frmSTotalSalesFood.lblSpagetti = itemName & " - " & itemPrice


ElseIf (itemId = 3) Then
frmSTotalSalesFood.lblNasiLemak = itemName & " - " & itemPrice


ElseIf (itemId = 4) Then
frmSTotalSalesFood.lblChickenGrill = itemName & " - " & itemPrice


ElseIf (itemId = 5) Then
frmSTotalSalesFood.lblMee = itemName & " - " & itemPrice


ElseIf (itemId = 6) Then
frmSTotalSalesDrinks.lblDetox.Caption = itemName & " - " & itemPrice


ElseIf (itemId = 7) Then
frmSTotalSalesDrinks.lblOreo.Caption = itemName & " - " & itemPrice


ElseIf (itemId = 8) Then
frmSTotalSalesDrinks.lblSkyJuice.Caption = itemName & " - " & itemPrice


ElseIf (itemId = 9) Then
frmSTotalSalesDrinks.lblOrange.Caption = itemName & " - " & itemPrice


ElseIf (itemId = 10) Then
frmSTotalSalesDrinks.lblWatermelon.Caption = itemName & " - " & itemPrice


ElseIf (itemId = 11) Then
frmsTotalSalesDessert.lblPisang = itemName & " - " & itemPrice


ElseIf (itemId = 12) Then
frmsTotalSalesDessert.lblWaffle = itemName & " - " & itemPrice


ElseIf (itemId = 13) Then
frmsTotalSalesDessert.lblChocolate = itemName & " - " & itemPrice


ElseIf (itemId = 14) Then
frmsTotalSalesDessert.lblCake = itemName & " - " & itemPrice


ElseIf (itemId = 15) Then
frmsTotalSalesDessert.lblCornDog = itemName & " - " & itemPrice



End If

Loop

Close #1

frmSMenu.Hide
frmSGrandTotal.Show
End Sub

Private Sub cmdOrder_Click()

Me.Hide
frmSListOrder.Show

End Sub

Private Sub cmdSignOut_Click()

Me.Hide
frmWelcome.Show

End Sub
