VERSION 5.00
Begin VB.Form frmSGrandTotal 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   6
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H0080FFFF&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9720
      Width           =   1335
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
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdGrandTotal 
      Caption         =   "GRAND TOTAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7200
      TabIndex        =   3
      Top             =   6360
      Width           =   5775
   End
   Begin VB.CommandButton cmdGrandDessert 
      Caption         =   "DESSERT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   14640
      TabIndex        =   2
      Top             =   3360
      Width           =   3135
   End
   Begin VB.CommandButton cmdGrandDrinks 
      Caption         =   "DRINKS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8640
      TabIndex        =   1
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton cmdGrandFood 
      Caption         =   "FOOD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2640
      Picture         =   "frmSGrandTotal.frx":0000
      TabIndex        =   0
      Top             =   3480
      Width           =   3015
   End
End
Attribute VB_Name = "frmSGrandTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdGrandDessert_Click()

frmsTotalSalesDessert.lblDisplaySalesCake = gintCakeQuantity & " unit"
frmsTotalSalesDessert.lblDisplayProfitCake = FormatCurrency(gcurAllCakePrice, 2)
frmsTotalSalesDessert.lblDisplaySalesPisang = gintPisangQuantity & " unit"
frmsTotalSalesDessert.lblDisplayProfitPisang = FormatCurrency(gcurAllPisangPrice, 2)
frmsTotalSalesDessert.lblDisplaySalesChocolate = gintEclairsQuantity & " unit"
frmsTotalSalesDessert.lblDisplayProfitChocolate = FormatCurrency(gcurAllEclairsPrice, 2)
frmsTotalSalesDessert.lblDisplaySalesWaffle = gintWaffleQuantity & " unit"
frmsTotalSalesDessert.lblDisplayProfitWaffle = FormatCurrency(gcurAllWafflePrice, 2)
frmsTotalSalesDessert.lblDisplaySalesCornDog = gintCorndogQuantity & " unit"
frmsTotalSalesDessert.lblDisplayProfitCornDog = FormatCurrency(gcurAllCorndogPrice, 2)

frmsTotalSalesDessert.lblDisplayTotalSalesDessert.Caption = FormatCurrency(gcurAllDessertSale, 2)

frmSGrandTotal.Hide
frmsTotalSalesDessert.Show
End Sub

Private Sub cmdGrandDrinks_Click()

frmSTotalSalesDrinks.lblDisplaySalesStrawberry = gintStrawberryQuantity & " unit"
frmSTotalSalesDrinks.lblDisplayProfitStrawberry = FormatCurrency(gcurAllStrawberryPrice, 2)
frmSTotalSalesDrinks.lblDisplaySalesOreo = gintOreoQuantity & " unit"
frmSTotalSalesDrinks.lblDisplayProfitOreo = FormatCurrency(gcurAllOreoPrice, 2)
frmSTotalSalesDrinks.lblDisplaySalesWater = gintSkyQuantity & " unit"
frmSTotalSalesDrinks.lblDisplayProfitWater = FormatCurrency(gcurAllSkyPrice, 2)
frmSTotalSalesDrinks.lblDisplaySalesOrange = gintOrangeQuantity & " unit"
frmSTotalSalesDrinks.lblDisplayProfitOrange = FormatCurrency(gcurAllOrangePrice, 2)
frmSTotalSalesDrinks.lblDisplaySalesWatermelon = gintWatermelonQuantity & " unit"
frmSTotalSalesDrinks.lblDisplayProfitWatermelon = FormatCurrency(gcurAllWatermelonPrice, 2)

frmSTotalSalesDrinks.lblDisplayTotalSalesDrinks.Caption = FormatCurrency(gcurAllBeverageSale, 2)

frmSGrandTotal.Hide
frmSTotalSalesDrinks.Show
End Sub

Private Sub cmdGrandFood_Click()

frmSTotalSalesFood.lblDisplaySalesBurger = gintBurgerQuantity & " unit"
frmSTotalSalesFood.lblDisplayProfitBurger = FormatCurrency(gcurAllBurgerPrice, 2)
frmSTotalSalesFood.lblDisplaySalesSpagetti = gintNoodlesQuantity & " unit"
frmSTotalSalesFood.lblDisplayProfitSpagetti = FormatCurrency(gcurAllSpagettiPrice, 2)
frmSTotalSalesFood.lblDisplaySalesChickenGrill = gintChickenQuantity & " unit"
frmSTotalSalesFood.lblDisplayProfitChickenGrill = FormatCurrency(gcurAllChickenPrice, 2)
frmSTotalSalesFood.lblDisplaySalesMeeGoreng = gintNoodlesQuantity & " unit"
frmSTotalSalesFood.lblDisplayProfitMeeGoreng = FormatCurrency(gcurAllNoodlesPrice, 2)
frmSTotalSalesFood.lblDisplaySalesNasiLemak = gintNasiLemakQuantity & " unit"
frmSTotalSalesFood.lblDisplayProfitNasiLemak = FormatCurrency(gcurAllNasiLemakPrice, 2)

frmSTotalSalesFood.lblDisplayTotalSalesFood.Caption = FormatCurrency(gcurAllFoodSale, 2)

frmSGrandTotal.Hide
frmSTotalSalesFood.Show
End Sub

Private Sub cmdGrandTotal_Click()
MsgBox "THE GRAND TOTAL SALES OF THE DAY IS : " & FormatCurrency(gcurAllDessertSale + gcurAllBeverageSale + gcurAllFoodSale, 2), vbOKOnly, "GRAND TOTAL SALES"
End Sub

Private Sub cmdNext_Click()
Me.Hide
frmSMenu.Show
End Sub

Private Sub cmdSignOut_Click()

Me.Hide
frmWelcome.Show

End Sub
