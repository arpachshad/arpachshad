VERSION 5.00
Begin VB.Form frmGMenu 
   Caption         =   "MENU"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14760
   LinkTopic       =   "Form2"
   Picture         =   "FormGMenu.frx":0000
   ScaleHeight     =   9150
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H0000FF00&
      Caption         =   "Reset Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8520
      Width           =   2775
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   20040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdPlaceOrder 
      BackColor       =   &H000080FF&
      Caption         =   "Place Your Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8520
      Width           =   2775
   End
   Begin VB.Label lblTotalPayment 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblTCart 
      BackColor       =   &H80000005&
      Caption         =   "Total Payment :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblDesert 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dessert"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   16440
      TabIndex        =   2
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Label lblFood 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Food"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   11280
      TabIndex        =   1
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label lblDrink 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Drink"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   5400
      TabIndex        =   0
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Image imgDessert 
      Height          =   2745
      Left            =   15000
      Picture         =   "FormGMenu.frx":3DE87
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   4440
   End
   Begin VB.Image imgFood 
      Height          =   2745
      Left            =   9240
      Picture         =   "FormGMenu.frx":60FC5
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   4560
   End
   Begin VB.Image imgDrink 
      Height          =   2745
      Left            =   3600
      Picture         =   "FormGMenu.frx":86DE4
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   4200
   End
End
Attribute VB_Name = "frmGMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdPlaceOrder_Click()

gcurAllFoodSale = gcurAllFoodSale + gcurBurgerPrice + gcurNoodlesPrice + gcurNasiLemakPrice + gcurChickenPrice + gcurSpagettiPrice

gcurAllBeverageSale = gcurAllBeverageSale + gcurOreoPrice + gcurWatermelonPrice + gcurSkyPrice + gcurOrangePrice + gcurStrawberryPrice

gcurAllDessertSale = gcurAllDessertSale + gcurPisangPrice + gcurWafflePrice + gcurEclairsPrice + gcurCakePrice + gcurCorndogPrice



'Print food to receipt'
If frmGFood.chkBurger.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGFood.lblBurger & Space(38) & frmGFood.lblPriceBurger & Space(20) & frmGFood.txtBurger & Space(20) & FormatCurrency(gcurBurgerPrice, 2))
End If
If frmGFood.chkNoodle.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGFood.lblMee & Space(39) & frmGFood.lblPriceMeeGoreng & Space(20) & frmGFood.txtNoodle & Space(20) & FormatCurrency(gcurNoodlesPrice, 2))
End If
If frmGFood.chkNasiLemak.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGFood.lblNasiLemak & Space(39) & frmGFood.lblPriceNasiLemak & Space(20) & frmGFood.txtNasiLemak & Space(20) & FormatCurrency(gcurNasiLemakPrice, 2))
End If
If frmGFood.chkChicken.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGFood.lblChickenGrill & Space(30) & frmGFood.lblPriceChickenGrill & Space(20) & frmGFood.txtChicken & Space(20) & FormatCurrency(gcurChickenPrice, 2))
End If
If frmGFood.chkSpagetti.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGFood.lblSpagetti & Space(15) & frmGFood.lblPriceSpagetti & Space(20) & frmGFood.txtSpagetti & Space(20) & FormatCurrency(gcurSpagettiPrice, 2))
End If

'Print beverage to receipt'
If frmGDrink.chkOreo.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGDrink.lblOreo & Space(48) & frmGDrink.lblPriceOreo & Space(20) & frmGDrink.txtOreo & Space(20) & FormatCurrency(gcurOreoPrice, 2))
End If
If frmGDrink.chkWatermelon.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGDrink.lblWatermelon & Space(45) & frmGDrink.lblPriceWaterMelon & Space(20) & frmGDrink.txtWatermelon & Space(20) & FormatCurrency(gcurWatermelonPrice, 2))
End If
If frmGDrink.chkSky.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGDrink.lblSkyJuice & Space(63) & frmGDrink.lblPriceWater & Space(22) & frmGDrink.txtSky & Space(20) & FormatCurrency(gcurSkyPrice, 2))
End If
If frmGDrink.chkOrange.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGDrink.lblOrange & Space(55) & frmGDrink.lblPriceOrange & Space(20) & frmGDrink.txtOrange & Space(20) & FormatCurrency(gcurOrangePrice, 2))
End If
If frmGDrink.chkStrawberry.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGDrink.lblDetox & Space(31) & frmGDrink.lblPriceStrawberry & Space(20) & frmGDrink.txtStrawberry & Space(20) & FormatCurrency(gcurStrawberryPrice, 2))
End If


'Print dessert to receipt'
If frmGDessert.chkPisang.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGDessert.lblPisang & Space(22) & frmGDessert.lblPricePisang & Space(20) & frmGDessert.txtPisang & Space(20) & FormatCurrency(gcurPisangPrice, 2))
End If
If frmGDessert.chkWaffle.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGDessert.lblWaffle & Space(22) & frmGDessert.lblPriceWaffle & Space(20) & frmGDessert.txtWaffle & Space(20) & FormatCurrency(gcurWafflePrice, 2))
End If
If frmGDessert.chkEclairs.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGDessert.lblChocolate & Space(29) & frmGDessert.lblPriceChocolate & Space(20) & frmGDessert.txtEclairs & Space(20) & FormatCurrency(gcurEclairsPrice, 2))
End If
If frmGDessert.chkCake.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGDessert.lblCake & Space(38) & frmGDessert.lblPriceCake & Space(20) & frmGDessert.txtCake & Space(20) & FormatCurrency(gcurCakePrice, 2))
End If
If frmGDessert.chkCorndog.Value = 1 Then
frmGCheckout.lstReceipt.AddItem (frmGDessert.lblCornDog & Space(48) & frmGDessert.lblPriceCorndog & Space(20) & frmGDessert.txtCorndog & Space(20) & FormatCurrency(gcurCorndogPrice, 2))
End If

frmGCheckout.lblGTotal.Caption = FormatCurrency(gcurTotalCust, 2)


frmGMenu.Hide
frmGCheckout.Show


End Sub

Private Sub cmdReset_Click()

frmGMenu.imgFood.Enabled = True
frmGMenu.imgDrink.Enabled = True
frmGMenu.imgDessert.Enabled = True

gcurTotalCust = 0
gcurTotalCustCart = 0
lblTotalPayment.Caption = " "

'Clear Dessert data if user click'
mintcurTotalPriceDessert = 0

mcurPisangPrice = 0
gcurPisangPrice = 0
gcurAllPisangPrice = 0
gintPisangQuantity = 0
   
mcurWafflePrice = 0
gcurWafflePrice = 0
gcurAllWafflePrice = 0
gintWaffleQuantity = 0
   
mcurEclairsPrice = 0
gcurEclairsPrice = 0
gcurAllEclairsPrice = 0
gintEclairsQuantity = 0
   
mcurCakePrice = 0
gcurCakePrice = 0
gcurAllCakePrice = 0
gintCakeQuantity = 0
   
mcurCorndogPrice = 0
gcurCorndogPrice = 0
gcurAllCorndogPrice = 0
gintCorndogQuantity = 0
   

'Clear dessert form'
frmGDessert.txtPisang.Text = " "
frmGDessert.chkPisang.Value = 0
frmGDessert.txtWaffle.Text = " "
frmGDessert.chkWaffle.Value = 0
frmGDessert.txtEclairs.Text = " "
frmGDessert.chkEclairs.Value = 0
frmGDessert.txtCake.Text = " "
frmGDessert.chkCake.Value = 0
frmGDessert.txtCorndog.Text = " "
frmGDessert.chkCorndog.Value = 0


'Clear Drink data if user click'
mintcurTotalPriceDrink = 0

mcurOreoPrice = 0
gcurOreoPrice = 0
gcurAllOreoPrice = 0
gintOreoQuantity = 0
   
mcurWatermelonPrice = 0
gcurWatermelonPrice = 0
gcurAllWatermelonPrice = 0
gintWatermelonQuantity = 0

mcurSkyPrice = 0
gcurSkyPrice = 0
gcurAllSkyPrice = 0
gintSkyQuantity = 0

mcurOrangePrice = 0
gcrOrangePrice = 0
gcurAllOrangePrice = 0
gintOrangeQuantity = 0
   
mcurStrawberryPrice = 0
gcurStrawberryPrice = 0
gcurAllStrawberryPrice = 0
gintStrawberryQuantity = 0


'Clear drink form'
frmGDrink.txtOreo.Text = " "
frmGDrink.chkOreo.Value = 0
frmGDrink.txtWatermelon.Text = " "
frmGDrink.chkWatermelon.Value = 0
frmGDrink.txtSky.Text = " "
frmGDrink.chkSky.Value = 0
frmGDrink.txtOrange.Text = " "
frmGDrink.chkOrange.Value = 0
frmGDrink.txtStrawberry.Text = " "
frmGDrink.chkStrawberry.Value = 0

'Clear Food data if user click'
mintcurTotalPriceFood = 0

mcurBurgerPrice = 0
gcurBurgerPrice = 0
gcurAllBurgerPrice = 0
gintBurgerQuantity = 0

mcurNoodlesPrice = 0
gcurNoodlesPrice = 0
gcurAllNoodlesPrice = 0
gintNoodlesQuantity = 0

mcurNasiLemakPrice = 0
gcurNasiLemakPrice = 0
gcurAllNasiLemakPrice = 0
gintNasiLemakQuantity = 0

mcurChickenPrice = 0
gcurChickenPrice = 0
gcurAllChickenPrice = 0
gintChickenQuantity = 0


End Sub

Private Sub imgDessert_Click()

frmGMenu.Hide
frmGDessert.Show

End Sub


Private Sub imgDrink_Click()

frmGMenu.Hide
frmGDrink.Show

End Sub

Private Sub imgFood_Click()

frmGMenu.Hide
frmGFood.Show
End Sub


Private Sub Form_Activate()

lblTotalPayment.Caption = FormatCurrency(gcurTotalCust, 2)

End Sub

