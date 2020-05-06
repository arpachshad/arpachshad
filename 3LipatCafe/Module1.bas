Attribute VB_Name = "Module1"
Option Explicit

'To declare global variables
'Variables for food menus
Public gintBurgerQuantity As Integer
Public gcurBurgerPrice As Currency
Public gcurAllBurgerPrice As Currency


Public gintNoodlesQuantity As Integer
Public gcurNoodlesPrice As Currency
Public gcurAllNoodlesPrice As Currency

Public gintNasiLemakQuantity As Integer
Public gcurNasiLemakPrice As Currency
Public gcurAllNasiLemakPrice As Currency

Public gintSpagettiQuantity As Integer
Public gcurSpagettiPrice As Currency
Public gcurAllSpagettiPrice As Currency

Public gintChickenQuantity As Integer
Public gcurChickenPrice As Currency
Public gcurAllChickenPrice As Currency

Public gcurAllFoodSale As Currency

'Variables for beverages menus
Public gintOreoQuantity As Integer
Public gcurOreoPrice As Currency
Public gcurAllOreoPrice As Currency

Public gintWatermelonQuantity As Integer
Public gcurWatermelonPrice As Currency
Public gcurAllWatermelonPrice As Currency

Public gintSkyQuantity As Integer
Public gcurSkyPrice As Currency
Public gcurAllSkyPrice As Currency

Public gintOrangeQuantity As Integer
Public gcurOrangePrice As Currency
Public gcurAllOrangePrice As Currency

Public gintStrawberryQuantity As Integer
Public gcurStrawberryPrice As Currency
Public gcurAllStrawberryPrice As Currency

Public gcurAllBeverageSale As Currency


'Variables for dessert menus
Public gintPisangQuantity As Integer
Public gcurPisangPrice As Currency
Public gcurAllPisangPrice As Currency

Public gintWaffleQuantity As Integer
Public gcurWafflePrice As Currency
Public gcurAllWafflePrice As Currency

Public gintEclairsQuantity As Integer
Public gcurEclairsPrice As Currency
Public gcurAllEclairsPrice As Currency

Public gintCakeQuantity As Integer
Public gcurCakePrice As Currency
Public gcurAllCakePrice As Currency

Public gintCorndogQuantity As Integer
Public gcurCorndogPrice As Currency
Public gcurAllCorndogPrice As Currency

Public gcurAllDessertSale As Currency



'To declare calculation variables for customer

Public gcurTotalFood As Currency
Public gcurGSTFood As Currency
Public gcurTotalGSTFood As Currency

Public gcurTotalBeverage As Currency
Public gcurGSTBeverage As Currency
Public gcurTotalGSTBeverage As Currency

Public gcurTotalDessert As Currency
Public gcurGSTDessert As Currency
Public gcurTotalGSTDessert As Currency

Public gcurTotalCust As Currency
Public gcurTotalAll As Currency


'To declare calculation variable for staff record

Public gcurStaffGST As Currency
Public gcurStaffProfit As Currency
Public gcurStaffTotalFood As Currency
Public gcurStaffTotalBeverage As Currency
Public gcurStaffTotalDessert As Currency

'Function'
Public Function totalPriceItem(ByVal item1 As Currency, ByVal item2 As Currency, ByVal item3 As Currency, ByVal item4 As Currency, ByVal item5 As Currency) As Currency
'Calculate total price customer'

totalPriceItem = item1 + item2 + item3 + item4 + item5

End Function





