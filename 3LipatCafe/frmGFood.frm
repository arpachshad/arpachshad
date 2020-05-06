VERSION 5.00
Begin VB.Form frmGFood 
   BackColor       =   &H00808080&
   Caption         =   "Form2"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18465
   LinkTopic       =   "Form2"
   Picture         =   "frmGFood.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   18465
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cmdContFood 
      BackColor       =   &H00FF80FF&
      Caption         =   "CONTINUE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   15480
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9960
      Width           =   3135
   End
   Begin VB.TextBox txtSpagetti 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3480
      TabIndex        =   18
      Top             =   10200
      Width           =   615
   End
   Begin VB.TextBox txtNasiLemak 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   15480
      TabIndex        =   16
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox txtChicken 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   11760
      TabIndex        =   14
      Top             =   10200
      Width           =   615
   End
   Begin VB.CheckBox chkChicken 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14640
      TabIndex        =   13
      Top             =   9360
      Width           =   285
   End
   Begin VB.CheckBox chkNasiLemak 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   17520
      TabIndex        =   12
      Top             =   4440
      Width           =   285
   End
   Begin VB.CheckBox chkSpagetti 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7680
      TabIndex        =   11
      Top             =   9360
      Width           =   285
   End
   Begin VB.TextBox txtNoodle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9480
      TabIndex        =   9
      Top             =   5160
      Width           =   615
   End
   Begin VB.CheckBox chkNoodle 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11520
      TabIndex        =   8
      Top             =   4440
      Width           =   285
   End
   Begin VB.CheckBox chkBurger 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5760
      TabIndex        =   6
      Top             =   4440
      Width           =   285
   End
   Begin VB.TextBox txtBurger 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3720
      TabIndex        =   5
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lblPriceSpagetti 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   27
      Top             =   9840
      Width           =   2415
   End
   Begin VB.Label lblPriceNasiLemak 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13560
      TabIndex        =   26
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label lblPriceMeeGoreng 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   25
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblPriceChickenGrill 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   24
      Top             =   9840
      Width           =   2175
   End
   Begin VB.Label lblPriceBurger 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label lblQuantitySpageti 
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   19
      Top             =   10320
      Width           =   1695
   End
   Begin VB.Label lblQuantityNasiLemak 
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13560
      TabIndex        =   17
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label lblQuanityChicken 
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   15
      Top             =   10320
      Width           =   1695
   End
   Begin VB.Label lblQuantityMee 
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label lblBurger 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   4440
      Width           =   3855
   End
   Begin VB.Label lblQuantityBurger 
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label lblChickenGrill 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   3
      Top             =   9360
      Width           =   4695
   End
   Begin VB.Label lblSpagetti 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   9360
      Width           =   6015
   End
   Begin VB.Label lblNasiLemak 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13560
      TabIndex        =   1
      Top             =   4440
      Width           =   3855
   End
   Begin VB.Label lblMee 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   0
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Image imgBurger 
      Height          =   3495
      Left            =   1560
      Picture         =   "frmGFood.frx":A5D9E
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4335
   End
   Begin VB.Image imgMee 
      Height          =   3615
      Left            =   7680
      Picture         =   "frmGFood.frx":CCAE8
      Stretch         =   -1  'True
      Top             =   600
      Width           =   4095
   End
   Begin VB.Image imgNasiLemak 
      Height          =   3735
      Left            =   13320
      Picture         =   "frmGFood.frx":10585A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   4665
   End
   Begin VB.Image imgSpageti 
      Height          =   3375
      Left            =   2280
      Picture         =   "frmGFood.frx":15F55F
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   3705
   End
   Begin VB.Image imgChickenGrill 
      Height          =   3255
      Left            =   9720
      Picture         =   "frmGFood.frx":3E90AF
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   4065
   End
End
Attribute VB_Name = "frmGFood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mcurBurgerPrice As Currency
Dim mcurNoodlesPrice As Currency
Dim mcurNasiLemakPrice As Currency
Dim mcurChickenPrice As Currency
Dim mcurSpagettiPrice As Currency

Dim mintcurTotalPriceFood As Currency
Dim mint
Private Sub cmdBack_Click()

frmGFood.Hide

frmGMenu.Show

End Sub

Private Sub cmdContFood_Click()

'To calculate the price of the food

'BURGER
If chkBurger.Value = 1 Then
 'To calculate the price for each customer using module variable
   mcurBurgerPrice = 23 * Val(txtBurger.Text)
  'To calculate the price and quantity for all customer accumulate the prices using global variable
   gcurBurgerPrice = gcurBurgerPrice + mcurBurgerPrice
   gcurAllBurgerPrice = gcurAllBurgerPrice + gcurBurgerPrice
   gintBurgerQuantity = gintBurgerQuantity + Val(txtBurger.Text)

   
ElseIf chkBurger.Value = 0 Then
   mcurBurgerPrice = 0
   gcurBurgerPrice = 0
   gcurAllBurgerPrice = 0
   gintBurgerQuantity = 0
End If

'NOODLE
If chkNoodle.Value = 1 Then
   mcurNoodlesPrice = 15 * Val(txtNoodle.Text)
   gcurNoodlesPrice = gcurNoodlesPrice + mcurNoodlesPrice
   gcurAllNoodlesPrice = gcurAllNoodlesPrice + gcurNoodlesPrice
   gintNoodlesQuantity = gintNoodlesQuantity + Val(txtNoodle.Text)
   
ElseIf chkNoodle.Value = 0 Then

    mcurNoodlesPrice = 0
    gcurNoodlesPrice = 0
    gcurAllNoodlesPrice = 0
    gintNoodlesQuantity = 0
End If

'NASI LEMAK
If chkNasiLemak.Value = 1 Then
   mcurNasiLemakPrice = 18 * Val(txtNasiLemak.Text)
   gcurNasiLemakPrice = gcurNasiLemakPrice + mcurNasiLemakPrice
   gcurAllNasiLemakPrice = gcurAllNasiLemakPrice + gcurNasiLemakPrice
   gintNasiLemakQuantity = gintNasiLemakQuantity + Val(txtNasiLemak.Text)
   
ElseIf chkNasiLemak.Value = 0 Then
    mcurNasiLemakPrice = 0
    gcurNasiLemakPrice = 0
    gcurAllNasiLemakPrice = 0
    gintNasiLemakQuantity = 0
End If

'CHICKEN GRILLED
If chkChicken.Value = 1 Then
   mcurChickenPrice = 32 * Val(txtChicken.Text)
   gcurChickenPrice = gcurChickenPrice + mcurChickenPrice
   gcurAllChickenPrice = gcurAllChickenPrice + gcurChickenPrice
   gintChickenQuantity = gintChickenQuantity + Val(txtChicken.Text)

ElseIf chkChicken.Value = 0 Then
    mcurChickenPrice = 0
    gcurChickenPrice = 0
    gcurAllChickenPrice = 0
    gintChickenQuantity = 0
End If

'SPAGETTI
If chkSpagetti.Value = 1 Then
   mcurSpagettiPrice = 20 * Val(txtSpagetti.Text)
   gcurSpagettiPrice = gcurSpagettiPrice + mcurSpagettiPrice
   gcurAllSpagettiPrice = gcurAllSpagettiPrice + gcurSpagettiPrice
   gintSpagettiQuantity = gintSpagettiQuantity + Val(txtSpagetti.Text)

ElseIf chkSpagetti.Value = 0 Then
    mcurSpagettiPrice = 0
    gcurSpagettiPrice = 0
    gcurAllSpagettiPrice = 0
    gintSpagettiQuantity = 0
End If

'To calculate total price for food'
mintcurTotalPriceFood = totalPriceItem(mcurBurgerPrice, mcurNoodlesPrice, mcurNasiLemakPrice, mcurChickenPrice, mcurSpagettiPrice)
gcurAllFoodSale = gcurAllFoodSale + mintcurTotalPriceFood

'To calculate total price for customer'
gcurTotalCust = gcurTotalCust + mintcurTotalPriceFood

'To display the total price for the customer
MsgBox "The total Price is " & FormatCurrency(mintcurTotalPriceFood, 2), vbOKOnly


gcurTotalCustCart = gcurTotalCustCart + gcurTotalCust

frmGMenu.imgFood.Enabled = False


frmGFood.Hide
frmGMenu.Show

End Sub

Private Sub cmdExit_Click()
End
End Sub

