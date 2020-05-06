VERSION 5.00
Begin VB.Form frmGDessert 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form3"
   ClientHeight    =   10080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18255
   Picture         =   "FormGDessert.frx":0000
   ScaleHeight     =   10080
   ScaleWidth      =   18255
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   0
      Width           =   2055
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
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdContDessert 
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
      Height          =   855
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9600
      Width           =   2535
   End
   Begin VB.TextBox txtCake 
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
      Left            =   5400
      TabIndex        =   18
      Top             =   10320
      Width           =   495
   End
   Begin VB.TextBox txtEclairs 
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
      Width           =   495
   End
   Begin VB.TextBox txtCorndog 
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
      Left            =   12120
      TabIndex        =   14
      Top             =   10200
      Width           =   495
   End
   Begin VB.CheckBox chkCorndog 
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
      Left            =   13440
      TabIndex        =   13
      Top             =   9360
      Width           =   285
   End
   Begin VB.CheckBox chkEclairs 
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
      Left            =   17280
      TabIndex        =   12
      Top             =   4080
      Width           =   285
   End
   Begin VB.CheckBox chkCake 
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
      Top             =   9480
      Width           =   285
   End
   Begin VB.TextBox txtPisang 
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
      Left            =   4440
      TabIndex        =   3
      Top             =   5280
      Width           =   495
   End
   Begin VB.CheckBox chkPisang 
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
      Left            =   5400
      TabIndex        =   2
      Top             =   4080
      Width           =   285
   End
   Begin VB.CheckBox chkWaffle 
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
      Left            =   11760
      TabIndex        =   1
      Top             =   4080
      Width           =   285
   End
   Begin VB.TextBox txtWaffle 
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
      TabIndex        =   0
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label lblPriceCorndog 
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
      Left            =   10200
      TabIndex        =   27
      Top             =   9840
      Width           =   2175
   End
   Begin VB.Label lblPriceWaffle 
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
      TabIndex        =   26
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label lblPriceCake 
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
      Left            =   3480
      TabIndex        =   25
      Top             =   9960
      Width           =   2415
   End
   Begin VB.Label lblPriceChocolate 
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
      TabIndex        =   24
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label lblPricePisang 
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
      Left            =   2520
      TabIndex        =   23
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   3705
      Left            =   14040
      Picture         =   "FormGDessert.frx":A5D9E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3795
   End
   Begin VB.Label lblQuantityCake 
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
      Left            =   3480
      TabIndex        =   19
      Top             =   10440
      Width           =   1815
   End
   Begin VB.Label lblQuantityChocolate 
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
   Begin VB.Label lblQuanityCorndog 
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
      Left            =   10200
      TabIndex        =   15
      Top             =   10320
      Width           =   1695
   End
   Begin VB.Image imgChickenGrill 
      Height          =   3495
      Left            =   9840
      Picture         =   "FormGDessert.frx":CB209
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   4425
   End
   Begin VB.Image imgSpageti 
      Height          =   3495
      Left            =   3600
      Picture         =   "FormGDessert.frx":EBC11
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   4425
   End
   Begin VB.Image imgMee 
      Height          =   4215
      Left            =   7680
      Picture         =   "FormGDessert.frx":245607
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3825
   End
   Begin VB.Image imgBurger 
      Height          =   3615
      Left            =   2400
      Picture         =   "FormGDessert.frx":29F965
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4515
   End
   Begin VB.Label lblWaffle 
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
      Height          =   735
      Left            =   7560
      TabIndex        =   10
      Top             =   4080
      Width           =   4095
   End
   Begin VB.Label lblChocolate 
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
      Height          =   735
      Left            =   13560
      TabIndex        =   9
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label lblCake 
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
      Left            =   3480
      TabIndex        =   8
      Top             =   9480
      Width           =   4095
   End
   Begin VB.Label lblCorndog 
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
      Left            =   10200
      TabIndex        =   7
      Top             =   9360
      Width           =   3375
   End
   Begin VB.Label lblQuantityPisang 
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
      Left            =   2520
      TabIndex        =   6
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label lblPisang 
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
      Height          =   735
      Left            =   2520
      TabIndex        =   5
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label lblQuantityWaffle 
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
      TabIndex        =   4
      Top             =   5400
      Width           =   1815
   End
End
Attribute VB_Name = "frmGDessert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mcurPisangPrice As Currency
Dim mcurWafflePrice As Currency
Dim mcurEclairsPrice As Currency
Dim mcurCakePrice As Currency
Dim mcurCorndogPrice As Currency

Dim mintcurTotalPriceDessert As Currency

Private Sub cmdBack_Click()

frmGDessert.Hide

frmGMenu.Show

End Sub

Private Sub cmdContDessert_Click()

'To calculate the price of the dessert

'PISANG GORENG CHEESE
If chkPisang.Value = 1 Then
  'To calculate the price for each customer using module variable
   mcurPisangPrice = 15 * Val(txtPisang.Text)
   
  'To calculate the price and quantity for all customer accumulate the prices using global variable
   gcurPisangPrice = gcurPisangPrice + mcurPisangPrice
   gcurAllPisangPrice = gcurAllPisangPrice + gcurPisangPrice
   gintPisangQuantity = gintPisangQuantity + Val(txtPisang.Text)
   
   
Else
   mcurPisangPrice = 0
   gcurPisangPrice = 0
   gcurAllPisangPrice = 0
   gintPisangQuantity = 0
End If

'ICE CREAM WAFFLE
If chkWaffle.Value = 1 Then
  'To calculate the price for each customer using module variable
   mcurWafflePrice = 20 * Val(txtWaffle.Text)
   
  'To calculate the price and quantity for all customer accumulate the prices using global variable
   gcurWafflePrice = gcurWafflePrice + mcurWafflePrice
   gcurAllWafflePrice = gcurAllWafflePrice + gcurWafflePrice
   gintWaffleQuantity = gintWaffleQuantity + Val(txtWaffle.Text)
   
   
Else
   mcurWafflePrice = 0
   gcurWafflePrice = 0
   gcurAllWafflePrice = 0
   gintWaffleQuantity = 0
End If

'CHOCLATE ECLAIRS
If chkEclairs.Value = 1 Then
  'To calculate the price for each customer using module variable
   mcurEclairsPrice = 20 * Val(txtEclairs.Text)
   
  'To calculate the price and quantity for all customer accumulate the prices using global variable
   gcurEclairsPrice = gcurEclairsPrice + mcurEclairsPrice
   gcurAllEclairsPrice = gcurAllEclairsPrice + gcurEclairsPrice
   gintEclairsQuantity = gintEclairsQuantity + Val(txtEclairs.Text)
   
   
Else
   mcurEclairsPrice = 0
   gcurEclairsPrice = 0
   gcurAllEclairsPrice = 0
   gintEclairsQuantity = 0
End If

'MACPAS CAKE
If chkCake.Value = 1 Then
  'To calculate the price for each customer using module variable
   mcurCakePrice = 15.5 * Val(txtCake.Text)
   
  'To calculate the price and quantity for all customer accumulate the prices using global variable
   gcurCakePrice = gcurCakePrice + mcurCakePrice
   gcurAllCakePrice = gcurAllCakePrice + gcurCakePrice
   gintCakeQuantity = gintCakeQuantity + Val(txtCake.Text)
   
   
Else
   mcurCakePrice = 0
   gcurCakePrice = 0
   gcurAllCakePrice = 0
   gintCakeQuantity = 0
End If

'CORNDOG
If chkCorndog.Value = 1 Then
  'To calculate the price for each customer using module variable
   mcurCorndogPrice = 25 * Val(txtCorndog.Text)
   
  'To calculate the price and quantity for all customer accumulate the prices using global variable
   gcurCorndogPrice = gcurCorndogPrice + mcurCorndogPrice
   gcurAllCorndogPrice = gcurAllCorndogPrice + gcurCorndogPrice
   gintCorndogQuantity = gintCorndogQuantity + Val(txtCorndog.Text)
   
   
Else
   mcurCorndogPrice = 0
   gcurCorndogPrice = 0
   gcurAllCorndogPrice = 0
   gintCorndogQuantity = 0
End If


'To calculate total price for food'
mintcurTotalPriceDessert = totalPriceItem(mcurPisangPrice, mcurWafflePrice, mcurEclairsPrice, mcurCakePrice, mcurCorndogPrice)

'To display the total price for the customer
MsgBox "The total Price is " & FormatCurrency(mintcurTotalPriceDessert, 2), vbOKOnly

'To calculate total price for customer'
gcurTotalCust = gcurTotalCust + mintcurTotalPriceDessert

gcurTotalCustCart = gcurTotalCustCart + gcurTotalCust
frmGMenu.imgDessert.Enabled = False

frmGDessert.Hide
frmGMenu.Show
End Sub



Private Sub cmdExit_Click()
End
End Sub

