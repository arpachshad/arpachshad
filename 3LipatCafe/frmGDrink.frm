VERSION 5.00
Begin VB.Form frmGDrink 
   Caption         =   "Form1"
   ClientHeight    =   9930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15915
   LinkTopic       =   "Form1"
   Picture         =   "frmGDrink.frx":0000
   ScaleHeight     =   9930
   ScaleWidth      =   15915
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
      Left            =   18720
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   0
      Width           =   2055
   End
   Begin VB.TextBox txtStrawberry 
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
      Left            =   12600
      TabIndex        =   19
      Top             =   9960
      Width           =   615
   End
   Begin VB.CheckBox chkStrawberry 
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
      Left            =   15600
      TabIndex        =   18
      Top             =   9240
      Width           =   285
   End
   Begin VB.TextBox txtOrange 
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
      Left            =   4800
      TabIndex        =   16
      Top             =   9960
      Width           =   615
   End
   Begin VB.CheckBox chkOrange 
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
      TabIndex        =   15
      Top             =   9120
      Width           =   285
   End
   Begin VB.TextBox txtSky 
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
      Left            =   17040
      TabIndex        =   13
      Top             =   4920
      Width           =   615
   End
   Begin VB.CheckBox chkSky 
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
      Left            =   17160
      TabIndex        =   12
      Top             =   4080
      Width           =   285
   End
   Begin VB.TextBox txtWatermelon 
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
      Left            =   10320
      TabIndex        =   10
      Top             =   4920
      Width           =   615
   End
   Begin VB.CheckBox chkWatermelon 
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
      Left            =   11880
      TabIndex        =   9
      Top             =   4080
      Width           =   285
   End
   Begin VB.TextBox txtOreo 
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
      TabIndex        =   7
      Top             =   4920
      Width           =   615
   End
   Begin VB.CheckBox chkOreo 
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
      Top             =   4080
      Width           =   285
   End
   Begin VB.CommandButton cmdContDrink 
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
      Left            =   16320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9720
      Width           =   2535
   End
   Begin VB.Label lblOrange 
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
      Left            =   2880
      TabIndex        =   4
      Top             =   9120
      Width           =   2895
   End
   Begin VB.Label lblPriceOrange 
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
      Left            =   2880
      TabIndex        =   25
      Top             =   9600
      Width           =   2655
   End
   Begin VB.Label lblPriceStrawberry 
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
      Left            =   10800
      TabIndex        =   23
      Top             =   9600
      Width           =   2175
   End
   Begin VB.Label lblDetox 
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
      Left            =   10800
      TabIndex        =   5
      Top             =   9120
      Width           =   4575
   End
   Begin VB.Label lblPriceWater 
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
      Left            =   15240
      TabIndex        =   26
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label lblPriceWaterMelon 
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
      Left            =   8400
      TabIndex        =   27
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label lblPriceOreo 
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
      Left            =   2520
      TabIndex        =   24
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label lblQuantityDetox 
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
      Left            =   10800
      TabIndex        =   20
      Top             =   10080
      Width           =   1695
   End
   Begin VB.Label lblQuantityOrange 
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
      Left            =   2880
      TabIndex        =   17
      Top             =   10080
      Width           =   1695
   End
   Begin VB.Label lblQuantitySkyJuice 
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
      Left            =   15240
      TabIndex        =   14
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label lblQuanityWatermelon 
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
      Left            =   8400
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label lblQuanityOreoMilk 
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
      TabIndex        =   8
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label lblWatermelon 
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
      Left            =   8400
      TabIndex        =   2
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label lblOreo 
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
      Left            =   2520
      TabIndex        =   1
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Image Image4 
      Height          =   3495
      Left            =   11280
      Picture         =   "frmGDrink.frx":A5D9E
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   3825
   End
   Begin VB.Image Image3 
      Height          =   3735
      Left            =   2760
      Picture         =   "frmGDrink.frx":BED31
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   3075
   End
   Begin VB.Image imgOreoMilk 
      Height          =   3855
      Left            =   2520
      Picture         =   "frmGDrink.frx":F1448
      Stretch         =   -1  'True
      Top             =   360
      Width           =   3435
   End
   Begin VB.Image Image2 
      Height          =   4215
      Left            =   14880
      Picture         =   "frmGDrink.frx":13218D
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3465
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   8400
      Picture         =   "frmGDrink.frx":154A92
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3705
   End
   Begin VB.Label lblSkyJuice 
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
      Left            =   15240
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
End
Attribute VB_Name = "frmGDrink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mcurOreoPrice As Currency
Dim mcurWatermelonPrice As Currency
Dim mcurSkyPrice As Currency
Dim mcurOrangePrice As Currency
Dim mcurStrawberryPrice As Currency

Dim mintcurTotalPriceBeverage As Currency

Private Sub cmdBack_Click()

frmGDrink.Hide

frmGMenu.Show

End Sub

Private Sub cmdContDrink_Click()

'To calculate the price of the beverage

'OREO MILKSHAKE
If chkOreo.Value = 1 Then
  'To calculate the price for each customer using module variable
   mcurOreoPrice = 15 * Val(txtOreo.Text)
  'To calculate the price and quantity for all customer accumulate the prices using global variable
   gcurOreoPrice = gcurOreoPrice + mcurOreoPrice
   gcurAllOreoPrice = gcurAllOreoPrice + gcurOreoPrice
   gintOreoQuantity = gintOreoQuantity + Val(txtOreo.Text)
   

Else
   mcurOreoPrice = 0
   gcurOreoPrice = 0
   gcurAllOreoPrice = 0
   gintOreoQuantity = 0
End If

'WATERMELON JUICE
If chkWatermelon.Value = 1 Then
  'To calculate the price for each customer using module variable
   mcurWatermelonPrice = 12 * Val(txtWatermelon.Text)
  'To calculate the price and quantity for all customer accumulate the prices using global variable
   gcurWatermelonPrice = gcurWatermelonPrice + mcurWatermelonPrice
   gcurAllWatermelonPrice = gcurAllWatermelonPrice + gcurWatermelonPrice
   gintWatermelonQuantity = gintWatermelonQuantity + Val(txtWatermelon.Text)
   

Else
   mcurWatermelonPrice = 0
   gcurWatermelonPrice = 0
   gcurAllWatermelonPrice = 0
   gintWatermelonQuantity = 0
End If

'SKY JUICE
If chkSky.Value = 1 Then
  'To calculate the price for each customer using module variable
   mcurSkyPrice = 3.5 * Val(txtSky.Text)
  'To calculate the price and quantity for all customer accumulate the prices using global variable
   gcurSkyPrice = gcurSkyPrice + mcurSkyPrice
   gcurAllSkyPrice = gcurAllSkyPrice + gcurSkyPrice
   gintSkyQuantity = gintSkyQuantity + Val(txtSky.Text)
   

Else
   mcurSkyPrice = 0
   gcurSkyPrice = 0
   gcurAllSkyPrice = 0
   gintSkyQuantity = 0
End If

'ORANGE JUICE
If chkOrange.Value = 1 Then
  'To calculate the price for each customer using module variable
   mcurOrangePrice = 12 * Val(txtOrange.Text)
  'To calculate the price and quantity for all customer accumulate the prices using global variable
   gcurOrangePrice = gcurOrangePrice + mcurOrangePrice
   gcurAllOrangePrice = gcurAllOrangePrice + gcurOrangePrice
   gintOrangeQuantity = gintOrangeQuantity + Val(txtOrange.Text)
   

Else
   mcurOrangePrice = 0
   gcurOrangePrice = 0
   gcurAllOrangePrice = 0
   gintOrangeQuantity = 0
End If

'STRAWBERRY FRIZZY
If chkStrawberry.Value = 1 Then
  'To calculate the price for each customer using module variable
   mcurStrawberryPrice = 18.9 * Val(txtStrawberry.Text)
  'To calculate the price and quantity for all customer accumulate the prices using global variable
   gcurStrawberryPrice = gcurStrawberryPrice + mcurStrawberryPrice
   gcurAllStrawberryPrice = gcurAllStrawberryPrice + gcurStrawberryPrice
   gintStrawberryQuantity = gintStrawberryQuantity + Val(txtStrawberry.Text)
   

Else
   mcurStrawberryPrice = 0
   gcurStrawberryPrice = 0
   gcurAllStrawberryPrice = 0
   gintStrawberryQuantity = 0
End If


'To calculate total price for beverage'
mintcurTotalPriceBeverage = totalPriceItem(mcurOreoPrice, mcurWatermelonPrice, mcurSkyPrice, mcurOrangePrice, mcurStrawberryPrice)

'To display the total price for the customer
MsgBox "The total Price is " & FormatCurrency(mintcurTotalPriceBeverage, 2), vbOKOnly

'To calculate total price for customer'
gcurTotalCust = gcurTotalCust + mintcurTotalPriceBeverage

gcurTotalCustCart = gcurTotalCustCart + gcurTotalCust
frmGMenu.imgDrink.Enabled = False

frmGDrink.Hide
frmGMenu.Show

End Sub


Private Sub cmdExit_Click()
End
End Sub

