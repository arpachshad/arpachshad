VERSION 5.00
Begin VB.Form frmGDetail 
   Caption         =   "Guest Page"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18900
   LinkTopic       =   "Form1"
   Picture         =   "frmGDetail.frx":0000
   ScaleHeight     =   8055
   ScaleWidth      =   18900
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frmCust 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fill in the details below"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   7560
      TabIndex        =   0
      Top             =   2880
      Width           =   7455
      Begin VB.ComboBox cboOrderType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "frmGDetail.frx":2BA5E
         Left            =   3000
         List            =   "frmGDetail.frx":2BA68
         TabIndex        =   6
         Top             =   3240
         Width           =   3975
      End
      Begin VB.CommandButton cmdSubmit 
         BackColor       =   &H000080FF&
         Caption         =   "Submit"
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
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5520
         Width           =   2535
      End
      Begin VB.TextBox txtCustPax 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   4
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox txtCustName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   1440
         Width           =   3855
      End
      Begin VB.Label lblOrderType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Take away/ Eat here"
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
         Left            =   960
         TabIndex        =   7
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblCustPax 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No. of Pax"
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
         Left            =   960
         TabIndex        =   2
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblCustName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name"
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
         Left            =   960
         TabIndex        =   1
         Top             =   1440
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmGDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSubmit_Click()

'Reset the price for new customer'
gcurTotalCust = 0
gcurTotalCustCart = 0

gcurBurgerPrice = 0
gcurNoodlesPrice = 0
gcurNasiLemakPrice = 0
gcurChickenPrice = 0
gcurSpagettiPrice = 0

gcurOreoPrice = 0
gcurWatermelonPrice = 0
gcurSkyPrice = 0
gcurOrangePrice = 0
gcurStrawberryPrice = 0

gcurPisangPrice = 0
gcurWafflePrice = 0
gcurEclairsPrice = 0
gcurCakePrice = 0
gcurCorndogPrice = 0


'To check if the information is inserted correctly

 ' If txtCustName.Text = "" Then'
 '     MsgBox "Please enter your name", vbOKOnly + vbInformation, "Error"'
'  End If'

  ' If txtCustPax.Text = "" Then'
     ' MsgBox "Please enter the number pax", vbOKOnly + vbInformation, "Error" '
   
 '  Else'

    '  If IsNumeric(txtCustPax.Text) Then'
     ' Else'
      '   MsgBox "Non numeric data entered", vbOKOnly, "Invalid Data"'
     ' End If'
 '
     
'To hide the current form frmGuest
frmGDetail.Hide

'To show or Navigate the next form, formMenu
frmGMenu.Show

'End If'


'Clear cart'
gcurTotalCust = 0


'Retrieve all item data from text file'
Open "C:\Users\habib\Desktop\Project NewVb\item.txt" For Input As #1

Do Until EOF(1)

Input #1, itemId, itemName, itemPrice, totalPrice

If (itemId = 1) Then
frmGFood.lblBurger = itemName
frmGFood.lblPriceBurger = itemPrice

ElseIf (itemId = 2) Then
frmGFood.lblSpagetti = itemName
frmGFood.lblPriceSpagetti = itemPrice

ElseIf (itemId = 3) Then
frmGFood.lblNasiLemak = itemName
frmGFood.lblPriceNasiLemak = itemPrice

ElseIf (itemId = 4) Then
frmGFood.lblChickenGrill = itemName
frmGFood.lblPriceChickenGrill = itemPrice

ElseIf (itemId = 5) Then
frmGFood.lblMee = itemName
frmGFood.lblPriceMeeGoreng = itemPrice

ElseIf (itemId = 6) Then
frmGDrink.lblDetox.Caption = itemName
frmGDrink.lblPriceStrawberry.Caption = itemPrice

ElseIf (itemId = 7) Then
frmGDrink.lblOreo.Caption = itemName
frmGDrink.lblPriceOreo.Caption = itemPrice

ElseIf (itemId = 8) Then
frmGDrink.lblSkyJuice.Caption = itemName
frmGDrink.lblPriceWater.Caption = itemPrice

ElseIf (itemId = 9) Then
frmGDrink.lblOrange.Caption = itemName
frmGDrink.lblPriceOrange.Caption = itemPrice

ElseIf (itemId = 10) Then
frmGDrink.lblWatermelon.Caption = itemName
frmGDrink.lblPriceWaterMelon.Caption = itemPrice

ElseIf (itemId = 11) Then
frmGDessert.lblPisang = itemName
frmGDessert.lblPricePisang = itemPrice

ElseIf (itemId = 12) Then
frmGDessert.lblWaffle = itemName
frmGDessert.lblPriceWaffle = itemPrice

ElseIf (itemId = 13) Then
frmGDessert.lblChocolate = itemName
frmGDessert.lblPriceChocolate = itemPrice

ElseIf (itemId = 14) Then
frmGDessert.lblCake = itemName
frmGDessert.lblPriceCake = itemPrice

ElseIf (itemId = 15) Then
frmGDessert.lblCornDog = itemName
frmGDessert.lblPriceCorndog = itemPrice


End If

Loop

Close #1

   
End Sub

