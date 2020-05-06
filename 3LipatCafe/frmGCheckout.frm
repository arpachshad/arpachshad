VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGCheckout 
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15570
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmGCheckout.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   15570
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frmGTotal 
      Height          =   855
      Left            =   15000
      TabIndex        =   2
      Top             =   7800
      Width           =   1815
      Begin VB.Label lblGTotal 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc customer 
      Height          =   330
      Left            =   15240
      Top             =   9360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\habib\Desktop\Project NewVb\3LipatTest.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\habib\Desktop\Project NewVb\3LipatTest.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "customer"
      Caption         =   "customer"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ListBox lstReceipt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7620
      Left            =   2640
      TabIndex        =   1
      Top             =   1200
      Width           =   16695
   End
   Begin VB.CommandButton cmdConfirm 
      BackColor       =   &H00FF00FF&
      Caption         =   "Confirm"
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9240
      Width           =   5055
   End
   Begin MSAdodcLib.Adodc order 
      Height          =   330
      Left            =   15360
      Top             =   10080
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   2
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\habib\Desktop\Project NewVb\3LipatTest.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\habib\Desktop\Project NewVb\3LipatTest.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ord"
      Caption         =   "order"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmGCheckout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfirm_Click()


'To declare variable
Dim strName As String
Dim intPax As Integer
Dim curTotal As Currency


'To write the data into text file
'If use append it will continue storing data
'If use Output it will overwrite the data to the new Data
Open "C:\Users\habib\Desktop\Project NewVb\customer.txt" For Append As #1

'To check the loop
strName = frmGDetail.txtCustName.Text
intPax = Val(frmGDetail.txtCustPax.Text)
curTotal = frmGMenu.lblTotalPayment.Caption

    Write #1, strName, intPax, curTotal
   


Close #1

'Insert customer detail into database
With customer.Recordset
customer.Refresh
customer.Recordset.AddNew
customer.Recordset.Fields("cust_name").Value = frmGDetail.txtCustName
customer.Recordset.Fields("cust_pax").Value = frmGDetail.txtCustPax
customer.Recordset.Fields("cust_totalPay").Value = gcurTotalCust
customer.Recordset.Update
End With

'Insert order detail into database
With order.Recordset
order.Refresh
order.Recordset.AddNew
order.Recordset.Fields("order_grandTotal").Value = gcurTotalCust
order.Recordset.Fields("order_quantity").Value = Val(frmGFood.txtBurger.Text) + Val(frmGFood.txtNoodle.Text) + Val(frmGFood.txtNasiLemak.Text) + Val(frmGFood.txtChicken.Text) + Val(frmGFood.txtSpagetti.Text) + Val(frmGDrink.txtOreo.Text) + Val(frmGDrink.txtWatermelon.Text) + Val(frmGDrink.txtSky.Text) + Val(frmGDrink.txtOrange.Text) + Val(frmGDrink.txtStrawberry.Text) + Val(frmGDessert.txtPisang.Text) + Val(frmGDessert.txtWaffle.Text) + Val(frmGDessert.txtEclairs.Text) + Val(frmGDessert.txtCake.Text) + Val(frmGDessert.txtCorndog.Text)
order.Recordset.Fields("cust_name").Value = frmGDetail.txtCustName
order.Recordset.Update
End With

'To back to the main menu
frmGCheckout.Hide
frmWelcome.Show

'To clear previous output
lstReceipt.Clear




End Sub

Private Sub Form_Load()
lstReceipt.AddItem "NAME" & Space(72) & "PRICE" & Space(13) & "QUANTITY" & Space(15) & "TOTAL"
End Sub

