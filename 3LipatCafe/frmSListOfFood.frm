VERSION 5.00
Begin VB.Form frmSListOfFood 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   1065
   ClientTop       =   2730
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
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
      TabIndex        =   13
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H0080C0FF&
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9360
      Width           =   2175
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9600
      Width           =   1335
   End
   Begin VB.Label lblPriceSpagetti 
      Alignment       =   2  'Center
      Caption         =   "RM 20.00 net"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   10
      Top             =   10320
      Width           =   2295
   End
   Begin VB.Label lblPriceMeeGoreng 
      Alignment       =   2  'Center
      Caption         =   "RM 15.00 net"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11640
      TabIndex        =   9
      Top             =   9840
      Width           =   2655
   End
   Begin VB.Label lblPriceNasiLemak 
      Alignment       =   2  'Center
      Caption         =   "RM 18.00 net"
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
      Left            =   15240
      TabIndex        =   8
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label lblPriceChickenChop 
      Alignment       =   2  'Center
      Caption         =   "RM 32.00 net"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   7
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label lblMeeGoreng 
      Alignment       =   2  'Center
      Caption         =   "MEE GORENG MONSTER"
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
      Left            =   10320
      TabIndex        =   6
      Top             =   9360
      Width           =   5415
   End
   Begin VB.Label lblSpagetti 
      Alignment       =   2  'Center
      Caption         =   "SPAGETTI BOLOGNESE POSSO FARLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3480
      TabIndex        =   5
      Top             =   9480
      Width           =   5400
   End
   Begin VB.Label lblNasiLemak 
      Alignment       =   2  'Center
      Caption         =   "NASI LEMAK BERGANDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14280
      TabIndex        =   4
      Top             =   5160
      Width           =   3855
   End
   Begin VB.Label lblChickenChop 
      Alignment       =   2  'Center
      Caption         =   "CHICKEN GRILLED DELICIOUS"
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
      Left            =   7200
      TabIndex        =   3
      Top             =   5160
      Width           =   4455
   End
   Begin VB.Label lblPriceBurger 
      Alignment       =   2  'Center
      Caption         =   "RM 23.00 net"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   2
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label lblBurger 
      Alignment       =   2  'Center
      Caption         =   "BURGER TRIPLE TOWER"
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
      Left            =   840
      TabIndex        =   1
      Top             =   5160
      Width           =   4095
   End
   Begin VB.Image imgNasiLemak 
      Height          =   3615
      Left            =   13800
      Picture         =   "frmSListOfFood.frx":0000
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Image imgSpagetti 
      Height          =   3375
      Left            =   4080
      Picture         =   "frmSListOfFood.frx":59D05
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   4095
   End
   Begin VB.Image imgMeeGoreng 
      Height          =   2895
      Left            =   10800
      Picture         =   "frmSListOfFood.frx":2E3855
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   4215
   End
   Begin VB.Image imgChickenChop 
      Height          =   3615
      Left            =   7200
      Picture         =   "frmSListOfFood.frx":31C5C7
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Image imgBurger 
      Height          =   3720
      Left            =   600
      Picture         =   "frmSListOfFood.frx":34E0A5
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   4440
   End
   Begin VB.Label lblFoodList 
      Caption         =   "LIST OF FOOD :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmSListOfFood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdNext_Click()
Me.Hide
frmSMenu.Show
End Sub


Private Sub cmdUpdate_Click()
Me.Hide
frmSUpdate.Show
End Sub

