VERSION 5.00
Begin VB.Form frmSListOfDrinks 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
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
      Left            =   17280
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9600
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label lblPriceWaterMelon 
      Alignment       =   2  'Center
      Caption         =   "RM 12.00 net"
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
      Left            =   11160
      TabIndex        =   10
      Top             =   10080
      Width           =   2655
   End
   Begin VB.Label lblPriceWater 
      Alignment       =   2  'Center
      Caption         =   "RM 3.50 net"
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
      Left            =   5040
      TabIndex        =   9
      Top             =   10200
      Width           =   2535
   End
   Begin VB.Label lblPriceOrange 
      Alignment       =   2  'Center
      Caption         =   "RM 12.00 net"
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
      Left            =   14280
      TabIndex        =   8
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label lblPriceOreo 
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
      Height          =   495
      Left            =   8040
      TabIndex        =   7
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label lblPriceStrawberry 
      Alignment       =   2  'Center
      Caption         =   "RM 18.90 net"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label lblWater 
      Alignment       =   2  'Center
      Caption         =   "SKY JUICY"
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
      Left            =   4320
      TabIndex        =   5
      Top             =   9720
      Width           =   3855
   End
   Begin VB.Label lblWatermelon 
      Alignment       =   2  'Center
      Caption         =   "WATERMELON JUICY"
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
      Left            =   10680
      TabIndex        =   4
      Top             =   9600
      Width           =   3615
   End
   Begin VB.Label lblOrange 
      Alignment       =   2  'Center
      Caption         =   "ORANGE JUICY"
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
      Left            =   13920
      TabIndex        =   3
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label lblOrea 
      Alignment       =   2  'Center
      Caption         =   "OREO MILK SHAKE"
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
      TabIndex        =   2
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label lblStrawberry 
      Alignment       =   2  'Center
      Caption         =   "STRAWBERRY FIZZY SPLASH"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   5280
      Width           =   3855
   End
   Begin VB.Image imgWatermelon 
      Height          =   5415
      Left            =   10560
      Picture         =   "frmSListOfDrink.frx":0000
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   3735
   End
   Begin VB.Image imgWater 
      Height          =   4335
      Left            =   4320
      Picture         =   "frmSListOfDrink.frx":2EA2B
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   3855
   End
   Begin VB.Image imgOrange 
      Height          =   4695
      Left            =   13200
      Picture         =   "frmSListOfDrink.frx":51330
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Image imgOrea 
      Height          =   3975
      Left            =   7320
      Picture         =   "frmSListOfDrink.frx":83A47
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Image imgStrawberry 
      Height          =   3735
      Left            =   1080
      Picture         =   "frmSListOfDrink.frx":C478C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label lblListOfDrink 
      Caption         =   "LIST OF DRINKS :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmSListOfDrinks"
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
