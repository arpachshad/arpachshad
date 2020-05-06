VERSION 5.00
Begin VB.Form frmSListOfDessert 
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19065
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   19065
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
      Top             =   9480
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9600
      Width           =   1335
   End
   Begin VB.Label lblPricePisang 
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
      Left            =   5040
      TabIndex        =   10
      Top             =   10080
      Width           =   2175
   End
   Begin VB.Label lblPriceWaffle 
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
      Height          =   495
      Left            =   11640
      TabIndex        =   9
      Top             =   10080
      Width           =   2295
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      Caption         =   "RM 25.00 net"
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
      Left            =   14520
      TabIndex        =   8
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label lblPriceChocolate 
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
      Height          =   495
      Left            =   8040
      TabIndex        =   7
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label lblPriceCake 
      Alignment       =   2  'Center
      Caption         =   "RM 15.50 net"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label lblPisang 
      Alignment       =   2  'Center
      Caption         =   "PISANG GORENG CHEESE ( 5 pcs )"
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
      Left            =   3720
      TabIndex        =   5
      Top             =   9600
      Width           =   5175
   End
   Begin VB.Label lblWaffle 
      Alignment       =   2  'Center
      Caption         =   "WAFFLE DOUBLE LAYER with ICE CREAM"
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
      TabIndex        =   4
      Top             =   9600
      Width           =   5775
   End
   Begin VB.Label lblCornDog 
      Alignment       =   2  'Center
      Caption         =   "CORN DOG ( 3 pcs )"
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
      Left            =   13680
      TabIndex        =   3
      Top             =   5400
      Width           =   3975
   End
   Begin VB.Label lblChocolate 
      Alignment       =   2  'Center
      Caption         =   "CHOCOLATE ECLAIRS"
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
      Left            =   7680
      TabIndex        =   2
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label lblCake 
      Alignment       =   2  'Center
      Caption         =   "CAKE MACPAS ( 1 slices )"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   5280
      Width           =   4455
   End
   Begin VB.Image imgWaffle 
      Height          =   4260
      Left            =   10680
      Picture         =   "frmSListOfDessert.frx":0000
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   3840
   End
   Begin VB.Image imgPisang 
      Height          =   3720
      Left            =   4320
      Picture         =   "frmSListOfDessert.frx":5A35E
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   4800
   End
   Begin VB.Image imgCornDog 
      Height          =   3840
      Left            =   13200
      Picture         =   "frmSListOfDessert.frx":7B976
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   4620
   End
   Begin VB.Image imgChocolate 
      Height          =   5370
      Left            =   7440
      Picture         =   "frmSListOfDessert.frx":9C37E
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3705
   End
   Begin VB.Image imgCake 
      Height          =   3735
      Left            =   840
      Picture         =   "frmSListOfDessert.frx":BF652
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   4920
   End
   Begin VB.Label lblListOfDesser 
      Caption         =   "LIST OF DESSERT :"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmSListOfDessert"
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

