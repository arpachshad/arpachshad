VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSTotalSalesDessert 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      TabIndex        =   30
      Top             =   9720
      Width           =   1335
   End
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
      TabIndex        =   29
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame frmTotalSalesDessert 
      Height          =   2895
      Left            =   14400
      TabIndex        =   26
      Top             =   5640
      Width           =   5415
      Begin VB.Label lblTotalSalesDessert 
         Alignment       =   2  'Center
         Caption         =   "TOTAL SALES DESSERT :"
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
         Left            =   360
         TabIndex        =   28
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label lblDisplayTotalSalesDessert 
         Alignment       =   2  'Center
         Caption         =   "RM 15,000"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   480
         TabIndex        =   27
         Top             =   1080
         Width           =   4455
      End
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   8520
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblDisplayProfitCake 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   25
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblDisplayProfitPisang 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   24
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Label lblDisplayProfitChocolate 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12000
      TabIndex        =   23
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblDisplayProfitWaffle 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12120
      TabIndex        =   22
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label lblDisplayProfitCornDog 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18600
      TabIndex        =   21
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblDisplaySalesCake 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   20
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblDisplaySalesPisang 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   19
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label lblDisplaySalesChocolate 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11760
      TabIndex        =   18
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblDisplaySalesWaffle 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12000
      TabIndex        =   17
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label lblDisplaySalesCornDog 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18480
      TabIndex        =   16
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblTotalCake 
      Alignment       =   2  'Center
      Caption         =   "TOTAL PROFIT :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   15
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label lblTotalPisang 
      Alignment       =   2  'Center
      Caption         =   "TOTAL PROFIT :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   14
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label lblTotalChocolate 
      Alignment       =   2  'Center
      Caption         =   "TOTAL PROFIT :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9840
      TabIndex        =   13
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label lblTotalWaffle 
      Alignment       =   2  'Center
      Caption         =   "TOTAL PROFIT :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   12
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label lblTotalCornDog 
      Alignment       =   2  'Center
      Caption         =   "TOTAL PROFIT :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16440
      TabIndex        =   11
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label lblNumSalesCake 
      Alignment       =   2  'Center
      Caption         =   "NUM.SALES:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   10
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblNumSalesPisang 
      Alignment       =   2  'Center
      Caption         =   "NUM.SALES:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   9
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label lblNumSalesChocolate 
      Alignment       =   2  'Center
      Caption         =   "NUM.SALES:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblNumSalesWaffle 
      Alignment       =   2  'Center
      Caption         =   "NUM.SALES:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   7
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label lblNumSalesCornDog 
      Alignment       =   2  'Center
      Caption         =   "NUM.SALES:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16440
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblCake 
      Alignment       =   2  'Center
      Caption         =   "CAKE MACPAS ( 1 slice ) - RM 15.50"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   5
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label lblPisang 
      Alignment       =   2  'Center
      Caption         =   "PISANG GORENG CHEESE ( 5pcs) - RM 15.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   6240
      Width           =   3255
   End
   Begin VB.Label lblWaffle 
      Alignment       =   2  'Center
      Caption         =   "WAFFLE DOUBLE LAYER with ICE CREAM - RM 20.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9720
      TabIndex        =   3
      Top             =   6240
      Width           =   4215
   End
   Begin VB.Label lblChocolate 
      Alignment       =   2  'Center
      Caption         =   "CHOCOLATE ECLAIRS - RM 20.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      TabIndex        =   2
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label lblCornDog 
      Alignment       =   2  'Center
      Caption         =   "CORN DOG ( 3 pcs) - RM 25.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   16320
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Image imgPisang 
      Height          =   2655
      Left            =   240
      Picture         =   "frmSTotalSalesDessert.frx":0000
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   3480
   End
   Begin VB.Image imgChocolate 
      Height          =   3975
      Left            =   6960
      Picture         =   "frmSTotalSalesDessert.frx":21618
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2760
   End
   Begin VB.Image imgWaffle 
      Height          =   3135
      Left            =   6960
      Picture         =   "frmSTotalSalesDessert.frx":448EC
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   2880
   End
   Begin VB.Image imgCornDog 
      Height          =   2655
      Left            =   13440
      Picture         =   "frmSTotalSalesDessert.frx":9EC4A
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3120
   End
   Begin VB.Image imgStrawberry 
      Height          =   2655
      Left            =   240
      Picture         =   "frmSTotalSalesDessert.frx":BF652
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   3480
   End
   Begin VB.Label lblSaleDessert 
      Caption         =   "TOTAL SALES : DESSERT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditColor 
         Caption         =   "Text Color"
      End
      Begin VB.Menu mnuEditFont 
         Caption         =   "Font"
      End
   End
End
Attribute VB_Name = "frmsTotalSalesDessert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdNext_Click()
Me.Hide
frmSGrandTotal.Show
End Sub

Private Sub mnuEditColor_Click()
dlgCommon.ShowColor

frmsTotalSalesDessert.ForeColor = dlgCommon.Color

lblDisplaySalesCake.ForeColor = dlgCommon.Color
lblDisplaySalesPisang.ForeColor = dlgCommon.Color
lblDisplaySalesChocalate.ForeColor = dlgCommon.Color
lblDisplaySalesWaffle.ForeColor = dlgCommon.Color
lblDisplaySalesCornDog.ForeColor = dlgCommon.Color
lblDisplayTotalSalesDessert.ForeColor = dlgCommon.Color

End Sub

Private Sub mnuEditFont_Click()
dlgCommon.ShowFont
lblDisplaySalesCake.Font.Name = dlgCommon.FontName
lblDisplaySalesPisang.Font.Name = dlgCommon.FontName
lblDisplaySalesChocolate.Font.Name = dlgCommon.FontName
lblDisplaySalesWaffle.Font.Name = dlgCommon.FontName
lblDisplaySalesCornDog.Font.Name = dlgCommon.FontName
lblDisplayTotalSalesDessert.Font.Name = dlgCommon.FontName

lblDisplaySalesCake.Font.Size = dlgCommon.FontSize
lblDisplaySalesPisang.Font.Size = dlgCommon.FontSize
lblDisplaySalesChocolate.Font.Size = dlgCommon.FontSize
lblDisplaySalesWaffle.Font.Size = dlgCommon.FontSize
lblDisplaySalesCornDog.Font.Size = dlgCommon.FontSize
lblDisplayTotalSalesDessert.Font.Size = dlgCommon.FontSize
End Sub

Private Sub mnuFilePrint_Click()
dlgCommon.ShowPrinter
End Sub

Private Sub mnuFileSave_Click()
dlgCommon.ShowSave
End Sub
