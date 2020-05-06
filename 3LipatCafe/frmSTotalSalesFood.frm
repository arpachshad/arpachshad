VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSTotalSalesFood 
   Caption         =   "Form2"
   ClientHeight    =   10635
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   20250
   LinkTopic       =   "Form2"
   ScaleHeight     =   10635
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   6720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   30
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame frmTotalSalesFood 
      Height          =   2895
      Left            =   14640
      TabIndex        =   27
      Top             =   5640
      Width           =   5415
      Begin VB.Label lblDisplayTotalSalesFood 
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
         TabIndex        =   29
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label lblTotalSalesFood 
         Alignment       =   2  'Center
         Caption         =   "TOTAL SALES FOOD :"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Label lblDisplayProfitNasiLemak 
      Caption         =   "RM 15,000"
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
      Left            =   18480
      TabIndex        =   25
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblDisplayProfitBurger 
      Caption         =   "RM 15,000"
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
      Left            =   4800
      TabIndex        =   24
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblDisplaySalesSpagetti 
      Caption         =   "200 unit"
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
      Left            =   11400
      TabIndex        =   23
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblDisplaySalesChickenGrill 
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
      Left            =   4800
      TabIndex        =   22
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label lblDisplaySalesMeeGoreng 
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
      Left            =   11520
      TabIndex        =   21
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label lblDisplaySalesNasiLemak 
      Caption         =   "20000 unit"
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
      TabIndex        =   20
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblDisplayProfitMeeGoreng 
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
      Left            =   11640
      TabIndex        =   19
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Label lblDisplayProfitChickenGrill 
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
      Left            =   4800
      TabIndex        =   18
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label lblDisplayProfitSpagetti 
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
      Left            =   11520
      TabIndex        =   17
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblDisplaySalesBurger 
      Caption         =   "1500 unit"
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
      Left            =   4800
      TabIndex        =   16
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lblTotalMeeGoreng 
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
      Left            =   9480
      TabIndex        =   15
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label lblTotalSpagetti 
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
      Left            =   9360
      TabIndex        =   14
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblTotalNasiLemak 
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
      Left            =   16320
      TabIndex        =   13
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lblTotalChicken 
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
      Left            =   2640
      TabIndex        =   12
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label lblNasiLemak 
      Alignment       =   2  'Center
      Caption         =   "NASI LEMAK BERGANDA  -  RM 18.00"
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
      TabIndex        =   11
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label lblMee 
      Alignment       =   2  'Center
      Caption         =   "MEE GORENG MONSTER  -  RM 15.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9360
      TabIndex        =   10
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Label lblSpagetti 
      Alignment       =   2  'Center
      Caption         =   "SPAGETTI BOLOGNESE POSSO FARLE  -  RM 20.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9240
      TabIndex        =   9
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label lblNumSalesNasiLemak 
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
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblNumSalesMeeGoreng 
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
      Left            =   9480
      TabIndex        =   7
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label lblNumSalesSpagetti 
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
      Left            =   9360
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblNumSalesChicken 
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
      Left            =   2760
      TabIndex        =   5
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label lblChickenGrill 
      Alignment       =   2  'Center
      Caption         =   "CHICKEN GRILED DELICIOUS  - RM 32.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   4
      Top             =   5880
      Width           =   3735
   End
   Begin VB.Image Image3 
      Height          =   2640
      Left            =   13080
      Picture         =   "frmSTotalSalesFood.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   3240
   End
   Begin VB.Image Image2 
      Height          =   2415
      Left            =   6840
      Picture         =   "frmSTotalSalesFood.frx":59D05
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   2580
   End
   Begin VB.Image Image1 
      Height          =   2520
      Left            =   6480
      Picture         =   "frmSTotalSalesFood.frx":92A77
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2760
   End
   Begin VB.Image imgChickenChop 
      Height          =   2415
      Left            =   -120
      Picture         =   "frmSTotalSalesFood.frx":31C5C7
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   2970
   End
   Begin VB.Label lblTotalBurger 
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
      Left            =   2640
      TabIndex        =   3
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblNumSalesBurger 
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
      Left            =   2760
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblBurger 
      Alignment       =   2  'Center
      Caption         =   "BURGER TRIPLE TOWER  -  RM 23.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Image imgBurger 
      Height          =   2520
      Left            =   -240
      Picture         =   "frmSTotalSalesFood.frx":34E0A5
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   3120
   End
   Begin VB.Label lblSafeFood 
      Caption         =   "TOTAL SALES : FOOD"
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
      Width           =   4815
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
      Begin VB.Menu mnuFont 
         Caption         =   "Font"
      End
   End
End
Attribute VB_Name = "frmSTotalSalesFood"
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

frmSTotalSalesFood.ForeColor = dlgCommon.Color

lblDisplaySalesBurger.ForeColor = dlgCommon.Color
lblDisplaySalesSpagetti.ForeColor = dlgCommon.Color
lblDisplaySalesMeeGoreng.ForeColor = dlgCommon.Color
lblDisplaySalesNasiLemak.ForeColor = dlgCommon.Color
lblDisplaySalesChickenGrill.ForeColor = dlgCommon.Color
lblDisplayTotalSalesFood.ForeColor = dlgCommon.Color


End Sub

Private Sub mnuFilePrint_Click()
dlgCommon.ShowPrinter
End Sub

Private Sub mnuFileSave_Click()
dlgCommon.ShowSave
End Sub

Private Sub mnuFont_Click()
dlgCommon.ShowFont
lblDisplaySalesBurger.Font.Name = dlgCommon.FontName
lblDisplaySalesSpagetti.Font.Name = dlgCommon.FontName
lblDisplaySalesMeeGoreng.Font.Name = dlgCommon.FontName
lblDisplaySalesNasiLemak.Font.Name = dlgCommon.FontName
lblDisplaySalesChickenGrill.Font.Name = dlgCommon.FontName
lblDisplayTotalSalesFood.Font.Name = dlgCommon.FontName

lblDisplaySalesBurger.Font.Size = dlgCommon.FontSize
lblDisplaySalesSpagetti.Font.Size = dlgCommon.FontSize
lblDisplaySalesMeeGoreng.Font.Size = dlgCommon.FontSize
lblDisplaySalesNasiLemak.Font.Size = dlgCommon.FontSize
lblDisplaySalesChickenGrill.Font.Size = dlgCommon.FontSize
lblDisplayTotalSalesFood.Font.Size = dlgCommon.FontSize


End Sub
