VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSTotalSalesDrinks 
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
      Left            =   600
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
   Begin VB.Frame frmTotalSalesDrinks 
      Height          =   2895
      Left            =   14280
      TabIndex        =   26
      Top             =   5760
      Width           =   5415
      Begin VB.Label lblTotalSalesDrinks 
         Alignment       =   2  'Center
         Caption         =   "TOTAL SALES DRINKS :"
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
         TabIndex        =   28
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label lblDisplayTotalSalesDrinks 
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
         Top             =   1320
         Width           =   4455
      End
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   8640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblDisplayProfitStrawberry 
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
      Left            =   5040
      TabIndex        =   25
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblDisplayProfitOreo 
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
      Left            =   4920
      TabIndex        =   24
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label lblDisplayProfitWater 
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
      Left            =   11880
      TabIndex        =   23
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblDisplayProfitOrange 
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
      TabIndex        =   22
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label lblDisplayProfitWatermelon 
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
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblDisplaySalesStrawberry 
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
      Left            =   5040
      TabIndex        =   20
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblDisplaySalesOreo 
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
      Left            =   4920
      TabIndex        =   19
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label lblDisplaySalesWater 
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
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblDisplaySalesOrange 
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
      Left            =   11880
      TabIndex        =   17
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label lblDisplaySalesWatermelon 
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
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblTotalStrawberry 
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
      Left            =   2880
      TabIndex        =   15
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblTotalOreo 
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
      Left            =   2760
      TabIndex        =   14
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label lblTotalWater 
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
      Left            =   9720
      TabIndex        =   13
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblTotalOrange 
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
      TabIndex        =   12
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label lblTotalWatermelon 
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
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblNumSalesStrawberry 
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
      Left            =   3000
      TabIndex        =   10
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblNumSalesOreo 
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
      Left            =   2880
      TabIndex        =   9
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label lblNumSalesWater 
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
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblNumSalesOrange 
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
      Left            =   9840
      TabIndex        =   7
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label lblNumSalesWatermelon 
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
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblOreo 
      Alignment       =   2  'Center
      Caption         =   "OREO MILK SHAKE - RM 15.00"
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
      Left            =   2640
      TabIndex        =   5
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label lblSkyJuice 
      Alignment       =   2  'Center
      Caption         =   "SKY JUICE - RM 3.50"
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
      Left            =   9600
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblOrange 
      Alignment       =   2  'Center
      Caption         =   "ORANGE JUICE - RM 12.00"
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
      Width           =   2175
   End
   Begin VB.Label lblWatermelon 
      Alignment       =   2  'Center
      Caption         =   "WATERMELON JUICE - RM 12.00"
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
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label lblDetox 
      Alignment       =   2  'Center
      Caption         =   "STRAWBERRY FIZZY SPLASH - RM 18.90"
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
      Left            =   2640
      TabIndex        =   1
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Image imgOrange 
      Height          =   3375
      Left            =   7200
      Picture         =   "frmSTotalSalesDrink.frx":0000
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   2640
   End
   Begin VB.Image Image3 
      Height          =   3855
      Left            =   13560
      Picture         =   "frmSTotalSalesDrink.frx":32717
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2760
   End
   Begin VB.Image imgWater 
      Height          =   2895
      Left            =   7080
      Picture         =   "frmSTotalSalesDrink.frx":61142
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2400
   End
   Begin VB.Image imgOreo 
      Height          =   2535
      Left            =   360
      Picture         =   "frmSTotalSalesDrink.frx":83A47
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   2400
   End
   Begin VB.Image imgStrawberry 
      Height          =   2535
      Left            =   240
      Picture         =   "frmSTotalSalesDrink.frx":C478C
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2640
   End
   Begin VB.Label lblSaleDrinks 
      Caption         =   "TOTAL SALES : DRINKS"
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
      Width           =   6615
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
Attribute VB_Name = "frmSTotalSalesDrinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNext_Click()
Me.Hide
frmSGrandTotal.Show
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub mnuEditColor_Click()
dlgCommon.ShowColor
frmSTotalSalesDrinks.ForeColor = dlgCommon.Color


lblDisplaySalesStrawberry.ForeColor = dlgCommon.Color
lblDisplaySalesWater.ForeColor = dlgCommon.Color
lblDisplaySalesWatermelon.ForeColor = dlgCommon.Color
lblDisplaySalesOrange.ForeColor = dlgCommon.Color
lblDisplaySalesOreo.ForeColor = dlgCommon.Color
lblDisplayTotalSalesDrinks.ForeColor = dlgCommon.Color

End Sub

Private Sub mnuEditFont_Click()
dlgCommon.ShowFont
lblDisplaySalesStrawberry.Font.Name = dlgCommon.FontName
lblDisplaySalesWater.Font.Name = dlgCommon.FontName
lblDisplaySalesWatermelon.Font.Name = dlgCommon.FontName
lblDisplaySalesOrange.Font.Name = dlgCommon.FontName
lblDisplaySalesOreo.Font.Name = dlgCommon.FontName
lblDisplayTotalSalesDrinks.Font.Name = dlgCommon.FontName

lblDisplaySalesStrawberry.Font.Size = dlgCommon.FontSize
lblDisplaySalesWater.Font.Size = dlgCommon.FontSize
lblDisplaySalesWatermelon.Font.Size = dlgCommon.FontSize
lblDisplaySalesOrange.Font.Size = dlgCommon.FontSize
lblDisplaySalesOreo.Font.Size = dlgCommon.FontSize
lblDisplayTotalSalesDrinks.Font.Size = dlgCommon.FontSize

End Sub

Private Sub mnuFilePrint_Click()
dlgCommon.ShowPrinter
End Sub

Private Sub mnuFileSave_Click()
dlgCommon.ShowSave
End Sub
