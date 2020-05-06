VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSUpdate 
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   15045
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc item 
      Height          =   615
      Left            =   14040
      Top             =   3600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
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
      RecordSource    =   "item"
      Caption         =   "Select item"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Frame frmUpdate 
      Height          =   5175
      Left            =   4920
      TabIndex        =   0
      Top             =   3000
      Width           =   11415
      Begin VB.CommandButton cmdSubmit 
         BackColor       =   &H0080C0FF&
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox txtNewName 
         Alignment       =   2  'Center
         DataField       =   "item_name"
         DataSource      =   "item"
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
         Left            =   2880
         TabIndex        =   5
         Top             =   3000
         Width           =   6015
      End
      Begin VB.TextBox txtNewPrice 
         Alignment       =   2  'Center
         DataField       =   "item_price"
         DataSource      =   "item"
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
         Left            =   2880
         TabIndex        =   4
         Top             =   1920
         Width           =   6015
      End
      Begin VB.Label lblListMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "item_name"
         DataSource      =   "item"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2880
         TabIndex        =   9
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label lblNewName 
         Caption         =   "New Name :"
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
         Left            =   960
         TabIndex        =   3
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label lblNewPrice 
         Caption         =   "New Price (RM) :"
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
         Left            =   240
         TabIndex        =   2
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label lblMenu 
         Alignment       =   1  'Right Justify
         Caption         =   "Menu :"
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
         Left            =   1080
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Label lblUpdate 
      Alignment       =   2  'Center
      Caption         =   "UPDATE PAGE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7080
      TabIndex        =   7
      Top             =   1440
      Width           =   6735
   End
End
Attribute VB_Name = "frmSUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub cmdNext_Click()
Me.Hide
frmSMenu.Show
End Sub

Private Sub cmdSubmit_Click()
item.Recordset.Update
MsgBox "Data has been updated!"
frmSUpdate.Hide
frmSMenu.Show


End Sub

Private Sub Form_Load()

'To read and open the shopping.txt to read the data inside notepad
'Open "C:\Users\habib\Desktop\TestProject NewVb\menu.txt" For Input As #1

'To declare the variable
'Dim strMenu As String


'Do Until EOF(1)
'To retrieve from the text file
'Input #1, strMenu

'To store the data into the Combo Box
'cboMenu.AddItem strMenu

'Loop

'Close #1

'To add the item in the combo box during the run time

End Sub

