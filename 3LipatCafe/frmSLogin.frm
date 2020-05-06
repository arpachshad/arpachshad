VERSION 5.00
Begin VB.Form frmSLogin 
   Caption         =   "LOG IN"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
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
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLEAR"
      Height          =   735
      Left            =   7800
      TabIndex        =   5
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "ENTER"
      Height          =   735
      Left            =   5520
      TabIndex        =   4
      Top             =   7200
      Width           =   1695
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   735
      Left            =   5040
      TabIndex        =   3
      Top             =   5880
      Width           =   5175
   End
   Begin VB.TextBox txtUsename 
      Alignment       =   2  'Center
      Height          =   735
      Left            =   5040
      TabIndex        =   2
      Top             =   4800
      Width           =   5175
   End
   Begin VB.Label lblPassword 
      Caption         =   "STAFF PASSWORD : "
      Height          =   735
      Left            =   1800
      TabIndex        =   1
      Top             =   6000
      Width           =   2775
   End
   Begin VB.Label lblUsename 
      Caption         =   "STAFF USERNAME :"
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   4920
      Width           =   2775
   End
End
Attribute VB_Name = "frmSLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
txtUsename.Clear
txtPassword.Clear

End Sub

Private Sub cmdEnter_Click()

If (txtPassword.Text = "12345") Then

frmSLogIn.Hide
frmSMenu.Show

Else
    MsgBox "Wrong Password!"
End If

End Sub

Private Sub cmdNext_Click()

End Sub

Private Sub cmdExit_Click()
End
End Sub

