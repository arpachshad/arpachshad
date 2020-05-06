VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00FFFFFF&
   Caption         =   "WELCOME"
   ClientHeight    =   10485
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   19230
   LinkTopic       =   "Form1"
   Picture         =   "frmWelcome.frx":0000
   ScaleHeight     =   1928.276
   ScaleMode       =   0  'User
   ScaleWidth      =   1089.347
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClickStaff 
      BackColor       =   &H00FF8080&
      Caption         =   "Continue as staff"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   11880
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9480
      Width           =   5895
   End
   Begin VB.CommandButton cmdClickCustomer 
      BackColor       =   &H00FF80FF&
      Caption         =   "Continue as customer"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4920
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9480
      Width           =   5775
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClickCustomer_Click()

'To hide the current form, frmWelcome
frmWelcome.Hide

'To Show or Navigate user to frmGuest
frmGDetail.Show


End Sub

Private Sub cmdClickStaff_Click()

'To hide the current form, frmWelcome
frmWelcome.Hide

'To Show or Navigate user to frmStaff
frmSLogIn.Show

End Sub
