VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents MsgrClient As clsMsgr
Attribute MsgrClient.VB_VarHelpID = -1

Private Sub Command1_Click()
Set MsgrClient = New clsMsgr
MsgrClient.Login "Username@hotmail.com", "YourPassword"
End Sub

Private Sub Form_Load()
Set MsgrClient = New clsMsgr
MsgrClient.Login "Usernaame@hotmail.com", "YourPassword"
End Sub
