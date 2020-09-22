VERSION 5.00
Begin VB.Form frmOvoyOpts2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOvoyOpts2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox cboRemote 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Text            =   "Download it"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Double clicking a remote picture will"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmOvoyOpts2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboRemote_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub Command1_Click()
Me.Hide

End Sub

Private Sub Form_Load()
cboRemote.AddItem "Open it"
cboRemote.AddItem "Download it"
End Sub
