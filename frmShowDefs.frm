VERSION 5.00
Begin VB.Form frmShowDefs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Slide Show Settings"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3915
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdDirection 
      Caption         =   "Forward"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtDelay 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "2"
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3360
      Picture         =   "frmShowDefs.frx":0000
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "To stop the slide show, click the same button"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Direction"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Delay in seconds"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmShowDefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
Me.Hide

End Sub

Private Sub cmdDirection_Click()
If cmdDirection.Caption = "Forward" Then
cmdDirection.Caption = "Backward"
Else
cmdDirection.Caption = "Forward"
End If
txtDelay.SetFocus
SaveSetting App.Title, "Settings", "Direction", cmdDirection.Caption

End Sub

Private Sub cmdOk_Click()
frmGallery.tmr = True

Me.Hide

End Sub

Private Sub Form_Load()
FormOnTop Me
frmGallery.tmr.Interval = txtDelay * 1000

End Sub

Private Sub txtDelay_Change()
If txtDelay.Text = "" Then
txtDelay.Text = "1"
End If
If txtDelay > 60 Then
txtDelay.Text = 60
End If
frmGallery.tmr.Interval = txtDelay * 1000
SaveSetting App.Title, "Settings", "Delay", txtDelay.Text
End Sub

Private Sub txtDelay_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
KeyAscii = 8
GoTo MoveOn
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
Exit Sub
End If

MoveOn:
If txtDelay > 60 Then
txtDelay.Text = 60
End If

frmGallery.tmr.Interval = txtDelay * 1000
End Sub
