VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmUpdate2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "oVoy2 Updates"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5040
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
   ScaleHeight     =   4110
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRawInfo 
      Height          =   2295
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   4455
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   720
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
   Begin VB.CommandButton cmdForce 
      Caption         =   "Force Update"
      Height          =   855
      Left            =   120
      Picture         =   "frmUpdate2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
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
      Left            =   3960
      TabIndex        =   2
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtUpdateInfo 
      BackColor       =   &H8000000F&
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmUpdate2.frx":08CA
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label lblCurVer 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFC0C0&
      X1              =   1440
      X2              =   3840
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF8080&
      X1              =   1440
      X2              =   3840
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      X1              =   1440
      X2              =   3840
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   1440
      X2              =   3840
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label lblVer 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   3120
      Width           =   615
   End
End
Attribute VB_Name = "frmUpdate2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdForce_Click()

On Error GoTo Error
SaveSetting "MSVUD", "UD", "Key", "dfjkhijhg#rt%"
Shell App.Path & "\updater.exe", vbNormalFocus
End
Exit Sub
Error:
MsgBox "Couldn't find updater.exe. Make sure it's in the same folder as oVoy.", vbCritical, "Error"
SaveSetting "MSVUD", "UD", "Key", "ghgh55$%ijhg#rt%"
End Sub

Private Sub cmdOk_Click()
Me.Hide

End Sub

Private Sub Form_Load()
lblVer.Caption = "1.2"
Exit Sub

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
Select Case State
    Case 1
        lblStat.Caption = "Trying to resolve host..."
    Case 2
        lblStat.Caption = "Host is resolved"
    Case 3
        lblStat.Caption = "Sending connection request..."
    Case 4
        lblStat.Caption = "Connected"
    Case 5
        lblStat.Caption = "Sending request..."
    Case 6
        lblStat.Caption = "Request sent"
    Case 7
        lblStat.Caption = "Receiving response..."
    Case 8
        lblStat.Caption = "Response received"
    Case 9
        lblStat.Caption = "Disconnecting..."
    Case 10
        lblStat.Caption = "Disconnected"
    Case 11
        lblStat.Caption = "Error"
        Inet1.Cancel
    End Select
End Sub
