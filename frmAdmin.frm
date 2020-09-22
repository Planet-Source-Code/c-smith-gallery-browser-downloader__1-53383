VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmAdmin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Admin Update"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4635
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
   ScaleHeight     =   4200
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRawInfo 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   4455
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://casey:pr0nbox@"
      UserName        =   "casey"
      Password        =   "pr0nbox"
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Update info now"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
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
      Left            =   3600
      TabIndex        =   2
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtURL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "1.1"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtUpdate 
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Latest version"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Idle"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   3360
      Width           =   2775
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdUpload_Click()
'Update the text file via FTP and inet

On Error GoTo Error

txtRawInfo.Text = ""

If Inet1.StillExecuting Then
Inet1.Cancel
End If


txtRawInfo.Text = txtRawInfo.Text & "$%$updatemessage$%$" & txtUpdate.Text & "$%$updatemessage$%$" & _
vbNewLine & vbNewLine & "$%$ver$%$" & txtURL.Text & "$%$ver$%$" & vbNewLine & vbNewLine & "Last Updated - " & Date
lblStatus.Caption = "Writing text file..."

Open App.Path & "\Update2.txt" For Output As #1: Print #1, txtRawInfo.Text: Close #1
lblStatus.Caption = "Update file written"

Do While Inet1.StillExecuting
DoEvents
Loop

Dim inetURL As String
Dim UserName As String
Dim Password As String


Inet1.URL = inetURL
Inet1.UserName = UserName
Inet1.Password = Password

Inet1.Execute , "DIR"


Do While Inet1.StillExecuting
DoEvents
Loop

'You may not need this, or you may need more
'This changes directories on the ftp server

Inet1.Execute , "CD /web"

Do While Inet1.StillExecuting
DoEvents
Loop

SendToFTP


Exit Sub
Error:
MsgBox Err.Number & Err.Description, vbCritical, "Error"
End Sub

Private Sub Command1_Click()
Me.Hide

End Sub


Private Sub Form_Load()
FormOnTop Me

Me.ZOrder (0)

End Sub


Private Sub Inet1_StateChanged(ByVal State As Integer)
Select Case State
    Case 1
        lblStatus.Caption = "Trying to resolve host..."
    Case 2
        lblStatus.Caption = "Host is resolved"
    Case 3
        lblStatus.Caption = "Sending connection request..."
    Case 4
        lblStatus.Caption = "Connected"
    Case 5
        lblStatus.Caption = "Sending request..."
    Case 6
        lblStatus.Caption = "Request sent"
    Case 7
        lblStatus.Caption = "Receiving response..."
    Case 8
        lblStatus.Caption = "Response received"
    Case 9
        lblStatus.Caption = "Disconnecting..."
    Case 10
        lblStatus.Caption = "Disconnected"
    Case 11
        lblStatus.Caption = "ERROR"
        MsgBox Inet1.ResponseCode & ":" & Inet1.ResponseInfo
        
    Case 12
        'finished
        lblStatus.Caption = "Completed Upload"
    End Select
End Sub

Public Sub SendToFTP()
On Error GoTo Error

If Inet1.StillExecuting = False Then

Inet1.Execute , "PUT " & """" & App.Path & "\Update2.txt" & """" & " " & "Update2.txt"     ' issues the PUT command to ftp

lblStatus.Caption = "Uploading file..."


End If

Exit Sub
Error:
MsgBox "Error:" & Err.Description
    Select Case Err.Number
    Case 35764
    DoEvents
    Resume
    End Select
End Sub


Private Sub txtUpdate_Change()
SaveSetting App.Title, "Settings", "Update message", txtUpdate.Text

End Sub

Private Sub txtURL_Change()
SaveSetting App.Title, "Settings", "Update URL", txtURL.Text

End Sub
Public Sub InitFTP()

On Error GoTo Error:
    Inet1.Execute , "DIR"                ' issues the DIR command to ftp
        
     
Exit Sub

Error:

    Select Case Err.Number
    Case 35764
    DoEvents
    Resume
    End Select
End Sub
