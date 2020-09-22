VERSION 5.00
Begin VB.Form frmOnline 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Online Browser"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4335
   FillColor       =   &H80000012&
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
   ScaleHeight     =   1005
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   600
      Top             =   360
   End
   Begin VB.CommandButton cmdShow 
      Height          =   615
      Left            =   0
      Picture         =   "frmOnline.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Slide Show"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtURL 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton cmdFwdAll 
      Caption         =   ">>>"
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
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdFwd1 
      Caption         =   ">"
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
      Left            =   2160
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack1 
      Caption         =   "<"
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
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdBackAll 
      Caption         =   "<<<"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdBack1_Click()
On Error GoTo Error
If frmMain.lstRemote.ListCount <> 0 Then

If frmMain.lstRemote.ListIndex = 0 Then
frmMain.lstRemote.ListIndex = frmMain.lstRemote.ListCount - 1
HyperJump frmMain.lstRemote.Text
txtURL.Text = frmMain.lstRemote.Text
txtURL.SelStart = Len(txtURL.Text)
Exit Sub
End If

frmMain.lstRemote.ListIndex = frmMain.lstRemote.ListIndex - 1
frmMain.lstRemote.Selected(frmMain.lstRemote.ListIndex) = True

HyperJump frmMain.lstRemote.Text
txtURL.Text = frmMain.lstRemote.Text
txtURL.SelStart = Len(txtURL.Text)

End If

Exit Sub
Error:
frmMain.lstRemote.ListIndex = frmMain.lstRemote.ListCount - 1
frmMain.lstRemote.Selected(frmMain.lstRemote.ListIndex) = True
If frmMain.lstRemote.Text <> "" Then
HyperJump frmMain.lstRemote.Text
txtURL.Text = frmMain.lstRemote.Text
txtURL.SelStart = Len(txtURL.Text)

End If

End Sub

Private Sub cmdBackAll_Click()
If frmMain.lstRemote.ListCount <> 0 Then
frmMain.lstRemote.ListIndex = 0
frmMain.lstRemote.Selected(frmMain.lstRemote.ListIndex) = True
HyperJump frmMain.lstRemote.Text
txtURL.Text = frmMain.lstRemote.Text
txtURL.SelStart = Len(txtURL.Text)
End If
End Sub

Private Sub cmdFwd1_Click()

If frmMain.lstRemote.ListCount <> 0 Then

If frmMain.lstRemote.ListIndex = frmMain.lstRemote.ListCount - 1 Then
frmMain.lstRemote.ListIndex = 0
frmMain.lstRemote.Selected(frmMain.lstRemote.ListIndex) = True

HyperJump frmMain.lstRemote.Text
txtURL.Text = frmMain.lstRemote.Text
txtURL.SelStart = Len(txtURL.Text)
Exit Sub
End If

frmMain.lstRemote.ListIndex = frmMain.lstRemote.ListIndex + 1
frmMain.lstRemote.Selected(frmMain.lstRemote.ListIndex) = True

If frmMain.lstRemote.Text <> "" Then
HyperJump frmMain.lstRemote.Text
txtURL.Text = frmMain.lstRemote.Text
txtURL.SelStart = Len(txtURL.Text)

End If
End If

End Sub

Private Sub cmdFwdAll_Click()

If frmMain.lstRemote.ListCount <> 0 Then
frmMain.lstRemote.ListIndex = frmMain.lstRemote.ListCount - 1
frmMain.lstRemote.Selected(frmMain.lstRemote.ListIndex) = True
HyperJump frmMain.lstRemote.Text
txtURL.Text = frmMain.lstRemote.Text
txtURL.SelStart = Len(txtURL.Text)
End If

End Sub

Private Sub cmdShow_Click()
If tmr = False Then
frmShowDefsO.Show 1
Else
tmr = False
End If

End Sub


Private Sub Form_Load()
FormOnTop Me
txtURL.Text = frmMain.lstRemote.Text
txtURL.SelStart = Len(txtURL.Text)
End Sub
Public Sub HyperJump(ByVal URL As String)

    'Function to execute the Hyperlink
    Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Form_Unload(Cancel As Integer)
tmr = False

frmMain.WindowState = 0
End Sub

Private Sub tmr_Timer()

If frmShowDefsO.cmdDirection.Caption = "Forward" Then
Call cmdFwd1_Click
Else
Call cmdBack1_Click
End If

End Sub
