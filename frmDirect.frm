VERSION 5.00
Begin VB.Form frmDirect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Choose a folder"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3675
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
   ScaleHeight     =   5385
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNewDir 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   4
      Text            =   "New name"
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdNewDir 
      Caption         =   "Create new folder>"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3495
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
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
      Left            =   2640
      TabIndex        =   0
      Top             =   4920
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmDirect.frx":0000
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   480
   End
End
Attribute VB_Name = "frmDirect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNewDir_Click()
'Create a new folder

On Error GoTo Error
MkDir Dir1.Path & "\" & txtNewDir.Text
Dir1.Refresh
Exit Sub
Error:
MsgBox "There was a problem creating the new directory." & vbNewLine & "Make sure it doesn't already exist.", vbCritical, "Path Error"

End Sub

Private Sub Command1_Click()


If frmMain.lstFolder.ListCount = 0 Then
MsgBox "No picture files found in this folder!", vbInformation, "oVoy"
frmMain.txtStatus.Text = frmMain.txtStatus.Text & vbNewLine & "Couldn't find any pictures in specified folder"
Else
frmMain.txtStatus.Text = frmMain.txtStatus.Text & vbNewLine & "Found " & frmMain.lstFolder.ListCount & " picture files in " & frmMain.cboFolder.Text
End If
If Right(frmMain.cboFolder.Text, 1) <> "\" Then
frmMain.cboFolder.Text = frmMain.cboFolder.Text & "\"
End If
frmMain.cboFolder.AddItem frmMain.cboFolder.Text

frmMain.lstDirs.AddItem frmMain.cboFolder.Text

xListKillDupes frmMain.lstDirs
SaveListBox App.Path & "\RecentDirs.oVoy", frmMain.lstDirs

If frmMain.cboFolder.Text = "C:\" Then
frmMain.lstFolder.Enabled = False
frmMain.cmdDownload.Enabled = False
frmMain.cmdDownloadAll.Enabled = False
frmMain.mnuDownloadNow.Enabled = False
Else
frmMain.lstFolder.Enabled = True
frmMain.cmdDownload.Enabled = True
frmMain.cmdDownloadAll.Enabled = True
frmMain.mnuDownloadNow.Enabled = True
End If

Me.Hide

End Sub

Private Sub Dir1_Change()



frmMain.cboFolder.Text = Dir1.Path
frmMain.FileIndex.Text = "0"
On Error Resume Next

Dim B As String, d As Integer, e As Integer, f As Integer
addTheFiles

If Right(frmMain.cboFolder.Text, 1) <> "\" Then
frmMain.cboFolder.Text = frmMain.cboFolder.Text & "\"
End If

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive

End Sub

Private Sub Form_Load()

Me.ZOrder (0)
End Sub

Sub addTheFiles()
Dim B As String
frmMain.lstFolder.Clear
B = Dir(Dir1.Path & "\*.*")
If B <> "" Then
    If Right(B, 4) = ".jpg" Then
        frmMain.lstFolder.AddItem B
    End If
    If Right(B, 4) = ".bmp" Then
        frmMain.lstFolder.AddItem B
    End If
    If Right(B, 4) = ".art" Then
        frmMain.lstFolder.AddItem B
    End If
    If Right(B, 4) = ".gif" Then
        frmMain.lstFolder.AddItem B
    End If
Else
    Exit Sub
End If
Do: DoEvents
    B = Dir
    If B <> "" Then
        If Right(B, 4) = ".jpg" Then
            frmMain.lstFolder.AddItem B
        End If
        If Right(B, 4) = ".bmp" Then
            frmMain.lstFolder.AddItem B
        End If
        If Right(B, 4) = ".art" Then
            frmMain.lstFolder.AddItem B
        End If
        If Right(B, 4) = ".gif" Then
            frmMain.lstFolder.AddItem B
        End If
    Else
        Exit Do
    End If
Loop
End Sub
