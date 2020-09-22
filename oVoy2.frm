VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "oVoy2"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "oVoy2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCommand 
      Interval        =   500
      Left            =   5640
      Top             =   2040
   End
   Begin VB.TextBox txtSave 
      Height          =   885
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "oVoy2.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "oVoy2.frx":1B4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "oVoy2.frx":1E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "oVoy2.frx":2740
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "oVoy2.frx":3592
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "oVoy2.frx":91B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "oVoy2.frx":A636
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Tlbr1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   953
      ButtonWidth     =   1984
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View"
            Key             =   "Preview"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Browse Online"
            Key             =   "BrowseOnline"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Local Gallery"
            Key             =   "LocalGallery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Updates"
            Key             =   "Updates"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmr1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5400
      Top             =   5400
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      Picture         =   "oVoy2.frx":A950
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6600
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5160
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtOnlineGal 
      Height          =   285
      Left            =   4560
      TabIndex        =   24
      Text            =   "0"
      Top             =   7440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdPopulate 
      Caption         =   "Populate List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      Picture         =   "oVoy2.frx":AC5A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtStatus 
      Height          =   975
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Text            =   "oVoy2.frx":BA9C
      Top             =   7560
      Width           =   10455
   End
   Begin VB.CommandButton cmdDownloadAll 
      Caption         =   "Download All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4560
      Picture         =   "oVoy2.frx":BAA5
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Downloads all pictures"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      Picture         =   "oVoy2.frx":C8E7
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      Picture         =   "oVoy2.frx":CBF1
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Download selected pictures"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame fraLocal 
      Caption         =   "Local system"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   6000
      TabIndex        =   11
      Top             =   2040
      Width           =   4455
      Begin VB.ListBox lstDirs 
         Height          =   1815
         Left            =   720
         TabIndex        =   35
         Top             =   1440
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ListBox lstFolder 
         Height          =   4545
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   14
         Top             =   600
         Width           =   4215
      End
      Begin VB.CommandButton cmdFolder 
         Caption         =   "..."
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cboFolder 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Text            =   "cboFolder"
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame fraRemote 
      Caption         =   "Remote"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   4455
      Begin VB.ListBox lstPresets 
         Height          =   2010
         Left            =   1080
         TabIndex        =   36
         Top             =   1560
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ListBox lstRemote 
         Height          =   4935
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame fraURL 
      Caption         =   "Address info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   10455
      Begin VB.TextBox txtCommand 
         Height          =   285
         Left            =   4560
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   960
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CommandButton cmdCreateLink 
         Caption         =   "Create Link && Copy to Clipboard"
         Height          =   615
         Left            =   7320
         Picture         =   "oVoy2.frx":CD14
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtDelay 
         Height          =   285
         Left            =   6120
         TabIndex        =   32
         Text            =   "2"
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5760
         TabIndex        =   31
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog Cmdlg1 
         Left            =   6840
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.TextBox txtPrefix 
         Height          =   285
         Left            =   4560
         MaxLength       =   50
         TabIndex        =   28
         Text            =   "Prefix - "
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtCurFile 
         Height          =   285
         Left            =   4920
         TabIndex        =   25
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FileIndex 
         Height          =   285
         Left            =   6480
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtNav 
         Height          =   285
         Left            =   4920
         TabIndex        =   20
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Caption         =   "Bounds"
         Height          =   615
         Left            =   2520
         TabIndex        =   6
         Top             =   600
         Width           =   1815
         Begin VB.TextBox txtHigh 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   9
            Text            =   "13"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtLow 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            MaxLength       =   4
            TabIndex        =   7
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "to"
            Height          =   255
            Left            =   720
            TabIndex        =   8
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.TextBox txtFormat 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "00"
         Top             =   840
         Width           =   495
      End
      Begin VB.CheckBox chkLeading 
         Caption         =   "Leading zeroes?"
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox txtAddy 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Text            =   "http://ourworld.compuserve.com/homepages/Stefan_Radke/images/bssbld/%%NUM%%.jpg"
         Top             =   240
         Width           =   9375
      End
      Begin VB.Line Line2 
         X1              =   1320
         X2              =   1080
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         X1              =   1080
         X2              =   1080
         Y1              =   960
         Y2              =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Prefix this group of pictures with:"
         Height          =   255
         Left            =   4560
         TabIndex        =   29
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Format"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Basic URL"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPre 
         Caption         =   "&Presets"
         Begin VB.Menu mnuLes 
            Caption         =   "(1 default preset)"
         End
         Begin VB.Menu lklkiyhikuy7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAddPreset 
            Caption         =   "Add"
         End
         Begin VB.Menu mnuDeletePreset 
            Caption         =   "Delete..."
            Begin VB.Menu mnureretg 
               Caption         =   "Delete Preset"
            End
            Begin VB.Menu gh5trg 
               Caption         =   "-"
            End
            Begin VB.Menu mnuDeleteAr 
               Caption         =   "Delete"
               Index           =   0
               Visible         =   0   'False
            End
         End
         Begin VB.Menu kljhki7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAddAr 
            Caption         =   "Added"
            Index           =   0
            Visible         =   0   'False
         End
      End
      Begin VB.Menu llkjhlkjh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save as..."
      End
      Begin VB.Menu rgrg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuOpts 
         Caption         =   "&Options"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGoogle 
         Caption         =   "&Quick Google Search"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuGS 
         Caption         =   "&Getting Started"
      End
      Begin VB.Menu mnuUpdates 
         Caption         =   "&Check for updates"
      End
      Begin VB.Menu lkjhlkjhgh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuRemote 
      Caption         =   "Remote"
      Visible         =   0   'False
      Begin VB.Menu mnuPreviewPic 
         Caption         =   "Preview picture"
      End
      Begin VB.Menu lkhjglkhg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy address to clipboard"
      End
      Begin VB.Menu lkhglkjhg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDownloadNow 
         Caption         =   "Download selected"
      End
   End
   Begin VB.Menu mnuLocal 
      Caption         =   "Local"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete selected"
      End
      Begin VB.Menu lkjhlkjh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewGal 
         Caption         =   "View in gallery"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Admin"
      Visible         =   0   'False
      Begin VB.Menu mnuAdminUP 
         Caption         =   "Change update message"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private m_GettingFileSize     As Boolean
Private m_DownloadingFile     As Boolean
Private m_DownloadingFileSize As Long
Private m_LocalSaveFile       As String
Private Const CONV_AMP = "#$AMP$#"
Private Const CONV_EQUALS = "#$EQUALS$#"
Private Const CONV_PLUS = "#$PLUS$#"

Private Const FIND_AMP = "\&"
Private Const FIND_EQUALS = "\="
Private Const FIND_PLUS = "\+"

Private Const REG_AMP = "&"
Private Const REG_EQUALS = "="
Private Const REG_PLUS = "+"

Private Const FILE_HEADER = "# DO NOT DELETE THIS FILE OR HEADER" & vbCrLf

Private Const PTC_PROTOCOL = "oVoy"
Private Const PTC_APPNAME = "oVoy"
Dim lngReturn As Long
Dim lngExtent As Long
Dim Extracted As String

Public Sub HyperJump(ByVal URL As String)

    'Function to execute the Hyperlink
    Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub cboFolder_Change()
On Error Resume Next




cboFolder.SelStart = Len(cboFolder)
SaveSetting App.Title, "Settings", "Directory", cboFolder.Text




End Sub

Private Sub cboFolder_Click()
'Changing local directories

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

If cboFolder.Text = "[clear list]" Then
cboFolder.Clear
lstDirs.Clear

cboFolder.AddItem "C:\"
cboFolder.AddItem "[clear list]"
cboFolder.ListIndex = 0


lstDirs.AddItem "C:\"

SaveListBox App.Path & "\RecentDirs.oVoy", lstDirs
lstFolder.Clear

Else

addTheFiles
End If


cboFolder.SelStart = Len(cboFolder)
SaveSetting App.Title, "Settings", "Directory", cboFolder.Text

End Sub

Private Sub chkLeading_Click()
If chkLeading.Value = 0 Then
txtFormat.Enabled = False
Else
txtFormat.Enabled = True
End If

SaveSetting App.Title, "Settings", "Leading", chkLeading.Value


End Sub

Private Sub cmdCreateLink_Click()
'This creates the custom protocol links

If chkLeading.Value = 1 Then
Clipboard.Clear
Clipboard.SetText "oVoy://" & txtLow.Text & "-" & txtHigh.Text & ":" & txtFormat.Text & "@" & txtAddy.Text
Else
Clipboard.Clear
Clipboard.SetText "oVoy://" & txtLow.Text & "-" & txtHigh.Text & "@" & txtAddy.Text

End If
End Sub

Private Sub cmdDelete_Click()


If lstFolder.Text <> "" Then

Dim i#
For i = 0 To lstFolder.ListCount - 1
lstFolder.ListIndex = i

If lstFolder.Selected(lstFolder.ListIndex) Then
txtStatus.Text = txtStatus.Text & vbNewLine & "Deleting " & txtCurFile.Text & "..."
DeleteFile cboFolder.Text & lstFolder.Text
txtStatus.Text = txtStatus.Text & vbNewLine & txtCurFile.Text & " Deleted"
End If

Next i

addTheFiles
End If



End Sub

Private Sub cmdDownload_Click()
'Download a single file
'To download all, the timer just
'goes down the listindex clicking
'the download button each time

On Error GoTo Error

If lstRemote.Text <> "" Then

If lstRemote.Selected(lstRemote.ListIndex) = True Then

txtCurFile.Text = Mid(lstRemote.Text, InStrRev(lstRemote.Text, "/", Len(lstRemote.Text), vbTextCompare) + 1)


Do While Inet1.StillExecuting
DoEvents
Loop


If txtNav.Text <> "" Then
Dim download() As Byte

download() = Inet1.OpenURL(lstRemote.Text, icByteArray)
    Open cboFolder.Text & txtPrefix.Text & txtCurFile.Text For Binary As #1
    Put #1, , download()
    Close #1
   txtStatus.Text = txtStatus.Text & vbNewLine & txtCurFile.Text & " - saved"

End If

addTheFiles

End If
End If

Exit Sub
Error:
tmr1 = False
lstRemote.Enabled = True
cmdPopulate.Enabled = True
cmdDownloadAll.Caption = "STOP"
txtStatus.Text = txtStatus.Text & vbNewLine & "An error occured while downloading " & txtCurFile.Text
End Sub

Private Sub cmdDownloadAll_Click()
'Download the entire picture series
'Appends the prefix(if specified)
'to each picture name

If lstRemote.ListCount <> 0 Then
lstRemote.ListIndex = 0
lstRemote.Selected(lstRemote.ListIndex) = True

If lstRemote.Text <> "" Then
If cmdDownloadAll.Caption = "Download All" Then
lstRemote.ListIndex = 0
lstRemote.Selected(lstRemote.ListIndex) = True
cmdDownloadAll.Caption = "STOP"
lstRemote.Enabled = False
cmdPopulate.Enabled = False
txtPrefix.Enabled = False
Call cmdDownload_Click
tmr1 = True
Else
cmdDownloadAll.Caption = "Download All"
tmr1 = False
txtStatus.Text = txtStatus.Text & vbNewLine & "Cancelled operation"
Inet1.Cancel
lstRemote.Enabled = True
cmdPopulate.Enabled = True
txtPrefix.Enabled = True
tmr1 = False
End If
End If
End If


End Sub

Private Sub cmdFolder_Click()
frmDirect.Show 1

End Sub

Private Sub cmdPopulate_Click()
'Populate the remote list according
'to the settings entered

If txtAddy.Text <> "" Then

lstRemote.Clear
txtStatus.Text = txtStatus.Text & vbNewLine & "Populating list with links..."
Dim i#
For i = txtLow.Text To txtHigh.Text
lblNum.Caption = i

If chkLeading.Value = 1 Then
If lblNum.Caption = "10" Then
lblNum.Caption = String(Len(txtFormat) - 1, "0") & "10"
End If


Select Case Val(lblNum)


Case Is < 10
lblNum.Caption = String(Len(txtFormat), "0") & Val(lblNum.Caption)

Case Is < 100
If lblNum.Caption > 10 Then
lblNum.Caption = String(Len(txtFormat) - 1, "0") & Val(lblNum.Caption)
End If

End Select
End If

txtNav.Text = Replace(txtAddy.Text, "%%NUM%%", lblNum.Caption)
lstRemote.AddItem txtNav.Text

Next i



txtStatus.Text = txtStatus.Text & vbNewLine & "Finished creating list"
lngExtent = 3 * (lstRemote.Width / Screen.TwipsPerPixelX)
lngReturn = SendMessage(lstRemote.hWnd, LB_SETHORIZONTALEXTENT, lngExtent, 0&)

End If


End Sub

Private Sub cmdPreview_Click()
'Decide which list was clicked last
'(remote/local) and previews accordingly

On Error GoTo Error

Select Case txtOnlineGal.Text
Case 1

frmGallery.Visible = False

If lstRemote.Text <> "" Then
HyperJump lstRemote.Text
End If

Case 0

If lstFolder.Text <> "" Then
frmGallery.Img2.Picture = LoadPicture(cboFolder.Text & lstFolder.Text)
frmGallery.Img1.Picture = frmGallery.Img2.Picture
frmGallery.Caption = frmMain.cboFolder.Text & frmMain.lstFolder.Text

If frmGallery.Visible = False Then
frmGallery.Show
End If
frmGallery.ZOrder (0)
End If

End Select


Exit Sub
Error:
MsgBox "There was an error reading the picture.", vbCritical, "Error"

End Sub



Private Sub Form_Load()

On Error Resume Next


Dim Admin As String



'Only allow one instance of the program
If App.PrevInstance = True Then
If Command = "" Then
End
Else
Call ExecuteCommand(Command)
End If
End If


'Add an option to clear the combo box
cboFolder.AddItem "[clear list]"

txtAddy.Text = GetSetting(App.Title, "Settings", "URL")
txtPrefix.Text = GetSetting(App.Title, "Settings", "Prefix")
chkLeading.Value = GetSetting(App.Title, "Settings", "Leading")
txtFormat.Text = GetSetting(App.Title, "Settings", "Format")
txtLow.Text = GetSetting(App.Title, "Settings", "Low")
txtHigh.Text = GetSetting(App.Title, "Settings", "High")
frmShowDefsO.txtDelay.Text = GetSetting(App.Title, "Settings", "ODelay")
frmOnline.tmr.Interval = frmShowDefsO.txtDelay * 1000
frmShowDefs.txtDelay.Text = GetSetting(App.Title, "Settings", "Delay")
frmGallery.tmr.Interval = frmShowDefs.txtDelay * 1000
frmGallery.chkStretch.Value = GetSetting(App.Title, "Settings", "Stretch")
frmGallery.chkScale.Value = GetSetting(App.Title, "Settings", "Scale")

'Load our recently used directory list
Loadlistbox App.Path & "\RecentDirs.oVoy", lstDirs
lstDirs.AddItem "C:\"
xListKillDupes lstDirs

'If there's anything to load, put it in the combo box
If lstDirs.ListCount <> 0 Then
Dim X%
For X = 0 To lstDirs.ListCount - 1
lstDirs.ListIndex = X
cboFolder.AddItem lstDirs.Text
Next X
End If


'The directory last used gets used upon opening
cboFolder.Text = GetSetting(App.Title, "Settings", "Directory")
If cboFolder.Text = "" Then
cboFolder.Text = "C:\"
End If

'load the presets
Loadlistbox "C:\Presets.oVoy", lstPresets

'If there are presets to load, throw them in the dynamic menu
If lstPresets.ListCount <> 0 Then
For X = 0 To lstPresets.ListCount - 1
lstPresets.ListIndex = X
lstPresets.Selected(X) = True
Load mnuAddAr(mnuAddAr.UBound + 1)
mnuAddAr(mnuAddAr.UBound).Caption = lstPresets.Text
mnuAddAr(mnuAddAr.UBound).Visible = True
Load mnuDeleteAr(mnuDeleteAr.UBound + 1)
mnuDeleteAr(mnuDeleteAr.UBound).Caption = lstPresets.Text
mnuDeleteAr(mnuDeleteAr.UBound).Visible = True
Next X
End If


'This decides whether or not to show the admin menu
'The only thing you can do there is change the update messsage
Admin = GetSetting(App.Title, "Settings", "Unlock")
'I used a gibberish string for the menu code
If Admin = "jhjh$%^M%IK%^v$%FH^ghNn*67^$%3243$%fd&*%^&^retlohugikuyfikyutfol" Then
mnuAdmin.Visible = True
End If


lblNum.Caption = txtLow.Text

If chkLeading.Value = 1 Then
If lblNum.Caption = "10" Then
lblNum.Caption = String(Len(txtFormat) - 1, "0") & "10"
End If


Select Case Val(lblNum)


Case Is < 10
lblNum.Caption = String(Len(txtFormat), "0") & Val(lblNum.Caption)

Case Is < 100
If lblNum.Caption > 10 Then
lblNum.Caption = String(Len(txtFormat) - 1, "0") & Val(lblNum.Caption)
End If

End Select
End If

Call cmdPopulate_Click

txtOnlineGal.Text = "1"
'Custom Protocol stuff
Dim strPath As String
Dim strFile As String
Dim strFileName As String
    
    Call MakeFile
    
    Call CheckInstance
    
    strPath = App.Path
    strFile = App.EXEName & ".exe"
    
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    strFileName = strPath & strFile
    
    If Not IsProtocolForMe(PTC_PROTOCOL) Then Call AddProtocol(PTC_PROTOCOL, PTC_APPNAME, strFileName)
    
Call ExecuteCommand(Command)

'Add only picture files to the local folder list box
addTheFiles

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

'This would change along with newer versions
'ex. V1.2, V1.3 etc and shows a message box telling
'what's been updated. The box only shows one time per update.
If GetSetting(App.Title, "Settings", "V1.1") <> "E5D2G0W2W1F6H8" Then
MsgBox "What's new in version 1.1:" & vbNewLine & vbNewLine & "- You can add/delete your own presets" & vbNewLine & "- Fixed small quirks here and there", vbInformation, "Updated"
SaveSetting App.Title, "Settings", "V1.1", "E5D2G0W2W1F6H8"
End If


End Sub

Private Sub Form_Resize()

'Bunch of nifty control resizing, moving and shaking going on here
If Me.WindowState <> 1 Then

If Me.Width <= 10575 Then
Me.Width = 10575
End If

If Me.Height <= 9210 Then
Me.Height = 9210
End If


txtStatus.Width = Me.ScaleWidth - 20
txtStatus.Top = Me.ScaleHeight - txtStatus.Height - 20
fraRemote.Left = 20
fraRemote.Width = (Me.Width - 1575) / 2
lstRemote.Width = fraRemote.Width - 200
fraRemote.Height = Me.ScaleHeight - txtStatus.Height - 40 - fraRemote.Top
lstRemote.Height = fraRemote.Height - 250
fraLocal.Left = fraRemote.Left + fraRemote.Width + 1575 - 200
fraLocal.Width = (Me.Width - 1575) / 2
lstFolder.Width = fraLocal.Width - 200
fraLocal.Height = Me.ScaleHeight - txtStatus.Height - 40 - fraRemote.Top
lstFolder.Height = fraLocal.Height - cboFolder.Height - 300
cmdFolder.Left = lstFolder.Width - cmdFolder.Width
cboFolder.Width = cmdFolder.Left - 200
cmdPreview.Left = fraRemote.Width + ((1575 - cmdPreview.Width) / 2) / 2
cmdDownload.Left = fraRemote.Width + ((1575 - cmdDownload.Width) / 2) / 2
cmdDownloadAll.Left = fraRemote.Width + ((1575 - cmdDownloadAll.Width) / 2) / 2
fraURL.Width = Me.ScaleWidth - 40
txtAddy.Width = fraURL.Width - txtAddy.Left - 100
cmdPopulate.Left = fraRemote.Width + ((1575 - cmdPopulate.Width) / 2) / 2
cmdDelete.Left = cmdDownload.Left
cmdCreateLink.Left = fraURL.Width - 100 - cmdCreateLink.Width

End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
'Close
End

End Sub
Private Sub Inet1_StateChanged(ByVal State As Integer)
'Keep the status box up to date on inet info

Select Case State
    Case 1
        txtStatus.Text = txtStatus.Text & vbNewLine & "Trying to resolve host..."
    Case 2
        txtStatus.Text = txtStatus.Text & vbNewLine & "Host is resolved"
    Case 3
        txtStatus.Text = txtStatus.Text & vbNewLine & "Sending connection request..."
    Case 4
        txtStatus.Text = txtStatus.Text & vbNewLine & "Connected"
    Case 5
        txtStatus.Text = txtStatus.Text & vbNewLine & "Sending request..."
    Case 6
        txtStatus.Text = txtStatus.Text & vbNewLine & "Request sent"
    Case 7
        txtStatus.Text = txtStatus.Text & vbNewLine & "Receiving response..."
    Case 8
        txtStatus.Text = txtStatus.Text & vbNewLine & "Response received"
    Case 9
        txtStatus.Text = txtStatus.Text & vbNewLine & "Disconnecting..."
    Case 10
        txtStatus.Text = txtStatus.Text & vbNewLine & "Disconnected"
    Case 11
        txtStatus.Text = txtStatus.Text & vbNewLine & "ERROR "
        NoInternetConnection Inet1.ResponseCode, Inet1.ResponseInfo
        Inet1.Cancel
    End Select
End Sub

Private Sub lstFolder_Click()
'A Boolean variable should have been used for this
'but I'm a little more visual and a 0/1 label was easier
'to remember for me. =)~
txtOnlineGal.Text = "0"
txtCurFile.Text = lstFolder.Text

End Sub

Private Sub lstFolder_DblClick()
'Auto-preview when double clicked
Call cmdPreview_Click

End Sub

Private Sub lstFolder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Show the popup menu
On Error GoTo Error
If Button = 2 Then
If lstFolder.Selected(lstFolder.ListIndex) = True Then
If lstFolder.Text <> "" Then
PopupMenu mnuLocal
End If
End If
End If
Error:
End Sub

Private Sub lstFolder_Scroll()
'Again with the boolean alternative
txtOnlineGal.Text = "0"
txtCurFile.Text = lstFolder.Text

End Sub

Private Sub lstRemote_Click()
'Another place where boolean would be used
txtOnlineGal.Text = "1"

End Sub

Private Sub lstRemote_DblClick()
'Auto preview on double click
Call cmdPreview_Click

End Sub

Private Sub lstRemote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Another popup menu
If Button = 2 Then
If lstRemote.Text <> "" Then
PopupMenu mnuRemote
End If
End If

End Sub

Private Sub lstRemote_Scroll()
'Another boolean place
txtOnlineGal.Text = "1"

End Sub

Private Sub mnuAbout_Click()
'About form
frmPAbout.Show 1

End Sub

Private Sub mnuAddAr_Click(index As Integer)
'This is the dynamic menu so it's a little tricky

If mnuAddAr(index).Caption <> "" Then

'Get the matching registry setting from the menu caption
txtSave.Text = GetSetting(App.Title, "Presets", mnuAddAr(index).Caption)

'Decrypt the file
txtSave.Text = DeHex(txtSave)

'My little extraction function
GetBetween "Addy", txtSave
txtAddy.Text = Extracted
GetBetween "Low", txtSave
txtLow.Text = Extracted
GetBetween "High", txtSave
txtHigh.Text = Extracted
GetBetween "Chk", txtSave
chkLeading.Value = Extracted
GetBetween "Format", txtSave
txtFormat.Text = Extracted
GetBetween "Delay", txtSave
txtDelay.Text = Extracted

lblNum.Caption = txtLow.Text

If chkLeading.Value = 1 Then

If lblNum.Caption = "10" Then
lblNum.Caption = String(Len(txtFormat) - 1, "0") & "10"
End If

Select Case Val(lblNum)

Case Is < 10
lblNum.Caption = String(Len(txtFormat), "0") & Val(txtLow.Text)

Case Is < 100
If lblNum.Caption > 10 Then
lblNum.Caption = String(Len(txtFormat) - 1, "0") & Val(txtLow.Text)
End If

End Select
Else
lblNum.Caption = txtLow.Text

End If

Call cmdPopulate_Click


End If

End Sub

Private Sub mnuAddPreset_Click()
'Add presets to the menu
'This instr stuff just grabs a default name
'for the new preset name
If InStr(1, txtAddy.Text, "http://www.", vbTextCompare) <> 0 Then
txtAddy.SelStart = InStr(1, txtAddy.Text, "http://www.", vbTextCompare) + 10
txtAddy.SelLength = InStr((InStr(1, txtAddy.Text, "http://www.", vbTextCompare)), txtAddy.Text, ".", vbTextCompare) - InStr(1, txtAddy.Text, "http://www.", vbTextCompare)
Text1.Text = txtAddy.SelText

Else

If InStr(1, txtAddy.Text, "http://", vbTextCompare) <> 0 Then
txtAddy.SelStart = InStr(1, txtAddy.Text, "http://", vbTextCompare) + 6
txtAddy.SelLength = InStr(1, txtAddy.Text, ".", vbTextCompare) - 8
Text1.Text = txtAddy.SelText
End If
End If

'Put the default name into the textbox
frmPreset.txtPreset.Text = Text1.Text
txtAddy.SelStart = 0

'show the preset form
frmPreset.Show 1

'Save all the configuration settings into a flat file
txtSave.Text = "$%$preset$%$" & frmPreset.txtPreset.Text & "$%$preset$%$" & Chr(13) & Chr(10) & _
"$%$Addy$%$" & txtAddy.Text & "$%$Addy$%$" & Chr(13) & Chr(10) & _
"$%$Low$%$" & txtLow.Text & "$%$Low" & Chr(13) & Chr(10) & _
"$%$High$%$" & txtHigh.Text & "$%$High$%$" & Chr(13) & Chr(10) & _
"$%$Chk$%$" & chkLeading.Value & "$%$Chk$%$" & Chr(13) & Chr(10) & _
"$%$Format$%$" & txtFormat.Text & "$%$Format$%$" & Chr(13) & Chr(10) & _
"$%$Delay$%$" & txtDelay.Text & "$%$Delay$%$"

'I used a simple hex encryption so people were
'likely to mess around with them and possibly
'give themselves errors
txtSave.Text = EnHex(txtSave)

'Save the name into our presets area in the registry
SaveSetting App.Title, "Presets", frmPreset.txtPreset.Text, txtSave.Text

'Load the new preset into the presets menu
'AND the delete menu
Load mnuAddAr(mnuAddAr.UBound + 1)
mnuAddAr(mnuAddAr.UBound).Caption = frmPreset.txtPreset.Text
mnuAddAr(mnuAddAr.UBound).Visible = True
Load mnuDeleteAr(mnuDeleteAr.UBound + 1)
mnuDeleteAr(mnuDeleteAr.UBound).Caption = frmPreset.txtPreset.Text
mnuDeleteAr(mnuDeleteAr.UBound).Visible = True
lstPresets.AddItem frmPreset.txtPreset.Text

'Save the new list
SaveListBox "C:\Presets.oVoy", lstPresets

End Sub

Private Sub mnuAdminUP_Click()

On Error Resume Next

Dim updateFile As String 'Remote text file to read

'Read the update file and load it up for us
frmAdmin.Inet1.RemotePort = "80"
frmAdmin.txtRawInfo.Text = frmAdmin.Inet1.OpenURL(updateFile)
GetBetween "updatemessage", frmAdmin.txtRawInfo
frmAdmin.txtUpdate = Extracted
GetBetween "ver", frmAdmin.txtRawInfo
frmAdmin.txtURL.Text = Extracted

frmAdmin.Show

End Sub

Private Sub mnuCopy_Click()
'Copy a single picture address to the clipboard
Clipboard.Clear
Clipboard.SetText lstRemote.Text

End Sub

Private Sub mnuDelete_Click()
'Delete
Call cmdDelete_Click

End Sub

Private Sub mnuDeleteAr_Click(index As Integer)
'WARNING: HERE I GOT LAZY
'This only deletes the preset from the menu
'It will still sit in the registry, further
'steps should be taken to do so
Dim X As Integer

For X = 0 To lstPresets.ListCount - 1
lstPresets.ListIndex = X
lstPresets.Selected(X) = True

If lstPresets.Text = mnuDeleteAr(index).Caption Then
lstPresets.RemoveItem X
GoTo Continue
End If

Next X

Exit Sub
Continue:
Unload mnuDeleteAr(index)
Unload mnuAddAr(index)
SaveListBox "C:\Presets.oVoy", lstPresets


End Sub

Private Sub mnuDownloadNow_Click()
'Download
Call cmdDownload_Click

End Sub

Private Sub mnuExit_Click()
'Exit
End

End Sub

Private Sub mnuGoogle_Click()
'Quick google image search

Me.WindowState = 1
frmGoogle.Show

End Sub

Private Sub mnuGS_Click()
'Getting started

Me.WindowState = 1
frmGetting.Show

End Sub

Private Sub mnuLes_Click()
'I added a single default preset to
'show how the configuration works

txtAddy.Text = "http://ourworld.compuserve.com/homepages/Stefan_Radke/images/bssbld/%%NUM%%.jpg"
chkLeading.Value = "1"
txtHigh.Text = "13"
txtLow.Text = "1"
txtFormat.Text = "0"
Call cmdPopulate_Click

End Sub

Private Sub mnuOpen_Click()
On Error GoTo Cancel
'Open saved .pr0n files
'You can rename the extensions to anything
'you'd like, just do it throughout the
'entire program


'common dialog stuff
Cmdlg1.InitDir = App.Path
Cmdlg1.Filter = "pr0n files (*.pr0n)|*.pr0n"
Cmdlg1.DialogTitle = "Open..."


Cmdlg1.ShowOpen
Open Cmdlg1.FileName For Input As #1
txtSave.Text = Input$(LOF(1), #1)
Close #1

'That status again
txtStatus.Text = txtStatus.Text & vbNewLine & Cmdlg1.FileTitle & " - loaded"

'Decrypt the file
txtSave.Text = DeHex(txtSave)


'Parse out the settings
GetBetween "Addy", txtSave
txtAddy.Text = Extracted
GetBetween "Low", txtSave
txtLow.Text = Extracted
GetBetween "High", txtSave
txtHigh.Text = Extracted
GetBetween "Chk", txtSave
chkLeading.Value = Extracted
GetBetween "Format", txtSave
txtFormat.Text = Extracted
GetBetween "Delay", txtSave
txtDelay.Text = Extracted

lblNum.Caption = txtLow.Text

If chkLeading.Value = 1 Then

If lblNum.Caption = "10" Then
lblNum.Caption = String(Len(txtFormat) - 1, "0") & "10"
End If

Select Case Val(lblNum)

Case Is < 10
lblNum.Caption = String(Len(txtFormat), "0") & Val(txtLow.Text)

Case Is < 100
If lblNum.Caption > 10 Then
lblNum.Caption = String(Len(txtFormat) - 1, "0") & Val(txtLow.Text)
End If

End Select
Else
lblNum.Caption = txtLow.Text

End If

Call cmdPopulate_Click

Exit Sub
Cancel:
Exit Sub

End Sub

Private Sub mnuOpts_Click()
'Options
'I ended up not using this, but it's
'here anyway :)
frmOvoyOpts2.Show 1

End Sub

Private Sub mnuPreviewPic_Click()
'Preview
Call cmdPreview_Click

End Sub

Private Sub mnuSave_Click()
'Save settings into a flat file
txtSave.Text = ""

On Error GoTo Cancel
txtSave.Text = "$%$Addy$%$" & txtAddy.Text & "$%$Addy$%$" & Chr(13) & Chr(10) & _
"$%$Low$%$" & txtLow.Text & "$%$Low" & Chr(13) & Chr(10) & _
"$%$High$%$" & txtHigh.Text & "$%$High$%$" & Chr(13) & Chr(10) & _
"$%$Chk$%$" & chkLeading.Value & "$%$Chk$%$" & Chr(13) & Chr(10) & _
"$%$Format$%$" & txtFormat.Text & "$%$Format$%$" & Chr(13) & Chr(10) & _
"$%$Delay$%$" & txtDelay.Text & "$%$Delay$%$"

txtSave.Text = EnHex(txtSave)

If InStr(1, txtAddy.Text, "http://www.", vbTextCompare) <> 0 Then
txtAddy.SelStart = InStr(1, txtAddy.Text, "http://www.", vbTextCompare) + 10
txtAddy.SelLength = InStr((InStr(1, txtAddy.Text, "http://www.", vbTextCompare)), txtAddy.Text, ".", vbTextCompare) - InStr(1, txtAddy.Text, "http://www.", vbTextCompare)
Text1.Text = txtAddy.SelText

Else

If InStr(1, txtAddy.Text, "http://", vbTextCompare) <> 0 Then
txtAddy.SelStart = InStr(1, txtAddy.Text, "http://", vbTextCompare) + 6
txtAddy.SelLength = InStr(1, txtAddy.Text, ".", vbTextCompare) - 8
Text1.Text = txtAddy.SelText
End If
End If


txtAddy.SelStart = 1

Cmdlg1.FileName = Text1.Text & ".pr0n"
Cmdlg1.InitDir = App.Path
Cmdlg1.Filter = "pr0n files (*.pr0n)|*.pr0n"
Cmdlg1.DialogTitle = "Save as..."


Cmdlg1.ShowSave
Open Cmdlg1.FileName For Output As #1
Print #1, txtSave.Text
Close #1

txtStatus.Text = txtStatus.Text & vbNewLine & Cmdlg1.FileTitle & " - saved"

Exit Sub
Cancel:
Exit Sub
End Sub

Private Sub mnuUpdates_Click()

'Check for updates!

frmUpdate2.Show
If frmUpdate2.Inet1.StillExecuting Then frmUpdate2.Inet1.Cancel

'You're probably thinking this should be
'a global declaration. Well, in the finished
'program, this variable doesn't exist. Since
'the update location stays the same, it's
'easier to just put the address in OpenURL

Dim updatFile As String

frmUpdate2.Inet1.RemotePort = "80"
frmUpdate2.txtRawInfo.Text = frmUpdate2.Inet1.OpenURL(updateFile)
GetBetween "updatemessage", frmUpdate2.txtRawInfo
frmUpdate2.txtUpdateInfo = Extracted
GetBetween "ver", frmUpdate2.txtRawInfo
frmUpdate2.lblCurVer.Caption = Extracted

'See if the versions match and give an option
'to update now
If frmUpdate2.lblCurVer <> frmUpdate2.lblVer Then
Dim a%
a% = MsgBox("There is an update available. Would you like to download it now?", vbYesNo, "Update available")
If a = 6 Then

'If they want to update, run the updater
On Error GoTo Error
SaveSetting "MSVUD", "UD", "Key", "dfjkhijhg#rt%"
Shell App.Path & "\updater.exe", vbNormalFocus
End
Exit Sub
Error:
'If the updater gets renamed or removed...
MsgBox "Couldn't find updater.exe. Make sure it's in the same folder as oVoy.", vbCritical, "Error"
SaveSetting "MSVUD", "UD", "Key", "ghgh55$%ijhg#rt%"
End If

End If


End Sub

Private Sub mnuViewGal_Click()
'preview
Call cmdPreview_Click

End Sub

Private Sub Tlbr1_ButtonClick(ByVal Button As MSComctlLib.Button)
'The toolbar button click handlers

Select Case Button.Key

Case "Open"
Call mnuOpen_Click

Case "Save"
Call mnuSave_Click

Case "Preview"
Call cmdPreview_Click

Case "LocalGallery"
frmGallery.Show
frmGallery.ZOrder (0)

Case "Updates"
Call mnuUpdates_Click


Case "BrowseOnline"

Me.WindowState = 1
frmOnline.Show
End Select


End Sub

Private Sub tmr1_Timer()

'The download all feature

On Error Resume Next

If cmdDownloadAll.Caption <> "Stop" Then

If lstRemote.ListIndex < lstRemote.ListCount - 1 Then
Call cmdDownload_Click
lstRemote.ListIndex = lstRemote.ListIndex + 1
lstRemote.Selected(lstRemote.ListIndex) = True
Else
txtStatus.Text = txtStatus.Text & vbNewLine & "Finished downloading"
Inet1.Cancel
cmdDownloadAll.Caption = "Download All"
lstRemote.Enabled = True
cmdPopulate.Enabled = True
txtPrefix.Enabled = True
tmr1 = False
End If

End If

Exit Sub

End Sub

Private Sub tmrCommand_Timer()
'Custom protocol stuff

 Dim strCommand As String
    Dim strFileName As String
    Dim FileNum As Integer
    Dim intLen As Integer
    
    strFileName = App.Path
    If Right(strFileName, 1) <> "\" Then strFileName = strFileName & "\"
    strFileName = strFileName & "command.tmp"
    
    intLen = FileLen(strFileName)
    
    If (intLen > Len(FILE_HEADER)) Then
        strCommand = GetCommandFromFile()
        Call ExecuteCommand(strCommand)
        
        FileNum = FreeFile
        Open strFileName For Output As FileNum
        Close FileNum
        
        FileNum = FreeFile
        Open strFileName For Binary Access Write As FileNum
            Put #FileNum, , FILE_HEADER
        Close FileNum
    Else
        Exit Sub
    End If
End Sub



Private Sub txtAddy_Change()
SaveSetting App.Title, "Settings", "URL", txtAddy.Text

End Sub

Private Sub txtFormat_Change()
If txtFormat.Text = "" Then
txtFormat.Text = "0"
End If

SaveSetting App.Title, "Settings", "Format", txtFormat.Text


End Sub

Private Sub txtFormat_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
GoTo MoveOn
End If

KeyAscii = Asc("0")
MoveOn:

End Sub

Private Sub txtHigh_Change()
lblNum.Caption = txtLow.Text

If txtHigh.Text = "" Then
txtHigh.Text = "2"
End If


If chkLeading.Value = 1 Then

If lblNum.Caption = "10" Then
lblNum.Caption = String(Len(txtFormat) - 1, "0") & "10"
End If

Select Case Val(lblNum)

Case Is < 10
lblNum.Caption = String(Len(txtFormat), "0") & Val(lblNum.Caption)

Case Is < 100
If lblNum.Caption > 10 Then
lblNum.Caption = String(Len(txtFormat) - 1, "0") & Val(lblNum.Caption)
End If

End Select
End If

SaveSetting App.Title, "Settings", "High", txtHigh.Text


End Sub

Private Sub txtHigh_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
KeyAscii = 8
GoTo MoveOn
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If

MoveOn:
If txtHigh.Text = "" Then
txtHigh.Text = "2"
End If

Select Case Val(lblNum)


Case Is < 10
lblNum.Caption = String(Len(txtFormat), "0") & Val(lblNum.Caption)

Case Is < 100
If lblNum.Caption > 10 Then
lblNum.Caption = String(Len(txtFormat) - 1, "0") & Val(lblNum.Caption)
End If

End Select
End Sub

Private Sub txtLow_Change()
lblNum.Caption = txtLow.Text

If txtLow.Text = "" Then
txtLow.Text = "1"
End If


If chkLeading.Value = 1 Then
If lblNum.Caption = "10" Then
lblNum.Caption = String(Len(txtFormat) - 1, "0") & "10"
End If


Select Case Val(lblNum)


Case Is < 10
lblNum.Caption = String(Len(txtFormat), "0") & Val(lblNum.Caption)

Case Is < 100
If lblNum.Caption > 10 Then
lblNum.Caption = String(Len(txtFormat) - 1, "0") & Val(lblNum.Caption)
End If

End Select
End If

SaveSetting App.Title, "Settings", "Low", txtLow.Text


End Sub

Private Sub txtLow_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then
KeyAscii = 8
GoTo MoveOn
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
KeyAscii = 0
End If
MoveOn:

If txtLow.Text = "" Then
txtLow.Text = "1"
End If

Select Case Val(lblNum)


Case Is < 10
lblNum.Caption = String(Len(txtFormat), "0") & Val(lblNum.Caption)

Case Is < 100
If lblNum.Caption > 10 Then
lblNum.Caption = String(Len(txtFormat) - 1, "0") & Val(lblNum.Caption)
End If

End Select
End Sub

Private Sub txtPrefix_Change()
'Unlock the admin menu
'I just did it this way because average users will
'never even know it's there whereas if you put
'some sort of option called 'admin login'
'people might fish around looking how to get in it

If txtPrefix.Text = "!showmetheadminmenu!" Then
SaveSetting App.Title, "Settings", "Unlock", "jhjh$%^M%IK%^v$%FH^ghNn*67^$%3243$%fd&*%^&^retlohugikuyfikyutfol"
mnuAdmin.Visible = True
txtPrefix.Text = ""
End If

SaveSetting App.Title, "Settings", "Prefix", txtPrefix.Text

End Sub

Private Sub txtStatus_Change()
txtStatus.SelStart = Len(txtStatus)

End Sub

Sub addTheFiles()
'add only picture files to the picture list

Dim B As String
frmMain.lstFolder.Clear
B = Dir(cboFolder.Text & "*.*")
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

Private Function GetHTTPFileSize(strHTTPFile As String) As Long
'Get a remote file's size

On Error GoTo ErrorHandler
    Dim GetValue As String
    Dim GetSize  As Long
    
    m_GettingFileSize = True
    
    Inet1.Execute strHTTPFile, "HEAD " & Chr(34) & strHTTPFile & Chr(34)

    Do Until Inet1.StillExecuting = False
        DoEvents
    Loop

    GetValue = Inet1.GetHeader("Content-length")
    
    Do Until Inet1.StillExecuting = False
        DoEvents
    Loop
    
    If IsNumeric(GetValue) = True Then
        GetSize = CLng(GetValue)
    Else
        GetSize = -1
    End If

    If GetSize <= 0 Then GetSize = -1

    m_GettingFileSize = False
    GetHTTPFileSize = GetSize
Exit Function

ErrorHandler:
    m_GettingFileSize = False
    GetHTTPFileSize = -1
End Function

Private Sub NoInternetConnection(ResponseCode As String, ResponseInfo As String)

On Error Resume Next
    Select Case ResponseCode
        Case 12007: MsgBox "You are not connected to the internet, or the server is temporarly offline." & vbCrLf & _
                           "Check your internet settings and try again.", vbOKOnly, "Check if you're connected to the internet"
        Case 35761: MsgBox "There has been occured a problem downloading." & vbCrLf & _
                           "It looks like your ISP is having a few problems." & vbCrLf & _
                           "Check your internet settings and try again.", vbOKOnly, "Problems downloading files"
        Case Else:  MsgBox "There has been a problem while downloading." & vbCrLf & _
                           "It looks like your ISP is having a few problems." & vbCrLf & _
                           "Check your internet settings and try again.", vbOKOnly, "Problems downloading files"
    End Select
End Sub

Public Sub DeleteFile(FileName As String)

    On Error GoTo DelError
    Kill FileName
    Exit Sub
DelError:
    MsgBox "Error deleting File", vbCritical, "Error"
End Sub
Public Function EnHex(Data As String) As String
'simple hex encryption

    Dim iCount As Double
    Dim sTemp As String

    For iCount = 1 To Len(Data)
        sTemp = Hex$(Asc(Mid$(Data, iCount, 1)))
        If Len(sTemp) < 2 Then sTemp = "0" & sTemp
        EnHex = EnHex & sTemp
    Next iCount

End Function


Public Function DeHex(Data As String) As String
'Decryption

    Dim iCount As Double

    For iCount = 1 To Len(Data) Step 2
        DeHex = DeHex & Chr$(Val("&H" & Mid$(Data, iCount, 2)))
    Next iCount

End Function

Public Sub GetBetween(strToFind As String, TextBoxToSearch As TextBox)
'My little parser

On Error Resume Next

Dim a As Integer, B As Integer, C As Integer
If InStr(1, TextBoxToSearch, strToFind, vbTextCompare) <> 0 Then
a = InStr(1, TextBoxToSearch, strToFind, vbTextCompare)
B = InStr(a + 3, TextBoxToSearch, strToFind, vbTextCompare)
C = InStrRev(TextBoxToSearch, "$%$", B - 2, vbTextCompare)

TextBoxToSearch.SelStart = C + 2
TextBoxToSearch.SelLength = B - C - 6
Extracted = TextBoxToSearch.SelText

End If

End Sub
Public Sub MakeFile()
'Protocol stuff

    On Error Resume Next
    
    Dim strFileName As String
    Dim intLen As Integer
    Dim FileNum As Integer
    
    strFileName = App.Path
    If Right(strFileName, 1) <> "\" Then strFileName = strFileName & "\"
    strFileName = strFileName & "command.tmp"
    
    intLen = FileLen(strFileName)
    If (intLen < Len(FILE_HEADER)) Then
        
        
        FileNum = FreeFile
        Open strFileName For Output As FileNum
        Close FileNum
        
        FileNum = FreeFile
        Open strFileName For Binary Access Write As FileNum
            Put #FileNum, , FILE_HEADER
        Close FileNum
    End If
End Sub

Public Function FormatHex(ByVal strHex As String) As String
    Do While Len(strHex) < 2
        strHex = "0" & strHex
    Loop
    
    FormatHex = strHex
End Function

Public Function GetCommandFromFile() As String
    Dim strCommand As String
    Dim strHeader As String
    Dim strFileName As String
    Dim FileNum As Integer
    
    strFileName = App.Path
    If Right(strFileName, 1) <> "\" Then strFileName = strFileName & "\"
    strFileName = strFileName & "command.tmp"
    
    FileNum = FreeFile
    Open strFileName For Input As FileNum
        Input #FileNum, strHeader
        Input #FileNum, strCommand
    Close FileNum
    
    GetCommandFromFile = strCommand
End Function

Public Sub WriteCommandToFile(ByVal strCommand As String)
    Dim strFileName As String
    Dim FileNum As Integer
    
    strFileName = App.Path
    If Right(strFileName, 1) <> "\" Then strFileName = strFileName & "\"
    strFileName = strFileName & "command.tmp"
    
    FileNum = FreeFile
    Open strFileName For Output As FileNum
    Close FileNum
    
    FileNum = FreeFile
    Open strFileName For Binary Access Write As FileNum
        Put #FileNum, , FILE_HEADER
        Put #FileNum, , strCommand
    Close FileNum
End Sub

Public Sub CheckInstance()
    Dim strCommand As String
    
    strCommand = Command
    
    If App.PrevInstance Then
        If (Trim(strCommand) = vbNullString) Then
            End
        End If
        
        Call WriteCommandToFile(strCommand)
        End
    End If
End Sub
Public Function IsProtocolForMe(ByVal gpProtocol As String) As Boolean
    Dim strFileName As String
    Dim strPath As String
    
    gpProtocol = Replace(gpProtocol, ":", "")
    gpProtocol = Replace(gpProtocol, " ", "")
    
    strFileName = ReadRegistry(HKEY_CLASSES_ROOT, gpProtocol & "\shell\open\command", "")
    strFileName = UCase(strFileName)
    strFileName = Replace(strFileName, Chr(34), "")
    
    If Right(strFileName, 3) = " %1" Then strFileName = Left(strFileName, Len(strFileName) - 3)
    
    strPath = App.Path
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    IsProtocolForMe = IIf(UCase(strPath & App.EXEName & ".exe") = UCase(strFileName), True, False)
End Function

Public Sub AddProtocol(ByVal apProtocol As String, ByVal apAppName As String, ByVal apFileName As String)
    apProtocol = Replace(apProtocol, ":", "")
    apProtocol = Replace(apProtocol, " ", "")
    apAppName = Trim(apAppName)
    
    Call WriteRegistry(HKEY_CLASSES_ROOT, apProtocol, "", ValString, "URL: " & apAppName & " Protocol")
    Call WriteRegistry(HKEY_CLASSES_ROOT, apProtocol, "URL Protocol", ValString, "")
    Call WriteRegistry(HKEY_CLASSES_ROOT, apProtocol & "\shell", "", ValString, Chr(0))
    Call WriteRegistry(HKEY_CLASSES_ROOT, apProtocol & "\shell\open", "", ValString, Chr(0))
    Call WriteRegistry(HKEY_CLASSES_ROOT, apProtocol & "\shell\open\command", "", ValString, Chr(34) & UCase(apFileName) & Chr(34) & " " & Chr(34) & "%1" & Chr(34))
End Sub
Public Sub ExecuteCommand(ByVal CommandToExecute As String)
'This is the part that adds all the configuration
'settings when someone clicks an oVoy:// link
'and the program launches

On Error GoTo Error

txtCommand.Text = CommandToExecute
txtCommand.Text = Replace(txtCommand.Text, Chr(34), "")
txtCommand.Text = Replace(txtCommand.Text, "ovoy://", "")

Dim pos1%

If InStr(1, txtCommand.Text, ":00", vbTextCompare) <> 0 Or InStr(1, txtCommand.Text, ":0", vbTextCompare) <> 0 Then
chkLeading.Value = 1
txtLow.Text = Mid(txtCommand.Text, 1, InStr(1, txtCommand.Text, "-", vbTextCompare) - 1)
pos1 = InStr(1, txtCommand.Text, "-", vbTextCompare) - 1
txtHigh.Text = Mid(txtCommand.Text, InStr(1, txtCommand.Text, "-", vbTextCompare) + 1, InStr(1, txtCommand.Text, ":", vbTextCompare) - 2 - pos1)
txtFormat.Text = Mid(txtCommand.Text, InStr(1, txtCommand.Text, ":", vbTextCompare) + 1, InStr(1, txtCommand.Text, "@", vbTextCompare) - 1)
txtFormat.Text = Replace(txtFormat.Text, "@", "")
txtAddy.Text = Mid(txtCommand.Text, InStr(1, txtCommand.Text, "@", vbTextCompare) + 1)
txtAddy.Text = Replace(txtAddy.Text, "1ttp://", "http://")

GoTo Cont
ElseIf InStr(1, txtCommand.Text, ":00", vbTextCompare) = 0 Or InStr(1, txtCommand.Text, ":0", vbTextCompare) = 0 Then

If txtCommand.Text <> "" Then
chkLeading.Value = 0
txtLow.Text = Mid(txtCommand.Text, 1, InStr(1, txtCommand.Text, "-", vbTextCompare) - 1)
pos1 = InStr(1, txtCommand.Text, "-", vbTextCompare) - 1
txtHigh.Text = Mid(txtCommand.Text, InStr(1, txtCommand.Text, "-", vbTextCompare) + 1, InStr(1, txtCommand.Text, "@", vbTextCompare) - 2 - pos1)
txtCommand.SelStart = 1
txtCommand.SelLength = InStr(1, txtCommand.Text, "@", vbTextCompare)
txtCommand.SelText = ""
txtAddy.Text = txtCommand.Text
txtAddy.Text = Replace(txtAddy.Text, "1ttp://", "http://")

End If

End If



Cont:
Call cmdPopulate_Click
Exit Sub
Error:
txtAddy.Text = GetSetting(App.Title, "Settings", "URL")
txtPrefix.Text = GetSetting(App.Title, "Settings", "Prefix")
chkLeading.Value = GetSetting(App.Title, "Settings", "Leading")
txtFormat.Text = GetSetting(App.Title, "Settings", "Format")
txtLow.Text = GetSetting(App.Title, "Settings", "Low")
txtHigh.Text = GetSetting(App.Title, "Settings", "High")
End Sub
