VERSION 5.00
Begin VB.Form frmGallery 
   Caption         =   "Gallery"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGallery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   495
      Left            =   3960
      Picture         =   "frmGallery.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Delete Current Picture"
      Top             =   4440
      Width           =   615
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3600
      Top             =   4200
   End
   Begin VB.CommandButton cmdShow 
      Height          =   495
      Left            =   3360
      Picture         =   "frmGallery.frx":114C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Slide Show"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CheckBox chkStretch 
      Appearance      =   0  'Flat
      Caption         =   "Stretch to fit"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   4560
      Width           =   495
   End
   Begin VB.CheckBox chkScale 
      Appearance      =   0  'Flat
      Caption         =   "Constrain proportions"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Image Img2 
      Height          =   975
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Img1 
      Height          =   4455
      Left            =   0
      Picture         =   "frmGallery.frx":1A16
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6720
   End
End
Attribute VB_Name = "frmGallery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkScale_Click()
SaveSetting App.Title, "Settings", "Scale", chkScale.Value

If chkStretch.Value = 0 Then
chkScale.Enabled = False
Img2.Visible = True
Img1.Visible = False
Else
chkScale.Enabled = True
If chkScale.Value = 1 Then
ConstrainProps
Else
Img1.Width = Me.ScaleWidth - 20
Img1.Height = Me.ScaleHeight - 120 - cmdOk.Height
End If
Img2.Visible = False
Img1.Visible = True
End If


End Sub

Private Sub chkStretch_Click()
If chkStretch.Value = 0 Then
chkScale.Enabled = False
Img2.Visible = True
Img1.Visible = False
Else
chkScale.Enabled = True
If chkScale.Value = 1 Then
ConstrainProps
Else
Img1.Width = Me.ScaleWidth - 20
Img1.Height = Me.ScaleHeight - 120 - cmdOk.Height
End If
Img2.Visible = False
Img1.Visible = True
End If

SaveSetting App.Title, "Settings", "Stretch", chkStretch.Value

End Sub



Private Sub cmdBack_Click()
Down1
End Sub



Private Sub cmdDelete_Click()
On Error GoTo Error
If frmMain.lstFolder.Text <> "" Then
Kill frmMain.cboFolder.Text & frmMain.lstFolder.Text
frmMain.addTheFiles
End If
Exit Sub
Error:
MsgBox "An error occured while deleting the file.", vbCritical, "Error"
End Sub

Private Sub cmdForward_Click()
Up1
End Sub



Private Sub cmdOk_Click()
tmr = False

Me.Hide

End Sub

Private Sub cmdShow_Click()
If tmr = False Then
frmShowDefs.Show 1
Else
tmr = False
End If

End Sub



Private Sub Form_Load()

If chkStretch.Value = 0 Then
chkScale.Enabled = False
Img2.Visible = True
Img1.Visible = False
Else
chkScale.Enabled = True
Img2.Visible = False
Img1.Visible = True
End If

Img2.Picture = Img1.Picture
chkScale.ZOrder (0)

cmdOk.Top = Me.ScaleHeight - cmdOk.Height - 20
cmdOk.Left = Me.ScaleWidth - cmdOk.Width - 20
chkStretch.Top = Me.ScaleHeight - chkStretch.Height - 20
cmdBack.Top = Me.ScaleHeight - cmdBack.Height - 80
cmdForward.Top = cmdBack.Top
chkScale.Top = chkStretch.Top + chkStretch.Height
cmdShow.Top = Me.ScaleHeight - cmdShow.Height - 20
cmdDelete.Top = cmdShow.Top

FitAll
End Sub



Private Sub Form_Resize()
If Me.WindowState <> 1 Then

If Me.Width <= 6885 Then
Me.Width = 6885
End If

If Me.Height <= 5335 Then
Me.Height = 5335
End If



cmdOk.Top = Me.ScaleHeight - cmdOk.Height - 20
cmdOk.Left = Me.ScaleWidth - cmdOk.Width - 20
chkStretch.Top = Me.ScaleHeight - (chkStretch.Height * 2) - 20
cmdBack.Top = Me.ScaleHeight - cmdBack.Height - 10
cmdForward.Top = cmdBack.Top
chkScale.Top = chkStretch.Top + chkStretch.Height
cmdShow.Top = Me.ScaleHeight - cmdShow.Height - 20
cmdDelete.Top = cmdShow.Top

If chkStretch.Value = 1 Then
If chkScale.Value = 1 Then
ConstrainProps
Exit Sub
End If
Img1.Height = Me.ScaleHeight - 120 - cmdOk.Height
Img1.Width = Me.ScaleWidth - 20
End If

End If

End Sub

Public Sub ConstrainProps()

'This was a pain and it's still just
'a little buggy

Img2.Height = Img2.Picture.Height
Img2.Width = Img2.Picture.Width

Rat = Img2.Picture.Width / Img2.Picture.Height


    If Rat > 1 Then
        
        If Img1.Width > frmGallery.Width Or Img1.Width < frmGallery.Width Then
         
        Img1.Width = frmGallery.Width
        Img1.Height = Img1.Width / Rat
        
        
        End If
        
            Else
            
        If Img1.Height > frmGallery.Height Or Img1.Height < frmGallery.Height Then
                
        Img1.Height = Me.ScaleHeight - 120 - cmdOk.Height
        Img1.Width = Img1.Height * Rat
        
        End If
        
                
        If Img1.Width >= Img1.Height Then
        Img1.Width = frmGallery.ScaleWidth - 130 - cmdOk.Height
        Img1.Width = frmGallery.Width
        Img1.Height = Img1.Width / Rat
        GoTo done
        Else
        Img1.Height = (Me.ScaleHeight - 130 - cmdOk.Height)
        Img1.Width = Img1.Height * Rat
        GoTo done
        
        End If
        
    End If
    
  

done:
Img1.Refresh

End Sub
Public Sub FitAll()

Me.Width = 6885

Me.Height = 5335


End Sub

Public Sub Up1()

On Error Resume Next

If frmMain.lstFolder.ListCount <> 0 Then

If frmMain.lstFolder.Text = "" Then
frmMain.lstFolder.ListIndex = 0
frmMain.lstFolder.Selected(0) = True
End If

If frmMain.lstFolder.ListCount > 1 Then

If frmMain.lstFolder.ListIndex = frmMain.lstFolder.ListCount - 1 Then
frmMain.lstFolder.Selected(frmMain.lstFolder.ListIndex) = False
frmMain.lstFolder.ListIndex = 0
frmMain.lstFolder.Selected(frmMain.lstFolder.ListIndex) = True
Img1.Picture = LoadPicture(frmMain.cboFolder.Text & frmMain.lstFolder.Text)
Img2.Picture = Img1.Picture
frmGallery.Caption = frmMain.cboFolder.Text & frmMain.lstFolder.Text
GoTo ConstX
End If

frmMain.lstFolder.Selected(frmMain.lstFolder.ListIndex) = False
frmMain.lstFolder.ListIndex = frmMain.lstFolder.ListIndex + 1
frmMain.lstFolder.Selected(frmMain.lstFolder.ListIndex) = True

End If

If frmMain.cboFolder.Text <> "C:\" Then
Img1.Picture = LoadPicture(frmMain.cboFolder.Text & frmMain.lstFolder.Text)
Img2.Picture = Img1.Picture
frmGallery.Caption = frmMain.cboFolder.Text & frmMain.lstFolder.Text
End If

ConstX:
ConstrainProps
End If

End Sub
Public Sub Down1()

On Error Resume Next


If frmMain.lstFolder.ListCount <> 0 Then

If frmMain.lstFolder.Text = "" Then
frmMain.lstFolder.ListIndex = 0
frmMain.lstFolder.Selected(0) = True
End If

If frmMain.lstFolder.ListCount > 1 Then

If frmMain.lstFolder.ListIndex = 0 Then
frmMain.lstFolder.Selected(frmMain.lstFolder.ListIndex) = False
frmMain.lstFolder.ListIndex = frmMain.lstFolder.ListCount - 1
frmMain.lstFolder.Selected(frmMain.lstFolder.ListIndex) = True
Img1.Picture = LoadPicture(frmMain.cboFolder.Text & frmMain.lstFolder.Text)
Img2.Picture = Img1.Picture
frmGallery.Caption = frmMain.cboFolder.Text & frmMain.lstFolder.Text
GoTo ConstX
End If

frmMain.lstFolder.Selected(frmMain.lstFolder.ListIndex) = False
frmMain.lstFolder.ListIndex = frmMain.lstFolder.ListIndex - 1
frmMain.lstFolder.Selected(frmMain.lstFolder.ListIndex) = True

End If

If frmMain.cboFolder.Text <> "C:\" Then
Img1.Picture = LoadPicture(frmMain.cboFolder.Text & frmMain.lstFolder.Text)
Img2.Picture = Img1.Picture
frmGallery.Caption = frmMain.cboFolder.Text & frmMain.lstFolder.Text
End If

ConstX:
ConstrainProps
End If






End Sub

Private Sub Form_Unload(Cancel As Integer)
tmr = False

End Sub

Private Sub tmr_Timer()
If frmShowDefs.cmdDirection.Caption = "Forward" Then
Call cmdForward_Click
Else
Call cmdBack_Click
End If
End Sub
