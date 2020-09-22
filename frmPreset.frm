VERSION 5.00
Begin VB.Form frmPreset 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Name of new preset"
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
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
   ScaleHeight     =   585
   ScaleWidth      =   4680
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
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtPreset 
      Height          =   285
      HideSelection   =   0   'False
      Left            =   120
      MaxLength       =   30
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmPreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If txtPreset.Text <> "" Then

Dim X As Integer

For X = 0 To frmMain.lstPresets.ListCount - 1
frmMain.lstPresets.ListIndex = X
frmMain.lstPresets.Selected(X) = True

If frmMain.lstPresets.Text = txtPreset.Text Then
GoTo Duped
End If

Next X

Me.Hide
frmMain.Show

End If

Exit Sub
Duped:
MsgBox "A preset by that name already exists." & vbNewLine & "Please choose a different name.", vbInformation, "Duplicate name"
Exit Sub

End Sub

Private Sub Form_Load()

FormOnTop Me

End Sub

Private Sub txtPreset_KeyPress(KeyAscii As Integer)
'See if a preset by this name
'already exists, if not, add it

If txtPreset.Text <> "" Then
If KeyAscii = 13 Then
If txtPreset.Text <> "" Then

Dim X As Integer

For X = 0 To frmMain.lstPresets.ListCount - 1
frmMain.lstPresets.ListIndex = X
frmMain.lstPresets.Selected(X) = True

If frmMain.lstPresets.Text = txtPreset.Text Then
GoTo Duped
End If

Next X

Me.Hide

End If

Exit Sub
Duped:
MsgBox "A preset by that name already exists." & vbNewLine & "Please choose a different name.", vbInformation, "Duplicate name"
Exit Sub


End If
End If

End Sub
