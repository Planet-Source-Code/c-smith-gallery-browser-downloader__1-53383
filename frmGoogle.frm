VERSION 5.00
Begin VB.Form frmGoogle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Google Image Search"
   ClientHeight    =   555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmGoogle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
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
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "whatever"
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmGoogle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim FindX As String
Private Sub Command1_Click()
'This is all self explanitory
'Does a google image search

FindX = Replace(Text1.Text, " ", "+")
SaveSetting App.Title, "Settings", "Google", Text1.Text

HyperJump "http://images.google.com/images?q=" & FindX & "&ie=UTF-8&oe=UTF-8&hl=en&btnG=Google+Search"


End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.WindowState = 0

End Sub

Private Sub Text1_Change()
FindX = Replace(Text1.Text, " ", "+")
SaveSetting App.Title, "Settings", "Google", Text1.Text
End Sub
Public Sub HyperJump(ByVal URL As String)

    'Function to execute the Hyperlink
    Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If

End Sub
