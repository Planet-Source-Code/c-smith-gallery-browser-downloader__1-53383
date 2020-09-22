VERSION 5.00
Begin VB.Form frmGetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Getting Started With oVoy2"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGetting.frx":030A
   ScaleHeight     =   3495
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   3495
      Left            =   5640
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lbl5 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGetting.frx":3AF1
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   4
      Top             =   3480
      Width           =   5655
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGetting.frx":3C69
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   5655
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGetting.frx":3D7B
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "1) On the right side of the program, select a folder to download your pictures to."
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
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Getting Started:"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmGetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim top1%
Dim top2%
Dim top3%
Dim top4%
Dim top5%

Private Sub Form_Load()
'Pretty much all the code here
'is for the neat-o scrolling effect
'It's really un-necessary

FormOnTop Me

top1 = lbl1.Top
top2 = lbl2.Top
top3 = lbl3.Top
top4 = lbl4.Top
top5 = lbl5.Top

VScroll1.Max = lbl1.Height + lbl2.Height + lbl3.Height + lbl4.Height + lbl5.Height
VScroll1.Min = 1
VScroll1.SmallChange = 100
VScroll1.LargeChange = 1000
VScroll1.Value = VScroll1.Min




End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.WindowState = 0

End Sub


Private Sub VScroll1_Change()

lbl1.Top = top1 - VScroll1.Value
lbl2.Top = top2 - VScroll1.Value
lbl3.Top = top3 - VScroll1.Value
lbl4.Top = top4 - VScroll1.Value
lbl5.Top = top5 - VScroll1.Value

End Sub

Private Sub VScroll1_Scroll()
lbl1.Top = top1 - VScroll1.Value
lbl2.Top = top2 - VScroll1.Value
lbl3.Top = top3 - VScroll1.Value
lbl4.Top = top4 - VScroll1.Value
lbl5.Top = top5 - VScroll1.Value

End Sub
