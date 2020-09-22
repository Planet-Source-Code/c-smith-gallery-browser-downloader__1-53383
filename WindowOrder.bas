Attribute VB_Name = "Module1"

Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTTOPMOST = -2
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE



Sub FormNormal(Frm As Form)

Call SetWindowPos(Frm.hWnd, HWND_NOTTOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub



Sub FormOnTop(Frm As Form)



Call SetWindowPos(Frm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub


Public Function GethWndByWinTitle(winTitle As String) As Long
    Dim retval As Long
    GethWndByWinTitle = FindWindow(vbNullString, winTitle)
End Function


