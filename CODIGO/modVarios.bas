Attribute VB_Name = "modVarios"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const LW_KEY = &H1
Const G_E = (-20)
Const W_E = &H80000
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2


Public Sub Skin(Frm As Form, Color As Long)
Frm.BackColor = Color
Dim Ret As Long
Ret = GetWindowLong(Frm.hwnd, G_E)
Ret = Ret Or W_E
SetWindowLong Frm.hwnd, G_E, Ret
SetLayeredWindowAttributes Frm.hwnd, Color, 0, LW_KEY
End Sub


Public Sub Auto_Drag(ByVal hwnd As Long)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
End Sub

