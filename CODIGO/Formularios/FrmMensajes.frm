VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmMensajes 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   7470
   ClientTop       =   2820
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   Picture         =   "FrmMensajes.frx":0000
   ScaleHeight     =   4515
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   840
      Top             =   1680
   End
   Begin RichTextLib.RichTextBox mensajes 
      Height          =   3015
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5318
      _Version        =   393217
      BackColor       =   -2147483641
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"FrmMensajes.frx":19DDF
   End
End
Attribute VB_Name = "FrmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Const LW_KEY = &H1
Const G_E = (-20)
Const W_E = &H80000
Private Sub Form_Load()
    mensajes.LoadFile App.path & "/INIT/" & UserName & ".txt", rtfText
    mensajes.Locked = True

    mensajes.SelLength = Len(mensajes)
    mensajes.SelColor = RGB(255, 255, 255)

    mensajes.SelLength = 0
    Skin Me, vbMagenta
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MoverVentana (Me.hwnd)
End Sub

Private Sub mensajes_GotFocus()
mensajes.SelStart = Len(mensajes.Text)
End Sub

Private Sub mensajes_KeyUp(KeyCode As Integer, Shift As Integer)
mensajes.SelStart = Len(mensajes.Text)
End Sub

Private Sub mensajes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mensajes.SelStart = Len(mensajes.Text)
End Sub


Sub Skin(Frm As Form, Color As Long)
Frm.BackColor = Color
Dim ret As Long
ret = GetWindowLong(Frm.hwnd, G_E)
ret = ret Or W_E
SetWindowLong Frm.hwnd, G_E, ret
SetLayeredWindowAttributes Frm.hwnd, Color, 0, LW_KEY
End Sub

Private Sub Timer1_Timer()
mensajes.LoadFile App.path & "/INIT/" & UserName & ".txt", rtfText
mensajes.SelStart = Len(mensajes.Text)
'mensajes.Refresh
End Sub
