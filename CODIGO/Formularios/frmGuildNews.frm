VERSION 5.00
Begin VB.Form frmGuildNews 
   BorderStyle     =   0  'None
   Caption         =   "GuildNews"
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   334
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtClanesAliados 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1020
      Left            =   315
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   5040
      Width           =   4275
   End
   Begin VB.TextBox txtClanesGuerra 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1020
      Left            =   315
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3480
      Width           =   4275
   End
   Begin VB.TextBox news 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2100
      Left            =   315
      Locked          =   -1  'True
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   825
      Width           =   4275
   End
   Begin VB.Image imgAceptar 
      Height          =   375
      Left            =   315
      Tag             =   "1"
      Top             =   6240
      Width           =   4350
   End
End
Attribute VB_Name = "frmGuildNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M?rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat?as Fernando Peque?o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n?mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C?digo Postal 1900
'Pablo Ignacio M?rquez

Option Explicit

Private cBotonAceptar As clsGraphicalButton

Public LastPressed    As clsGraphicalButton

Private Sub aliados_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              x As Single, _
                              y As Single)
    LastPressed.ToggleToNormal

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Or KeyCode = 27 Then
        imgAceptar_Click

    End If

End Sub

Private Sub Form_Load()
    'Call SetTranslucent(Me.hwnd, NTRANS_GENERAL)
    Me.Picture = LoadPictureEX("VentanaGuildNews.jpg")
    
    LoadButtons

End Sub

Private Sub LoadButtons()

    Set cBotonAceptar = New clsGraphicalButton
    Set LastPressed = New clsGraphicalButton
    
    Call cBotonAceptar.Initialize(imgAceptar, "BotonAceptarGuildNews.jpg", "BotonAceptarRolloverGuildNews.jpg", "BotonAceptarClickGuildNews.jpg", Me)
   
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then MoverVentana (Me.hwnd)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastPressed.ToggleToNormal

End Sub

Private Sub imgAceptar_Click()

    On Error Resume Next

    Unload Me
    frmMain.SetFocus

End Sub

Private Sub imgAceptar_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 y As Single)
    LastPressed.ToggleToNormal

End Sub

