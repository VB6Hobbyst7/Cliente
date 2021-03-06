VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1470
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   3240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   98
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   216
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCantidad 
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
      Height          =   225
      Left            =   495
      MaxLength       =   5
      TabIndex        =   0
      Top             =   540
      Width           =   2205
   End
   Begin VB.Image imgTirarTodo 
      Height          =   375
      Left            =   1680
      Tag             =   "1"
      Top             =   975
      Width           =   1335
   End
   Begin VB.Image imgTirar 
      Height          =   375
      Left            =   210
      Tag             =   "1"
      Top             =   975
      Width           =   1335
   End
End
Attribute VB_Name = "frmCantidad"
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

Private cBotonTirar     As clsGraphicalButton

Private cBotonTirarTodo As clsGraphicalButton

Public LastPressed      As clsGraphicalButton

Private Sub Form_Load()
    'Call SetTranslucent(Me.hwnd, NTRANS_GENERAL)
    Me.Picture = LoadPictureEX("VentanaTirarOro.jpg")
    
    Call LoadButtons

End Sub

Private Sub LoadButtons()
    
    Set cBotonTirar = New clsGraphicalButton
    Set cBotonTirarTodo = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton

    Call cBotonTirar.Initialize(imgTirar, "BotonTirar.jpg", "BotonTirarRollover.jpg", "BotonTirarClick.jpg", Me)
    Call cBotonTirarTodo.Initialize(imgTirarTodo, "BotonTirarTodo.jpg", "BotonTirarTodoRollover.jpg", "BotonTirarTodoClick.jpg", Me)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then MoverVentana (Me.hwnd)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    LastPressed.ToggleToNormal

End Sub

Private Sub imgTirar_Click()

    If LenB(txtCantidad.Text) > 0 Then
        If Not IsNumeric(txtCantidad.Text) Then Exit Sub  'Should never happen
        
        Call WriteDrop(Inventario.SelectedItem, frmCantidad.txtCantidad.Text, Inventario.DropX, Inventario.DropY)
        frmCantidad.txtCantidad.Text = ""

    End If
    
    Unload Me

End Sub

Private Sub imgTirarTodo_Click()

    If Inventario.SelectedItem = 0 Then Exit Sub
    
    If Inventario.SelectedItem <> FLAGORO Then
        Call WriteDrop(Inventario.SelectedItem, Inventario.Amount(Inventario.SelectedItem), Inventario.DropX, Inventario.DropY)
        Unload Me
    Else

        If UserGLD > 10000 Then
            Call WriteDrop(Inventario.SelectedItem, 10000)
            Unload Me
        Else
            Call WriteDrop(Inventario.SelectedItem, UserGLD)
            Unload Me

        End If

    End If

    frmCantidad.txtCantidad.Text = ""

End Sub

Private Sub txtCantidad_Change()

    On Error GoTo ErrHandler

    If Val(txtCantidad.Text) < 0 Then
        txtCantidad.Text = "1"

    End If
    
    If Val(txtCantidad.Text) > MAX_INVENTORY_OBJS Then
        txtCantidad.Text = "10000"

    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    txtCantidad.Text = "1"

End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0

        End If

    End If

End Sub

Private Sub txtCantidad_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        imgTirar_Click
    ElseIf KeyCode = 27 Then
        Unload Me

    End If

End Sub

Private Sub txtCantidad_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
    LastPressed.ToggleToNormal

End Sub
