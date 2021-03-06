VERSION 5.00
Begin VB.Form frmPeaceProp 
   BorderStyle     =   0  'None
   Caption         =   "Ofertas de paz"
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5070
   ClipControls    =   0   'False
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
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lista 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   1785
      ItemData        =   "frmPeaceProp.frx":0000
      Left            =   240
      List            =   "frmPeaceProp.frx":0002
      TabIndex        =   0
      Top             =   600
      Width           =   4620
   End
   Begin VB.Image imgRechazar 
      Height          =   480
      Left            =   3840
      Top             =   2520
      Width           =   960
   End
   Begin VB.Image imgAceptar 
      Height          =   480
      Left            =   2640
      Top             =   2520
      Width           =   960
   End
   Begin VB.Image imgDetalle 
      Height          =   480
      Left            =   1440
      Top             =   2520
      Width           =   960
   End
   Begin VB.Image imgCerrar 
      Height          =   480
      Left            =   240
      Top             =   2520
      Width           =   960
   End
End
Attribute VB_Name = "frmPeaceProp"
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

Private cBotonAceptar  As clsGraphicalButton

Private cBotonCerrar   As clsGraphicalButton

Private cBotonDetalles As clsGraphicalButton

Private cBotonRechazar As clsGraphicalButton

Public LastPressed     As clsGraphicalButton

Private TipoProp       As TIPO_PROPUESTA

Public Enum TIPO_PROPUESTA

    ALIANZA = 1
    PAZ = 2

End Enum

Private Sub Form_Load()
    'Call SetTranslucent(Me.hwnd, NTRANS_GENERAL)

    Call LoadBackGround
    Call LoadButtons

End Sub

Private Sub LoadButtons()

    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonDetalles = New clsGraphicalButton
    Set cBotonRechazar = New clsGraphicalButton
    
    Set LastPressed = New clsGraphicalButton
    
    Call cBotonAceptar.Initialize(imgAceptar, "BotonAceptarOferta.jpg", "BotonAceptarRolloverOferta.jpg", "BotonAceptarClickOferta.jpg", Me)

    Call cBotonCerrar.Initialize(imgCerrar, "BotonCerrarOferta.jpg", "BotonCerrarRolloverOferta.jpg", "BotonCerrarClickOferta.jpg", Me)

    Call cBotonDetalles.Initialize(imgDetalle, "BotonDetallesOferta.jpg", "BotonDetallesRolloverOferta.jpg", "BotonDetallesClickOferta.jpg", Me)

    Call cBotonRechazar.Initialize(imgRechazar, "BotonRechazarOferta.jpg", "BotonRechazarRolloverOferta.jpg", "BotonRechazarClickOferta.jpg", Me)

End Sub

Private Sub LoadBackGround()

    If TipoProp = TIPO_PROPUESTA.ALIANZA Then
        Me.Picture = LoadPictureEX("VentanaOfertaAlianza.jpg")
    Else
        Me.Picture = LoadPictureEX("VentanaOfertaPaz.jpg")

    End If

End Sub

Public Property Let ProposalType(ByVal nValue As TIPO_PROPUESTA)
    TipoProp = nValue

End Property

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then MoverVentana (Me.hwnd)

End Sub

Private Sub imgAceptar_Click()

    If lista.ListIndex = -1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("?Debes selecionar un propuesta!", .red, .green, .blue, .bold, .italic)

        End With

        Exit Sub

    End If

    If TipoProp = PAZ Then
        Call WriteGuildAcceptPeace(lista.List(lista.ListIndex))
    Else
        Call WriteGuildAcceptAlliance(lista.List(lista.ListIndex))

    End If

    If frmGuildLeader.solicitudes.ListIndex > -1 Then
        frmGuildLeader.solicitudes.RemoveItem (frmGuildLeader.solicitudes.ListIndex)

    End If
    
    Me.Hide
    
    Unload Me

End Sub

Private Sub imgCerrar_Click()
    Unload Me

End Sub

Private Sub imgDetalle_Click()

    If lista.ListIndex = -1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("?Debes selecionar un propuesta!", .red, .green, .blue, .bold, .italic)

        End With

        Exit Sub

    End If

    If TipoProp = PAZ Then
        Call WriteGuildPeaceDetails(lista.List(lista.ListIndex))
    Else
        Call WriteGuildAllianceDetails(lista.List(lista.ListIndex))

    End If

End Sub

Private Sub imgRechazar_Click()

    If lista.ListIndex = -1 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("?Debes selecionar un propuesta!", .red, .green, .blue, .bold, .italic)

        End With

        Exit Sub

    End If

    If TipoProp = PAZ Then
        Call WriteGuildRejectPeace(lista.List(lista.ListIndex))
    Else
        Call WriteGuildRejectAlliance(lista.List(lista.ListIndex))

    End If

    If frmGuildLeader.solicitudes.ListIndex > -1 Then
        frmGuildLeader.solicitudes.RemoveItem (frmGuildLeader.solicitudes.ListIndex)

    End If

    Me.Hide
    
    Unload Me

End Sub
