VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Option Base 0

Private Const MAX_PASOS As Byte = 100

Private LastPaso        As Byte

Private Type tPasos

    Pasos       As Integer
    VidaPasos   As Integer
    x           As Integer
    y           As Integer

End Type

Private cPasos(1 To MAX_PASOS) As tPasos
    
Public Sub RenderPasos(ByVal PixelOffSetXTemp As Integer, _
                       ByVal PixelOffSetYTemp As Integer, _
                       ByVal i As Byte)

    On Error GoTo 0

    Dim color(0 To 3) As Long

    With cPasos(i)

        If .VidaPasos > 0 Then

            Dim iGrhIndex As Integer

            .VidaPasos = .VidaPasos - 2
            
            color(0) = D3DColorARGB(.VidaPasos, 255, 255, 255)
            color(1) = color(0)
            color(2) = color(0)
            color(3) = color(0)
        
            iGrhIndex = .Pasos
            
            Call Engine_Render_Rectangle(PixelOffSetXTemp, PixelOffSetYTemp, GrhData(iGrhIndex).PixelWidth, GrhData(iGrhIndex).PixelHeight, GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, GrhData(iGrhIndex).PixelWidth, GrhData(iGrhIndex).PixelHeight, , , 0, GrhData(iGrhIndex).FileNum, color(0), color(1), color(2), color(3))
                        
            If .VidaPasos <= 0 Then
                Call RemovePaso(i)

            End If
                        
        End If

    End With

End Sub

Public Sub CreatePasos(ByVal x As Integer, ByVal y As Integer, ByVal Pasos As Integer)

    On Error GoTo 0

    If Not InMapBounds(x, y) Then Exit Sub
    If HayAgua(x, y) Then Exit Sub
    If Pasos = 0 Then Exit Sub
    
    Dim i    As Byte

    Dim slot As Byte
    
    For i = 1 To MAX_PASOS

        If cPasos(i).VidaPasos = 0 Then
            slot = i
            Exit For

        End If

    Next i
    
    If slot >= MAX_PASOS Then Exit Sub
    If slot > LastPaso Then LastPaso = slot
    
    With cPasos(slot)
        .x = x
        .y = y
        .Pasos = Pasos
        .VidaPasos = 250

    End With
    
    MapData(x, y).PasosIndex = slot
    
End Sub

Public Sub RemovePaso(ByVal i As Integer)

    On Error GoTo 0

    With cPasos(i)

        If i = LastPaso Then LastPaso = LastPaso - 1
        If .x <> 0 And .y <> 0 Then
            MapData(.x, .y).PasosIndex = 0
            .Pasos = 0
            .VidaPasos = 0
            .x = 0
            .y = 0

        End If

    End With

End Sub

Public Sub RemoveAllPasos()

    On Error GoTo 0

    Dim i As Long
    
    For i = LastPaso To 1 Step -1
        RemovePaso i
    Next i

End Sub

