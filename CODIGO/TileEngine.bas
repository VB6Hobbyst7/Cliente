Attribute VB_Name = "Mod_TileEngine"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez



Option Explicit
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDCDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long

'Map sizes in tiles
Public Const XMaxMapSize As Integer = 1100
Public Const YMaxMapSize As Integer = 1500

Public Const RelacionMiniMapa As Single = 1.92120075046904

Public Const GrhFogata As Integer = 1521
''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1


'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    x As Long
    y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    x As Integer
    y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Integer
    
    PixelWidth As Integer
    PixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Integer
    
    Speed As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
End Type

'Lista de cuerpos
Type BodyData
    Walk(E_Heading.north To E_Heading.west) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Type HeadData
    Head(E_Heading.north To E_Heading.west) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.north To E_Heading.west) As Grh
    '[ANIM ATAK]
    WeaponAttack As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.north To E_Heading.west) As Grh
    '[ANIM ATAK]
    ShieldAttack As Byte
End Type

Public NPCMuertos As New Collection

'Apariencia del personaje
Public Type Char
    'Render
    Elv As Byte
    Gld As Long
    Clase As Byte

    equitando As Boolean
    congelado As Boolean
    Chiquito As Boolean
    nadando As Boolean
    inmovilizado As Boolean
    
    ACTIVE As Byte
    Heading As E_Heading
    Pos As Position
    LastPos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    fX As Grh
    FxIndex As Integer
    
    Criminal As Byte
    
    nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    logged As Boolean
    muerto As Boolean
    invisible As Boolean
    oculto As Boolean
    Alpha As Byte
    ContadorInvi As Integer
    iTick As Long
    priv As Byte
    
    Quieto As Byte
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

Public Type MapInformation
    name As String
    MapVersion As Integer
    Width As Integer
    Height As Integer
    Offset As Integer
    Date As String
    
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    PasosIndex As Byte

    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    Particle_Group_Index As Integer
    
    Blocked As Byte
    Trigger As Byte
    
    Map As Byte
    Elemento As Object
    
    Light_Value(3) As Long
    Hora As Byte
    
    fX As Integer
    fXGrh As Grh
End Type

Public IniPath As String
Public MapPath As String


'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Private fpsLastCheck As Long

'Tamaño del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer

'Offset del desde 0,0 del main view
Private MainViewTop As Integer
Private MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Public timerElapsedTime As Single
Public timerTicksPerFrame As Single
Public engineBaseSpeed As Single


Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

Public NumChars As Integer
Public LastChar As Integer


Private MainDestRect   As RECT
Private MainViewRect   As RECT
Private BackBufferRect As RECT

Private MainViewWidth As Integer
Private MainViewHeight As Integer

Private MouseTileX As Integer
Private MouseTileY As Integer




'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Arrojas As New Collection
Public Tooltips As New Collection
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData(1 To XMaxMapSize, 1 To YMaxMapSize) As MapBlock ' Mapa
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?
Public bAlpha       As Byte
Public nAlpha       As Byte
Public tTick        As Long
Public ColorTecho   As Long
Public brstTick     As Long

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(7) As Integer

Public charlist(1 To 10000) As Char
Public AperturaPergamino As Single

#If SeguridadAlkon Then

Public MI(1 To 1233) As clsManagerInvisibles
Public CualMI As Integer

#End If

' Used by GetTextExtentPoint32
Private Type Size
    cx As Long
    cy As Long
End Type

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapaY As Single
Public VerMapa As Boolean
Public Entrada As Byte
Public FrameUseMotionBlur As Boolean

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

Public PosMapX As Single
Public PosMapY As Single

Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open App.path & "\init\Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
End Sub

Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    N = FreeFile()
    Open App.path & "\init\Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
End Sub


Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Integer, ByRef tY As Integer)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************

    tX = UserPos.x + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    tY = UserPos.y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
    
    

    'frmMain.lblPosTest2.Caption = "X: " & tX & "; Y:" & tY
   
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .ACTIVE = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        '[ANIM ATAK]
        .Arma.WeaponAttack = 0
        .Escudo.ShieldAttack = 0
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        .Alpha = 255
        .iTick = 0
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.x = x
        .Pos.y = y
        
        .muerto = Head = CASPER_HEAD Or Head = CASPER_HEAD_CRIMI Or Body = FRAGATA_FANTASMAL
        If .muerto Then .Alpha = 80 Else .Alpha = 255
        'Make active
        .ACTIVE = 1
        
    End With
    
    'Plot on map
    MapData(x, y).CharIndex = CharIndex
    
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        .Gld = 0
        .Elv = 1
        .Clase = 0
    
        .equitando = False
        .ACTIVE = 0
        .Criminal = 0
        .FxIndex = 0
        .invisible = False
        
        #If SeguridadAlkon Then
            Call MI(CualMI).ResetInvisible(CharIndex)
        #End If
        
        .Moving = 0
        .muerto = False
        .Alpha = 255
        .iTick = 0
        .ContadorInvi = 0
        .nombre = ""
        .pie = False
        .Pos.x = 0
        .Pos.y = 0
        .LastPos.x = 0
        .LastPos.y = 0
        .UsandoArma = False
    End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next

    With charlist(CharIndex)

    .ACTIVE = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).ACTIVE = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    If .Pos.x > 0 And .Pos.y > 0 Then
    MapData(.Pos.x, .Pos.y).CharIndex = 0
    
    If .FxIndex <> 0 And .fX.Loops > -1 Then
        MapData(.Pos.x, .Pos.y).fX = .FxIndex
        MapData(.Pos.x, .Pos.y).fXGrh = .fX
    End If
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
    End If
    
    End With
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    If Grh.GrhIndex = 0 Then Exit Sub
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************



    Dim AddX As Integer
    Dim AddY As Integer
    Dim x As Integer
    Dim y As Integer
    Dim nX As Integer
    Dim nY As Integer
    Dim tmpInt As Integer
    
    With charlist(CharIndex)
        x = .Pos.x
        y = .Pos.y
        
        If x = 0 Or y = 0 Then Exit Sub
        
        .LastPos.x = x
        .LastPos.y = y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.north
                AddY = -1
        
            Case E_Heading.east
                AddX = 1
        
            Case E_Heading.south
                AddY = 1
            
            Case E_Heading.west
                AddX = -1
        End Select
        
        nX = x + AddX
        nY = y + AddY
        
        If MapData(nX, nY).CharIndex > 0 Then
            tmpInt = MapData(nX, nY).CharIndex
            If charlist(tmpInt).muerto = False Then
                tmpInt = 0
            Else
                charlist(tmpInt).Pos.x = x
                charlist(tmpInt).Pos.y = y
                charlist(tmpInt).Heading = InvertHeading(nHeading)
                charlist(tmpInt).MoveOffsetX = 1 * (TilePixelWidth * AddX)
                charlist(tmpInt).MoveOffsetY = 1 * (TilePixelHeight * AddY)
                
                charlist(tmpInt).Moving = 1
                
                charlist(tmpInt).scrollDirectionX = -AddX
                charlist(tmpInt).scrollDirectionY = -AddY
                
                'Si el fantasma soy yo mueve la pantalla
                If tmpInt = UserCharIndex Then Call MoveScreen(charlist(tmpInt).Heading)
            End If
        Else
            tmpInt = 0
        End If
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.x = nX
        .Pos.y = nY
        MapData(x, y).CharIndex = tmpInt
        
        If UserEstado <> 1 Then
            Call vPasos.CreatePasos(x, y, DamePasos(nHeading))
        End If
        
        .MoveOffsetX = -1 * (TilePixelWidth * AddX)
        .MoveOffsetY = -1 * (TilePixelHeight * AddY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = AddX
        .scrollDirectionY = AddY
    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    'If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
    '    If CharIndex <> UserCharIndex Then
    '        Call EraseChar(CharIndex)
    '    End If
    'End If
End Sub
Public Function InvertHeading(ByVal nHeading As E_Heading) As E_Heading
    Select Case nHeading
        Case E_Heading.east
            InvertHeading = west
        Case E_Heading.west
            InvertHeading = east
        Case E_Heading.south
            InvertHeading = north
        Case E_Heading.north
            InvertHeading = south
    End Select
End Function
Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave(SND_FUEGO, location.x, location.y, LoopStyle.Enabled)
    End If
End Sub

Public Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .x > UserPos.x - 11 And .x < UserPos.x + 111 And .y > UserPos.y - 9 And .y < UserPos.y + 9
    End With
End Function

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim x As Integer
    Dim y As Integer
    Dim AddX As Integer
    Dim AddY As Integer
    Dim nHeading As E_Heading
    Dim tmpInt As Integer
    Dim hayColision As Boolean
    
    With charlist(CharIndex)
        x = .Pos.x
        y = .Pos.y
        
        If x > 0 And y > 0 Then
                
        AddX = nX - x
        AddY = nY - y
        
        If Sgn(AddX) = 1 Then
            nHeading = E_Heading.east
        ElseIf Sgn(AddX) = -1 Then
            nHeading = E_Heading.west
        ElseIf Sgn(AddY) = -1 Then
            nHeading = E_Heading.north
        ElseIf Sgn(AddY) = 1 Then
            nHeading = E_Heading.south
        End If
        
        If MapData(nX, nY).CharIndex > 0 Then
            tmpInt = MapData(nX, nY).CharIndex
            
            'Si está muerto lo pisamos
            If charlist(tmpInt).muerto = False Then
                'Si pisó el PJ volvemos a su posición anterior
                If MapData(nX, nY).CharIndex = UserCharIndex Then
                    Debug.Print "************************************************************ COLISIÓN ********************************************************************************"
                    WAIT_ACTION = eWAIT_FOR_ACTION.RPU
                    Call WriteRequestPositionUpdate
                End If
                tmpInt = 0
            Else
                charlist(tmpInt).Pos.x = x
                charlist(tmpInt).Pos.y = y
                charlist(tmpInt).Heading = InvertHeading(nHeading)
                charlist(tmpInt).MoveOffsetX = 1 * (TilePixelWidth * AddX)
                charlist(tmpInt).MoveOffsetY = 1 * (TilePixelHeight * AddY)
                
                charlist(tmpInt).Moving = 1
                
                charlist(tmpInt).scrollDirectionX = -AddX
                charlist(tmpInt).scrollDirectionY = -AddY
                
                'Si el fantasma soy yo mueve la pantalla
                If tmpInt = UserCharIndex Then Call MoveScreen(charlist(tmpInt).Heading)
            End If
        Else
            tmpInt = 0
        End If
        
        MapData(x, y).CharIndex = tmpInt
              
        
        MapData(nX, nY).CharIndex = CharIndex
        
       
'
'         If hayColision = True Then
'            'Call EraseChar(CharIndex)
'            'charlist(UserCharIndex).Pos.x = .LastPos.x
'            'charlist(UserCharIndex).Pos.y = .LastPos.y
'            'Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY, True, nX, nY)
'            'Call CharRender(charlist(UserCharIndex), UserCharIndex, charlist(UserCharIndex).Pos.X, charlist(UserCharIndex).Pos.Y)
'        End If
        
        
        .Pos.x = nX
        .Pos.y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * AddX)
        .MoveOffsetY = -1 * (TilePixelHeight * AddY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(AddX)
        .scrollDirectionY = Sgn(AddY)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
            .FxIndex = 0
        End If
        End If
        
        If Not EstaPCarea(CharIndex) Then
            Call Dialogos.RemoveDialog(CharIndex)
        Else
            If .muerto = False Then
                Call vPasos.CreatePasos(x, y, DamePasos(nHeading))
            End If
        End If
        
    End With
    
    
    
'    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
'        Call EraseChar(CharIndex)
'    End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim x As Integer
    Dim y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.north
            y = -1
        
        Case E_Heading.east
            x = 1
        
        Case E_Heading.south
            y = 1
        
        Case E_Heading.west
            x = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.y + y
    
    'Check to see if its out of bounds
    If tX < 1 Or tX > MapInfo.Width Or tY < 1 Or tY > MapInfo.Height Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.y = y
        UserPos.y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 7 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
    End If
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim J As Long
    Dim k As Long
    
    For J = UserPos.x - 8 To UserPos.x + 8
        For k = UserPos.y - 6 To UserPos.y + 6
            If InMapBounds(J, k) Then
                If MapData(J, k).ObjGrh.GrhIndex = GrhFogata Then
                    
                    location.x = J
                    location.y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next J
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim loopc As Long
    Dim Dale As Boolean
    
    loopc = 1
    Do While charlist(loopc).ACTIVE And Dale
        loopc = loopc + 1
        Dale = (loopc <= UBound(charlist))
    Loop
    
    NextOpenChar = loopc
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhData() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    Open IniPath & GraphicsFile For Binary Access Read As handle
    Seek #1, 1
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
        If Grh > 0 Then
        With GrhData(Grh)
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next Frame
                
                Get handle, , .Speed
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .PixelHeight = GrhData(.Frames(1)).PixelHeight
                'If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .PixelWidth = GrhData(.Frames(1)).PixelWidth
                'If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                'If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                'If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                
                Get handle, , .PixelWidth
                If .PixelWidth <= 0 Then GoTo ErrorHandler
                
                Get handle, , .PixelHeight
                If .PixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .PixelWidth / TilePixelHeight
                .TileHeight = .PixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
        End If
    Wend
    
    Close handle
    
    LoadGrhData = True
Exit Function

ErrorHandler:
    LoadGrhData = False
End Function

Function LegalPos(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If x < 1 Or x > MapInfo.Width Or y < 1 Or y > MapInfo.Height Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(x, y).Blocked = 1 Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(x, y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> HayAgua(x, y) Then
        Exit Function
    End If
    
    LegalPos = True
End Function

Function MoveToLegalPos(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 10/05/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If x < 1 Or x > MapInfo.Width Or y < 1 Or y > MapInfo.Height Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(x, y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(x, y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.x, UserPos.y).Blocked = 1 Then
            Exit Function
        End If
        
        With charlist(CharIndex)
            ' Si no es casper, no puede pasar
            If .muerto = False Or .nombre = "" Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.x, UserPos.y) Then
                    If Not HayAgua(x, y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(x, y) Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> HayAgua(x, y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
End Function

Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If x < 1 Or x > MapInfo.Width Or y < 1 Or y > MapInfo.Height Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Sub DrawGrhIndexLuz(ByVal GrhIndex As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByRef color() As Long)

    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        Call Engine_Render_Rectangle(x, y, .PixelWidth, .PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , , .FileNum, color(0), color(1), color(2), color(3))
    End With
End Sub

Sub DrawGrhIndex(ByVal GrhIndex As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal color As Long)

    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        Call Engine_Render_Rectangle(x, y, .PixelWidth, .PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , , .FileNum, color, color, color, color)
    End With
End Sub
Sub DrawGrhLuz(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Single, ByRef color() As Long)
    Dim CurrentGrhIndex As Integer
    
On Error GoTo Error

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * Animate * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    If Grh.GrhIndex > 0 Then
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        'If COLOR = -1 Then COLOR = Iluminacion

        Call Engine_Render_Rectangle(x, y, .PixelWidth, .PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , 0, .FileNum, color(0), color(1), color(2), color(3))
    End With
    End If
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub
Sub DrawGrhShadow(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Single, Optional Shadow As Byte = 0, Optional color As Long = -1, Optional ShadowAlpha As Single = 255, Optional Chiquitolin As Boolean = False)
    Dim CurrentGrhIndex As Integer
    
On Error GoTo Error

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * Animate * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        Dim PixelWidth As Integer
        Dim PixelHeight As Integer
        
         ' <<<< CHIQUITOLIN >>>>
        If Chiquitolin = True Then
            PixelWidth = PixelWidth * 0.7
            PixelHeight = PixelHeight * 0.7
        Else
            PixelWidth = .PixelWidth
            PixelHeight = .PixelHeight
        End If
                
        'Draw
        'Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
      
        If mOpciones.Shadows = True And Chiquitolin = False And Conectar = False Then
            If Shadow = 1 Then
                ShadowColor = D3DColorRGBA(0, 0, 0, ShadowAlpha * 100 / 255)
                Call Engine_Render_Rectangle(x, y, PixelWidth, PixelHeight, .sX, .sY, PixelWidth, PixelHeight, , , 0, .FileNum, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1, False)
            ElseIf Shadow = 2 Then
                ShadowColor = D3DColorRGBA(0, 0, 0, ShadowAlpha * 100 / 255)
                Call Engine_Render_Rectangle(x + 10, y - 16, .PixelWidth, PixelHeight, .sX, .sY, PixelWidth, PixelHeight, , , 0, .FileNum, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1, False)
            End If
        End If
      
        If color = -1 Then color = Iluminacion
    End With
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub
Sub DrawGrhShadowOff(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Single, Optional color As Long = -1, Optional Chiquitolin As Boolean = False)
    Dim CurrentGrhIndex As Integer
    
On Error GoTo Error
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        If color = -1 Then color = Iluminacion
        Dim PixelWidth As Integer
        Dim PixelHeight As Integer
        
        ' <<<< CHIQUITOLIN >>>>
        If Chiquitolin = True Then
            PixelWidth = .PixelWidth * 0.7
            PixelHeight = .PixelHeight * 0.7
        Else
            PixelWidth = .PixelWidth
            PixelHeight = .PixelHeight
        End If

        Call Engine_Render_Rectangle(x, y, PixelWidth, PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , 0, .FileNum, color, color, color, color)
    End With
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

Sub DrawGrh(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Single, Optional Shadow As Byte = 0, Optional color As Long = -1, Optional ShadowAlpha As Single = 255)
    Dim CurrentGrhIndex As Integer
    
On Error GoTo Error

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * Animate * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                        If Grh.Loops = 0 Then
                            Grh.Started = 0
                        End If
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                x = x - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                y = y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
                
        'Draw
        'Call BackBufferSurface.BltFast(X, Y, SurfaceDB.Surface(.FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        If Shadow = 1 Then
            ShadowColor = D3DColorRGBA(0, 0, 0, ShadowAlpha * 100 / 255)
            Call Engine_Render_Rectangle(x, y, .PixelWidth, .PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , 0, .FileNum, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1, False)
        ElseIf Shadow = 2 Then
            ShadowColor = D3DColorRGBA(0, 0, 0, ShadowAlpha * 100 / 255)
            Call Engine_Render_Rectangle(x + 10, y - 16, .PixelWidth, .PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , 0, .FileNum, ShadowColor, ShadowColor, ShadowColor, ShadowColor, 1, False)
        End If
        If color = -1 Then color = Iluminacion

        Call Engine_Render_Rectangle(x, y, .PixelWidth, .PixelHeight, .sX, .sY, .PixelWidth, .PixelHeight, , , 0, .FileNum, color, color, color, color)
    End With
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

Public Sub RenderNiebla()

If Zonas(ZonaActual).Niebla = 0 Then
    If nAlpha > 0 Then
        If GTCPres < (GetTickCount() And &H7FFFFFFF) - GTCInicial Then
            nAlpha = nAlpha - 1
            GTCPres = (GetTickCount() And &H7FFFFFFF)
        End If
    Else
        Exit Sub
    End If
End If

If nAlpha < Zonas(ZonaActual).Niebla Then
    If GTCPres < (GetTickCount() And &H7FFFFFFF) - GTCInicial Then
        nAlpha = nAlpha + IIf(nAlpha + 1 < Zonas(ZonaActual).Niebla, 1, Zonas(ZonaActual).Niebla - nAlpha)
        GTCPres = (GetTickCount() And &H7FFFFFFF)
    End If
End If

Dim Mueve As Single 'Niebla
Dim T As Single
Dim color As Long

GTCPres = Abs((GetTickCount() And &H7FFFFFFF) - GTCInicial)
T = (GTCPres - 4000) / 1000
Mueve = (T * 20) Mod 512

color = D3DColorRGBA(Zonas(ZonaActual).NieblaR, Zonas(ZonaActual).NieblaG, Zonas(ZonaActual).NieblaB, nAlpha)

Call Engine_Render_D3DXSprite(255, 255, 512 - Mueve, 512, Mueve, 0, color, 14706, 0)
Call Engine_Render_D3DXSprite(255, 767, 512 - Mueve, 256, Mueve, 0, color, 14706, 0)

Call Engine_Render_D3DXSprite(767 - Mueve, 255, 512, 512, 0, 0, color, 14706, 0)
Call Engine_Render_D3DXSprite(767 - Mueve, 767, 512, 256, 0, 0, color, 14706, 0)

Call Engine_Render_D3DXSprite(1279 - Mueve, 255, Mueve, 512, 0, 0, color, 14706, 0)
Call Engine_Render_D3DXSprite(1279 - Mueve, 767, Mueve, 256, 0, 0, color, 14706, 0)

End Sub


Function GetBitmapDimensions(ByVal BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
    Dim BMHeader As BITMAPFILEHEADER
    Dim BINFOHeader As BITMAPINFOHEADER
    
    Open BmpFile For Binary Access Read As #1
    
    Get #1, , BMHeader
    Get #1, , BINFOHeader
    
    Close #1
    
    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight
End Function

Sub DrawGrhtoHdc(ByVal hdc As Long, ByVal GrhIndex As Integer, ByRef SourceRect As RECT, ByRef destRect As RECT)
'*****************************************************************
'Draws a Grh's portion to the given area of any Device Context
'*****************************************************************
    'Call SurfaceDB.Surface(GrhData(GrhIndex).FileNum).BltToDC(hDC, SourceRect, destRect)
Call TransparentBlt(hdc, 0, 0, 32, 32, Inventario.Grafico(GrhData(GrhIndex).FileNum), 0, 0, 32, 32, vbMagenta)
End Sub
Public Sub CargarTile(x As Long, y As Long, ByRef DataMap() As Byte)
Dim ByFlags As Byte
Dim Rango As Byte
Dim i As Integer
Dim tmpInt As Integer

Dim Pos As Long

Pos = MapInfo.Offset + (x - 1) * 10 + (y - 1) * MapInfo.Width * 10

ByFlags = DataMap(Pos)
ByFlags = ByFlags Xor ((x Mod 200) + 55)
Pos = Pos + 1

If ByFlags = 50 Then
    MapData(x, y).Blocked = 1
Else
    MapData(x, y).Blocked = 0
End If
MapData(x, y).Trigger = ByFlags

For i = 1 To 4
    tmpInt = (DataMap(Pos + 1) And &H7F) * &H100 Or DataMap(Pos) Or -(DataMap(Pos + 1) > &H7F) * &H8000
    Pos = Pos + 2
    Select Case i
        Case 1
            MapData(x, y).Graphic(1).GrhIndex = (tmpInt Xor (y + 301) Xor (x + 721)) - x
        Case 2
            MapData(x, y).Graphic(2).GrhIndex = (tmpInt Xor (y + 501) Xor (x + 529)) - x
        Case 3
            MapData(x, y).Graphic(3).GrhIndex = (tmpInt Xor (x + 239) Xor (y + 319)) - x
        Case 4
            MapData(x, y).Graphic(4).GrhIndex = (tmpInt Xor (x + 671) Xor (y + 129)) - x
    End Select
    
    If MapData(x, y).Graphic(i).GrhIndex > 0 Then
        InitGrh MapData(x, y).Graphic(i), MapData(x, y).Graphic(i).GrhIndex
    End If
Next i
'Get ArchivoMapa, , Rango
Rango = DataMap(Pos)
Pos = Pos + 1

MapData(x, y).Map = UserMap

MapData(x, y).Light_Value(0) = D3DColorRGBA(255, 255, 255, 255)
MapData(x, y).Light_Value(1) = D3DColorRGBA(255, 255, 255, 255)
MapData(x, y).Light_Value(2) = D3DColorRGBA(255, 255, 255, 255)
MapData(x, y).Light_Value(3) = D3DColorRGBA(255, 255, 255, 255)
MapData(x, y).Hora = 99

Call Light_Destroy_ToMap(x, y)

If MapData(x, y).Graphic(3).GrhIndex < 0 Then
    Call Light_Create(x, y, 255, 255, 255, Rango, -MapData(x, y).Graphic(3).GrhIndex - 1)
End If
End Sub
Sub RenderScreen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffSetX As Single, ByVal PixelOffSetY As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim y           As Long     'Keeps track of where on map we are
    Dim x           As Long     'Keeps track of where on map we are
    Dim screenminY  As Integer  'Start Y pos on current screen
    Dim screenmaxY  As Integer  'End Y pos on current screen
    Dim screenminX  As Integer  'Start X pos on current screen
    Dim screenmaxX  As Integer  'End X pos on current screen
    Dim MinY        As Integer  'Start Y pos on current map
    Dim MaxY        As Integer  'End Y pos on current map
    Dim MinX        As Integer  'Start X pos on current map
    Dim MaxX        As Integer  'End X pos on current map
    Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    Dim PixelOffSetXTemp As Integer 'For centering grhs
    Dim PixelOffSetYTemp As Integer 'For centering grhs
    Dim tmpInt As Integer
    Dim tmpLong As Long
    Dim SupIndex As Integer
    Dim ByFlags As Byte
    Dim i As Integer
    Dim color As Long
        
    Dim Eliminados As Integer
    Dim Cant As Integer
    
    If UserMap = 0 Then Exit Sub
    
    'Figure out Ends and Starts of screen
    screenminY = TileY - HalfWindowTileHeight
    screenmaxY = TileY + HalfWindowTileHeight
    screenminX = TileX - HalfWindowTileWidth
    screenmaxX = TileX + HalfWindowTileWidth
    
    MinY = screenminY - TileBufferSize
    MaxY = screenmaxY + TileBufferSize
    MinX = screenminX - TileBufferSize
    MaxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If MinY < 1 Then
        minYOffset = 1 - MinY
        MinY = 1
    End If
    
    If MaxY > MapInfo.Height Then MaxY = MapInfo.Height
    
    If MinX < 1 Then
        minXOffset = 1 - MinX
        MinX = 1
    End If
    
    If MaxX > MapInfo.Width Then MaxX = MapInfo.Width
    
    'If we can, we render around the view area to make it smoother
    
    screenminY = screenminY - 1
    
    
    If screenmaxY < MapInfo.Height Then
        screenmaxY = screenmaxY + 1
    Else
        screenmaxY = MapInfo.Height
    End If
    
    screenminX = screenminX - 1

    
    If screenmaxX < MapInfo.Width Then
        screenmaxX = screenmaxX + 1
    Else
        screenmaxX = MapInfo.Width
    End If
    
        
    'Dim CambioHora As Boolean
    'Cargar mapa
    For y = MinY - 6 To MaxY + 6
        For x = MinX - 6 To MaxX + 6
            If x > 0 And y > 0 And x <= MapInfo.Width And y <= MapInfo.Height Then
                If MapData(x, y).Map <> UserMap Then
                    If UserMap = 1 Then
                        Call CargarTile(x, y, DataMap1)
                    Else
                        Call CargarTile(x, y, DataMap2)
                    End If
                End If
                If MapData(x, y).Hora <> Hora And UserMap = 1 Then
                    For i = 0 To 3
                        MapData(x, y).Light_Value(i) = Iluminacion
                    Next i
                    MapData(x, y).Hora = Hora
                    'CambioHora = True
                End If
            End If
        Next x
    Next y
        
    Light_Render_Area
    
    'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            If x > 0 And y > 0 And x <= MapInfo.Width And y <= MapInfo.Height Then
            'Layer 1 **********************************
            Call DrawGrhLuz(MapData(x, y).Graphic(1), _
                (ScreenX - 1) * TilePixelWidth + PixelOffSetX + TileBufferPixelOffsetX, _
                (ScreenY - 1) * TilePixelHeight + PixelOffSetY + TileBufferPixelOffsetY, _
                 0, 1, MapData(x, y).Light_Value)
            '******************************************
            End If
            ScreenX = ScreenX + 1
            
        Next x
        
        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next y
    
    'Draw floor layer 2
    ScreenY = minYOffset
    For y = MinY To MaxY
        ScreenX = minXOffset
        For x = MinX To MaxX
            If x > 0 And y > 0 And x <= MapInfo.Width And y <= MapInfo.Height Then
            'Layer 2 **********************************
            If MapData(x, y).Graphic(2).GrhIndex <> 0 Then
                Call DrawGrhLuz(MapData(x, y).Graphic(2), _
                        (ScreenX - 1) * TilePixelWidth + PixelOffSetX, _
                        (ScreenY - 1) * TilePixelHeight + PixelOffSetY, _
                        1, 1, MapData(x, y).Light_Value)
            End If
            End If
            '******************************************
            
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    
    Dim mNPCMuerto As clsNPCMuerto
    
    Eliminados = 0
    Cant = NPCMuertos.Count
    For i = 1 To Cant
        Set mNPCMuerto = NPCMuertos(i - Eliminados)
        Call mNPCMuerto.Update '(TileX, TileY, PixelOffSetX, PixelOffSetY)
        If mNPCMuerto.KillMe Then
            NPCMuertos.Remove (i - Eliminados)
            Eliminados = Eliminados + 1
        End If
    Next i
    
    
    'Draw Transparent Layers
    ScreenY = minYOffset
    For y = MinY To MaxY
        ScreenX = minXOffset
        For x = MinX To MaxX
            PixelOffSetXTemp = (ScreenX - 1) * TilePixelWidth + PixelOffSetX
            PixelOffSetYTemp = (ScreenY - 1) * TilePixelHeight + PixelOffSetY
            
            With MapData(x, y)
                'Pasos
                 If .PasosIndex <> 0 Then Call vPasos.RenderPasos(PixelOffSetXTemp, PixelOffSetYTemp, .PasosIndex)
                             
            
                'Object Layer **********************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Call DrawGrhLuz(.ObjGrh, _
                            PixelOffSetXTemp, PixelOffSetYTemp, 1, 1, MapData(x, y).Light_Value)

                End If
                '***********************************************
                
                
                If Not .Elemento Is Nothing Then 'Render de Npc Muertos
                    Call .Elemento.Render(PixelOffSetXTemp, PixelOffSetYTemp)
                End If
            
                'Char layer ************************************
                If .CharIndex <> 0 Then
                    Call CharRender(charlist(.CharIndex), .CharIndex, PixelOffSetXTemp, PixelOffSetYTemp)
                    If .CharIndex <> UserCharIndex And UserPos.x = charlist(.CharIndex).Pos.x And UserPos.y = charlist(.CharIndex).Pos.y Then
                         Debug.Print "ME PISO CHEEE ******************************************************************************"
                    End If
                End If
                
                If UserMap = 1 Then
                    Call RenderBarcos(x, y, TileX, TileY, PixelOffSetX, PixelOffSetY)
                End If
                
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex > 0 Then
                    'Draw
                    SupIndex = GrhData(.Graphic(3).GrhIndex).FileNum
                    If ((SupIndex >= 7000 And SupIndex <= 7008) Or (SupIndex >= 1261 And SupIndex <= 1287) Or SupIndex = 648 Or SupIndex = 645) Then
                        If mOpciones.TransparencyTree = True And UserPos.x >= x - 3 And UserPos.x <= x + 3 And UserPos.y >= y - 5 And UserPos.y <= y Then
                            Call DrawGrh(.Graphic(3), PixelOffSetXTemp, PixelOffSetYTemp, 1, 1, 0, D3DColorRGBA(IluRGB.r, IluRGB.G, IluRGB.B, 180))
                        Else
                            Call DrawGrhLuz(.Graphic(3), PixelOffSetXTemp, PixelOffSetYTemp, 1, 1, MapData(x, y).Light_Value)
                        End If
                    Else
                       Call DrawGrh(.Graphic(3), PixelOffSetXTemp, PixelOffSetYTemp, 1, 1)
                    End If
                End If
                '*************************************************
                
                
                'Layer 3 Plus FX *****************************************
                If .fXGrh.Started = 1 Then
                    Call DrawGrh(.fXGrh, PixelOffSetXTemp - FxData(.fX).OFFSETX, PixelOffSetYTemp - FxData(.fX).OFFSETY, 1, 1)
                    If .fXGrh.Started = 0 Then .fX = 0
                End If
                '************************************************
                
            End With
            
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    
    
    Dim mArroja As clsArroja
    Dim Elemento
    For Each Elemento In Arrojas
        Set mArroja = Elemento
            Call mArroja.Render(TileX, TileY, PixelOffSetX, PixelOffSetY)
    Next Elemento
    
    Dim mTooltip As clsToolTip
    
    Eliminados = 0
    Cant = Tooltips.Count
    For i = 1 To Cant
        Set mTooltip = Tooltips(i - Eliminados)
        Call mTooltip.Render(TileX - 1, TileY, PixelOffSetX, PixelOffSetY)
        If mTooltip.Alpha = 0 Then
            Tooltips.Remove (i - Eliminados)
            Eliminados = Eliminados + 1
        End If
    Next i
            
    If Not bTecho Then
        If bAlpha < 255 Then
            If tTick < (GetTickCount() And &H7FFFFFFF) - 30 Then
                bAlpha = bAlpha + IIf(bAlpha + 8 < 255, 8, 255 - bAlpha)
                ColorTecho = D3DColorRGBA(IluRGB.r, IluRGB.G, IluRGB.B, bAlpha)
                tTick = (GetTickCount() And &H7FFFFFFF)
            End If
        End If
    Else
        If bAlpha > 0 Then
            If tTick < (GetTickCount() And &H7FFFFFFF) - 30 Then
                bAlpha = bAlpha - IIf(bAlpha - 8 > 0, 8, bAlpha)
                ColorTecho = D3DColorRGBA(IluRGB.r, IluRGB.G, IluRGB.B, bAlpha)
                tTick = (GetTickCount() And &H7FFFFFFF)
            End If
        End If
    End If
    
    'Draw blocked tiles and grid
    ScreenY = minYOffset
    For y = MinY To MaxY
        ScreenX = minXOffset
        For x = MinX To MaxX
            'Layer 4 **********************************
            If MapData(x, y).Graphic(4).GrhIndex And bAlpha > 0 Then
                'Draw
                Call DrawGrhIndex(MapData(x, y).Graphic(4).GrhIndex, _
                    (ScreenX - 1) * TilePixelWidth + PixelOffSetX, _
                    (ScreenY - 1) * TilePixelHeight + PixelOffSetY, _
                    1, ColorTecho)
            End If
                '**********************************
                
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    'TODO : Check this!!
    Dim ColorLluvia As Long
    If ZonaActual > 0 Then
        If Zonas(ZonaActual).Terreno <> eTerreno.Dungeon Then
            If bRain Then
                'Figure out what frame to draw
                If llTick < (GetTickCount() And &H7FFFFFFF) - 50 Then
                    iFrameIndex = iFrameIndex + 1
                    If iFrameIndex > 7 Then iFrameIndex = 0
                    llTick = (GetTickCount() And &H7FFFFFFF)
                End If
                ColorLluvia = D3DColorRGBA(IluRGB.r, IluRGB.G, IluRGB.B, 140)
                For y = 0 To 5
                    For x = 0 To 6
                            Call Engine_Render_Rectangle(LTLluvia(x), LTLluvia(y), RLluvia(iFrameIndex).Right - RLluvia(iFrameIndex).Left, RLluvia(iFrameIndex).Bottom - RLluvia(iFrameIndex).Top, RLluvia(iFrameIndex).Left, RLluvia(iFrameIndex).Top, RLluvia(iFrameIndex).Right - RLluvia(iFrameIndex).Left, RLluvia(iFrameIndex).Bottom - RLluvia(iFrameIndex).Top, , , , 5556, ColorLluvia, ColorLluvia, ColorLluvia, ColorLluvia)
                    Next x
                Next y
            End If
        End If
    End If
    
    Call Dialogos.Render
    Call DibujarCartel
    Call DialogosClanes.Draw
    
    If CambioZona > 0 And ZonaActual > 0 Then
        If CambioZona > 300 Then
            tmpInt = 500 - CambioZona
        ElseIf CambioZona < 200 Then
            tmpInt = CambioZona
        Else
            tmpInt = 200
        End If
        If zTick < (GetTickCount() And &H7FFFFFFF) - 50 Then
            CambioZona = CambioZona - 5
            zTick = (GetTickCount() And &H7FFFFFFF)
        End If

        'Mensaje al cambiar de zona
        Call D3DX.DrawText(MainFont, D3DColorRGBA(0, 0, 0, tmpInt), Zonas(ZonaActual).nombre, DDRect(0, 10, 814, 220), DT_CENTER)
        Call D3DX.DrawText(MainFont, D3DColorRGBA(220, 215, 215, tmpInt), Zonas(ZonaActual).nombre, DDRect(5, 15, 814, 220), DT_CENTER)

        If CambioSegura Then
            Call DrawFont(IIf(Zonas(ZonaActual).Segura = 1, "Entraste a una zona segura", "Saliste de una zona segura"), 574, 345, D3DColorRGBA(255, 0, 0, tmpInt))
        End If
    End If
    
    
        If UseMotionBlur And mOpciones.BlurEffects = True Then
        
            AngMareoMuerto = AngMareoMuerto + timerElapsedTime * 0.002
            If AngMareoMuerto >= 6.28318530717959 Then
                AngMareoMuerto = 0
                'GoingHome = 0
            End If
            
            If GoingHome = 1 Then
                RadioMareoMuerto = RadioMareoMuerto + timerElapsedTime * 0.01
                If RadioMareoMuerto > 50 Then RadioMareoMuerto = 50
            ElseIf GoingHome = 2 Then
                RadioMareoMuerto = RadioMareoMuerto - timerElapsedTime * 0.02
                If RadioMareoMuerto <= 0 Then
                    RadioMareoMuerto = 0
                    GoingHome = 0
                End If
            End If
        
            If FrameUseMotionBlur Then
                FrameUseMotionBlur = False
                With D3DDevice
               
                    'Dim ValueEffect As Long
                    'ValueEffect = 2048
                   
                    'Perform the zooming calculations
                    ' * 1.333... maintains the aspect ratio
                    ' ... / 1024 is to factor in the buffer size
                    BlurTA(0).tu = ZoomLevel + RadioMareoMuerto / 2048 * Sin(AngMareoMuerto) + RadioMareoMuerto / 2048
                    BlurTA(0).tv = ZoomLevel + RadioMareoMuerto / 2048 * Cos(AngMareoMuerto) + RadioMareoMuerto / 2048
                    BlurTA(1).tu = ((ScreenWidth + 1 + Cos(AngMareoMuerto) * RadioMareoMuerto / 2 - RadioMareoMuerto / 2) / 800) - ZoomLevel
                    BlurTA(1).tv = ZoomLevel + RadioMareoMuerto / 2048 * Sin(AngMareoMuerto) + RadioMareoMuerto / 2048
                    BlurTA(2).tu = ZoomLevel + RadioMareoMuerto / 2048 * Cos(AngMareoMuerto) + RadioMareoMuerto / 2048
                    BlurTA(2).tv = ((ScreenHeight + 1 + Sin(AngMareoMuerto) * RadioMareoMuerto / 2 - RadioMareoMuerto / 2) / 608) - ZoomLevel
                    BlurTA(3).tu = BlurTA(1).tu
                    BlurTA(3).tv = BlurTA(2).tv
                   
                    'Draw what we have drawn thus far since the last .Clear
                    'LastTexture = -100
                    D3DDevice.EndScene
                    .SetRenderTarget pBackbuffer, Nothing, ByVal 0
                    
                    D3DDevice.BeginScene
                
                    .SetTexture 0, BlurTexture
                    .SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(BlurIntensity, 255, 255, 255)
                    .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TFACTOR
                    .DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, BlurTA(0), Len(BlurTA(0))
                    .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
               
                End With
            End If
        End If
    
            
    If VerMapa Then
        
        '420
        '0.46545454545454545454545454545455
        
        If UserMap = 1 Then
            PosMapX = -Int(UserPos.x * RelacionMiniMapa) + 32 + 398
            PosMapY = -Int(UserPos.y * RelacionMiniMapa) + 32 + 292
            
            If PosMapX > 0 Then PosMapX = 0
            If PosMapX < -1247 Then PosMapX = -1247
            If PosMapY > 0 Then PosMapY = 0
            If PosMapY < -2210 Then PosMapY = -2210
            
          
            
            color = D3DColorRGBA(255, 255, 255, 225)
            
            If PosMapX > -1024 Then 'Dibujo primera columna
                If PosMapY <= 0 And PosMapY > -1024 Then
                    Call Engine_Render_Rectangle(256, 256, _
                                                 800 + IIf(PosMapX < -288, PosMapX + 288, 0), 608 + IIf(PosMapY < -480, PosMapY + 480, 0), _
                                                 -PosMapX, -PosMapY, _
                                                 800 + IIf(PosMapX < -288, PosMapX + 288, 0), 608 + IIf(PosMapY < -480, PosMapY + 480, 0), , , , 14763, color, color, color, color)
                End If
                If PosMapY <= -480 And PosMapY > -2048 Then
                    Call Engine_Render_Rectangle(256, 256 + PosMapY + 1024, _
                                                 800 + IIf(PosMapX < -288, PosMapX + 288, 0), 608 + IIf(PosMapY + 1024 > 0, PosMapY + 1024, 0) + 480, _
                                                 -PosMapX, 0, _
                                                 800 + IIf(PosMapX < -288, PosMapX + 288, 0), 608 + IIf(PosMapY + 1024 > 0, PosMapY + 1024, 0) + 480, , , , 14765, color, color, color, color)
                End If
                If PosMapY <= -1504 Then
                    Call Engine_Render_Rectangle(256, 256 + PosMapY + 2048, _
                                                 800 + IIf(PosMapX < -288, PosMapX + 288, 0), 608 + IIf(PosMapY + 2048 > -480, PosMapY + 2048 + 480, 0), _
                                                 -PosMapX, 0, _
                                                 800 + IIf(PosMapX < -288, PosMapX + 288, 0), 608 + IIf(PosMapY + 2048 > -480, PosMapY + 2048 + 480, 0), , , , 14767, color, color, color, color)
                End If
                            
            End If
            
            If PosMapX < -288 Then 'Dibujo segunda columna
                If PosMapY <= 0 And PosMapY > -1024 Then
                    Call Engine_Render_Rectangle(256 + IIf(PosMapX > -1024, 736 + 288 + PosMapX, 0), 256, _
                                                 800 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 608 + IIf(PosMapY < -480, PosMapY + 480, 0), _
                                                 IIf(PosMapX < -1024, -PosMapX - 1024, 0), -PosMapY, _
                                                 800 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 608 + IIf(PosMapY < -480, PosMapY + 480, 0), , , , 14764, color, color, color, color)
                End If
                If PosMapY <= -480 And PosMapY > -2048 Then
                    Call Engine_Render_Rectangle(256 + IIf(PosMapX > -1024, 736 + 288 + PosMapX, 0), 256 + PosMapY + 1024, _
                                                 800 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 608 + IIf(PosMapY + 1024 > 0, PosMapY + 1024, 0) + 480, _
                                                 IIf(PosMapX < -1024, -PosMapX - 1024, 0), 0, _
                                                 800 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 608 + IIf(PosMapY + 1024 > 0, PosMapY + 1024, 0) + 480, , , , 14766, color, color, color, color)
                End If
                If PosMapY <= -1504 Then
                    Call Engine_Render_Rectangle(256 + IIf(PosMapX > -1024, 736 + 288 + PosMapX, 0), 256 + PosMapY + 2048, _
                                                 800 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 608 + IIf(PosMapY + 2048 > -480, PosMapY + 2048 + 480, 0), _
                                                 IIf(PosMapX < -1024, -PosMapX - 1024, 0), 0, _
                                                 800 + IIf(PosMapX < -1024, 0, -1024 - PosMapX), 608 + IIf(PosMapY + 2048 > -480, PosMapY + 2048 + 480, 0), , , , 14768, color, color, color, color)
                End If
            End If
            
            
            'If PosMap <= 210 Then '488
            '    MapaY = 0
            'ElseIf PosMap > 210 And PosMap < 492 Then
            '    MapaY = -(PosMap - 210)
            'Else
            '    MapaY = -282
            'End If
           
            
            'Call Engine_Render_Rectangle(256 + 0, 256 + MapaY, 512, 512, 0, 0, 512, 512, , , , 14404, color, color, color, color)
            'Call Engine_Render_Rectangle(256 + 0, 256 + 512 + MapaY, 512, 186, 0, 0, 512, 186, , , , 14405, color, color, color, color)
            color = D3DColorRGBA(255, 255, 255, 255)
            'Call Engine_Render_Rectangle(256 + UserPos.x * RelacionMiniMapa - 35 + PosMapX, 256 + UserPos.Y * RelacionMiniMapa - 35 + PosMapY, 5, 5, 0, 0, 5, 5, , , , 1, color, color, color, color)
            Call Engine_Render_Rectangle(256 + UserPos.x * RelacionMiniMapa - 35 + PosMapX, 256 + UserPos.y * RelacionMiniMapa - 35 + PosMapY, 5, 5, 0, 0, 5, 5, , , , 1, color, color, color, color)
           
           
            x = Int((frmMain.MouseX - PosMapX + 32) / RelacionMiniMapa)
            y = Int((frmMain.MouseY - PosMapY + 32) / RelacionMiniMapa)
            
            If x > 1 And x < 1100 And y > 1 And y < 1500 Then
                Call DrawFont("(" & x & "," & y & ")", frmMain.MouseX + 266, frmMain.MouseY + 266, D3DColorRGBA(255, 255, 255, 200))
                i = BuscarZona(x, y)
                If i > 0 Then
                    Call DrawFont(Zonas(i).nombre, frmMain.MouseX + 246, frmMain.MouseY + 266 + 13, D3DColorRGBA(255, 255, 255, 200))
                End If
            End If
        ElseIf ZonaActual = 33 Or ZonaActual = 34 Or ZonaActual = 35 Then 'Dungeon Newbie
            color = D3DColorRGBA(255, 255, 255, 190)
            Call Engine_Render_Rectangle(256 + 60, 256 + 3, 512, 512, 0, 0, 512, 512, , , , 14406, color, color, color, color)

            color = D3DColorRGBA(255, 255, 255, 255)
            Call Engine_Render_Rectangle(256 + 60 + (UserPos.x - 571) * 2.21105527638191, 256 + 5 + (UserPos.y - 311) * 2.21105527638191, 5, 5, 0, 0, 5, 5, , , , 1, color, color, color, color)
        Else
            'Mensaje al cambiar de zona
            Call D3DX.DrawText(MainFont, D3DColorRGBA(0, 0, 0, 200), Zonas(ZonaActual).nombre, DDRect(0, 10, 736, 200), DT_CENTER)
            Call D3DX.DrawText(MainFont, D3DColorRGBA(220, 215, 215, 200), Zonas(ZonaActual).nombre, DDRect(5, 15, 736, 200), DT_CENTER)
        End If
    End If
    
    
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    
     'Particle_Group_Render MapData(150, 800).particle_group_index, MouseX, MouseY
    
    
    'Dim tmplng As Long
    'Dim tmblng2 As Long
    'ScreenY = minYOffset '- TileBufferSize
    'For y = minY To maxY
    '    ScreenX = minXOffset '- TileBufferSize
    '    For x = minX To maxX
    '        With MapData(x, y)
    '            '*** Start particle effects ***
    '            If MapData(x, y).particle_group_index Then
    '                Particle_Group_Render MapData(x, y).particle_group_index, ScreenX, ScreenY
    '            End If
    '            '*** End particle effects ***
    '        End With
    '        ScreenX = ScreenX + 1
    '    Next x
    '    ScreenY = ScreenY + 1
    'Next y
'Call Engine_Render_Rectangle(frmMain.MouseX, frmMain.MouseY, 128, 128, 0, 256, 128, 128, , , 0, 14332)
                 
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
 

    If TiempoRetos > 0 Then
        '10 segundos de espera para empezar la ronda
        tmpLong = Abs((GetTickCount() And &H7FFFFFFF) - TiempoRetos)
        tmpInt = 10 - Int(tmpLong / 1000)
        
        If tmpLong < 10000 Then
            Call D3DX.DrawText(MainFontBig, D3DColorRGBA(0, 0, 0, 200), CStr(tmpInt), DDRect(0, 30, 736, 230), DT_CENTER)
            Call D3DX.DrawText(MainFontBig, D3DColorRGBA(220, 215, 215, 200), CStr(tmpInt), DDRect(5, 35, 736, 230), DT_CENTER)
        Else
            'Termino el tiempo de espera que empieze el reto
            TiempoRetos = 0
        End If
    End If
    
    
    If Entrada > 0 Then
        color = D3DColorRGBA(255, 255, 255, Entrada)
        Call Engine_Render_Rectangle(256, 256, 544, 416, 0, 0, 512, 416, , , , 14325, color, color, color, color)
        If zTick2 < (GetTickCount() And &H7FFFFFFF) - 75 Then
            Entrada = Entrada - 15
            zTick2 = (GetTickCount() And &H7FFFFFFF)
        End If
    End If
    
    If PergaminoDireccion > 0 Then
        If PergaminoTick < (GetTickCount() And &H7FFFFFFF) - 20 Then
            If PergaminoDireccion = 1 And AperturaPergamino < 240 Then
                AperturaPergamino = AperturaPergamino + (5 + Sqr(240 - AperturaPergamino) / 2)
                If AperturaPergamino > 240 Then AperturaPergamino = 240
            ElseIf PergaminoDireccion = 2 And AperturaPergamino > 0 Then
                AperturaPergamino = AperturaPergamino - (5 + Sqr(AperturaPergamino) / 2)
                If AperturaPergamino < 0 Then AperturaPergamino = 0
            End If
            PergaminoTick = (GetTickCount() And &H7FFFFFFF)
        End If

    End If
    
    If AperturaPergamino > 0 Then
        If DateDiff("s", TiempoAbierto, Now) > 10 Then
            PergaminoDireccion = 2
            TiempoAbierto = Now
        End If
        color = D3DColorRGBA(255, 255, 255, AperturaPergamino * 175 / 240)
                
        Call Engine_Render_Rectangle(256 + 10 - 5 + 240 - AperturaPergamino, 256 + 309 + 2, 28, 107, 0, 0, 28, 107, , , , 14687, color, color, color, color)
        Call Engine_Render_Rectangle(256 + 38 - 5 + 240 - AperturaPergamino, 256 + 336 + 2, AperturaPergamino, 74, 240 - AperturaPergamino, 108, AperturaPergamino, 74, , , , 14687, color, color, color, color)
        Call Engine_Render_Rectangle(256 + 517 - 5 - 240 + AperturaPergamino, 256 + 309 + 2, 26, 107, 29, 0, 26, 107, , , , 14687, color, color, color, color)
        Call Engine_Render_Rectangle(256 + 278 - 5, 256 + 335 + 2, AperturaPergamino, 74, 0, 182, AperturaPergamino, 74, , , , 14687, color, color, color, color)
    
        'If AperturaPergamino >= 232 Then
        '    Call Engine_Render_Rectangle(256 + 40, 256 + 335 + 11, 56, 56, 56, 0, 56, 56, , , , 14687, color, color, color, color)
        'ElseIf AperturaPergamino < 232 And AperturaPergamino >= 176 Then
        '    Call Engine_Render_Rectangle(256 + 40 + 232 - AperturaPergamino, 256 + 335 + 11, AperturaPergamino - 176, 56, 56 + 232 - AperturaPergamino, 0, AperturaPergamino - 176, 56, , , , 14687, color, color, color, color)
        'End If
        Call Engine_Render_D3DXTexture(256 + 38 - 5 + 240 - Int(AperturaPergamino), 256 + 342, Int(AperturaPergamino) * 2, 80, 240 - Int(AperturaPergamino), 0, color, pRenderTexture, 0)
    End If
         
    If mOpciones.Niebla = True Then
        Call RenderNiebla
    End If
    
    If AlphaCuenta > 0 Then
         Call RenderCuentaRegresiva
    End If
    
    If AlphaBlood > 0 Then
         Call RenderBlood
    End If
    
    If AlphaBloodUserDie > 0 Then
         Call RenderUserDieBlood
    End If
    
    If AlphaTextKills > 0 Then
         Call RenderTextKills
    End If
    
    If bRain = True Then
        Call RenderRelampago
    End If
    
    If UserCiego = True Then
         Call RenderCeguera
    End If
    
    If AlphaSalir > 0 Then
        Call RenderSaliendo
    End If

    If FPSFLAG Then Call DrawFont("FPS: " & FPS, 740, 260, D3DColorRGBA(255, 255, 255, 160))
End Sub

''*********************
'Sub UpdateRenderScreen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffSetX As Single, ByVal PixelOffSetY As Single)
'
'   '**************************************************************
''Author: Aaron Perkins
''Last Modify Date: 8/14/2007
''Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
''Renders everything to the viewport
''**************************************************************
'    Dim Y           As Long     'Keeps track of where on map we are
'    Dim X           As Long     'Keeps track of where on map we are
'    Dim minXOffset  As Integer
'    Dim minYOffset  As Integer
'    Dim PixelOffSetXTemp As Integer 'For centering grhs
'    Dim PixelOffSetYTemp As Integer 'For centering grhs
'    Dim tmpInt As Integer
'    Dim tmpLong As Long
'    Dim SupIndex As Integer
'    Dim i As Integer
'    Dim color As Long
'
'    Dim Eliminados As Integer
'    Dim Cant As Integer
'
'    If UserMap = 0 Then Exit Sub
'
'    'If MapData(TileX, TileY).Map <> UserMap Then
'    '    Call CargarTile(TileX, TileY)
'    'End If
'
'    If MapData(TileX, TileY).Hora <> Hora Then
'        For i = 0 To 3
'            MapData(TileX, TileY).Light_Value(i) = Iluminacion
'        Next i
'        MapData(TileX, TileY).Hora = Hora
'    End If
'
'
'    Light_Render_Area
'
'
'    'Layer 1 **********************************
'    Call DrawGrhLuz(MapData(TileX, TileY).Graphic(1), _
'       TilePixelWidth + PixelOffSetX + TileBufferPixelOffsetX, _
'       TilePixelHeight + PixelOffSetY + TileBufferPixelOffsetY, _
'         0, 1, MapData(TileX, TileY).Light_Value)
'    '******************************************
'
'
'
'    'Layer 2 **********************************
'    If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
'        Call DrawGrhLuz(MapData(TileX, TileY).Graphic(2), _
'                TilePixelWidth + PixelOffSetX, _
'                TilePixelHeight + PixelOffSetY, _
'                1, 1, MapData(TileX, TileY).Light_Value)
'    End If
'    '******************************************
'
'
'    'Draw Transparent Layers
'
'    PixelOffSetXTemp = TilePixelWidth + PixelOffSetX
'    PixelOffSetYTemp = TilePixelHeight + PixelOffSetY
'
'    With MapData(TileX, TileY)
'
'        'Object Layer **********************************
'        If .ObjGrh.GrhIndex <> 0 Then
'            Call DrawGrhLuz(.ObjGrh, _
'                    PixelOffSetXTemp, PixelOffSetYTemp, 1, 1, MapData(TileX, TileY).Light_Value)
'
'        End If
'        '***********************************************
'
'        'Char layer ************************************
'        If .CharIndex <> 0 Then
'            Call CharRender(charlist(.CharIndex), .CharIndex, PixelOffSetXTemp, PixelOffSetYTemp)
'        End If
'
'        'If UserMap = 1 Then
'        '    Call RenderBarcos(TileX, TileY, TileX, TileY, PixelOffSetX, PixelOffSetY)
'        'End If
'
'        'Layer 3 *****************************************
'        If .Graphic(3).GrhIndex > 0 Then
'            'Draw
'            SupIndex = GrhData(.Graphic(3).GrhIndex).FileNum
'            If ((SupIndex >= 7000 And SupIndex <= 7008) Or (SupIndex >= 1261 And SupIndex <= 1287) Or SupIndex = 648 Or SupIndex = 645) Then
'                If mOpciones.TransparencyTree = True And UserPos.X >= TileX - 3 And UserPos.X <= TileX + 3 And UserPos.Y >= TileY - 5 And UserPos.Y <= TileY Then
'                    Call DrawGrh(.Graphic(3), PixelOffSetXTemp, PixelOffSetYTemp, 1, 1, 0, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, 180))
'                Else
'                    Call DrawGrhLuz(.Graphic(3), PixelOffSetXTemp, PixelOffSetYTemp, 1, 1, MapData(TileX, TileY).Light_Value)
'                End If
'            Else
'               Call DrawGrh(.Graphic(3), PixelOffSetXTemp, PixelOffSetYTemp, 1, 1)
'            End If
'        End If
'        '*************************************************
'
'
'        'Layer 3 Plus FX *****************************************
'        If .fXGrh.Started = 1 Then
'            Call DrawGrh(.fXGrh, PixelOffSetXTemp - FxData(.fX).OFFSETX, PixelOffSetYTemp - FxData(.fX).OFFSETY, 1, 1)
'            If .fXGrh.Started = 0 Then .fX = 0
'        End If
'        '************************************************
'
'    End With
'
'
'    If Not bTecho Then
'        If bAlpha < 255 Then
'            If tTick < (GetTickCount() And &H7FFFFFFF) - 30 Then
'                bAlpha = bAlpha + IIf(bAlpha + 8 < 255, 8, 255 - bAlpha)
'                ColorTecho = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, bAlpha)
'                tTick = (GetTickCount() And &H7FFFFFFF)
'            End If
'        End If
'    Else
'        If bAlpha > 0 Then
'            If tTick < (GetTickCount() And &H7FFFFFFF) - 30 Then
'                bAlpha = bAlpha - IIf(bAlpha - 8 > 0, 8, bAlpha)
'                ColorTecho = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, bAlpha)
'                tTick = (GetTickCount() And &H7FFFFFFF)
'            End If
'        End If
'    End If
'
'
'    'Layer 4 **********************************
'    If MapData(TileX, TileY).Graphic(4).GrhIndex And bAlpha > 0 Then
'        'Draw
'        Call DrawGrhIndex(MapData(TileX, TileY).Graphic(4).GrhIndex, _
'            TilePixelWidth + PixelOffSetX, _
'            TilePixelHeight + PixelOffSetY, _
'            1, ColorTecho)
'    End If
'        '**********************************
''
''    'TODO : Check this!!
''    Dim ColorLluvia As Long
''    If ZonaActual > 0 Then
''        If Zonas(ZonaActual).Terreno <> eTerreno.Dungeon Then
''            If bRain Then
''                'Figure out what frame to draw
''                If llTick < (GetTickCount() And &H7FFFFFFF) - 50 Then
''                    iFrameIndex = iFrameIndex + 1
''                    If iFrameIndex > 7 Then iFrameIndex = 0
''                    llTick = (GetTickCount() And &H7FFFFFFF)
''                End If
''                ColorLluvia = D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.B, 140)
''                For Y = 0 To 5
''                    For X = 0 To 6
''                            Call Engine_Render_Rectangle(LTLluvia(X), LTLluvia(Y), RLluvia(iFrameIndex).Right - RLluvia(iFrameIndex).Left, RLluvia(iFrameIndex).Bottom - RLluvia(iFrameIndex).Top, RLluvia(iFrameIndex).Left, RLluvia(iFrameIndex).Top, RLluvia(iFrameIndex).Right - RLluvia(iFrameIndex).Left, RLluvia(iFrameIndex).Bottom - RLluvia(iFrameIndex).Top, , , , 5556, ColorLluvia, ColorLluvia, ColorLluvia, ColorLluvia)
''                    Next X
''                Next Y
''            End If
''        End If
''    End If
''
''    Call Dialogos.Render
''    Call DibujarCartel
''    Call DialogosClanes.Draw
'
'
''
''    If mOpciones.Niebla = True Then
''        Call RenderNiebla
''    End If
''
''    If AlphaCuenta > 0 Then
''         Call RenderCuentaRegresiva
''    End If
''
''    If AlphaBlood > 0 Then
''         Call RenderBlood
''    End If
''
''    If AlphaBloodUserDie > 0 Then
''         Call RenderUserDieBlood
''    End If
''
''    If AlphaTextKills > 0 Then
''         Call RenderTextKills
''    End If
''
''    If bRain = True Then
''        Call RenderRelampago
''    End If
''
''    If UserCiego = True Then
''         Call RenderCeguera
''    End If
''
''    If FPSFLAG Then Call DrawFont("FPS: " & FPS, 740, 260, D3DColorRGBA(255, 255, 255, 160))
'End Sub

Function CalcAlpha(Tiempo As Long, STiempo As Long, MaxAlpha As Byte, Tempo As Single) As Byte
Dim tmpInt As Long

tmpInt = (Tiempo - STiempo) / Tempo
If tmpInt >= 0 Then
CalcAlpha = IIf(tmpInt > MaxAlpha, MaxAlpha, tmpInt)
End If
End Function


Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
If ZonaActual > 0 Then
    If Zonas(ZonaActual).Terreno <> eTerreno.Dungeon Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave(SND_LLUVIAIN, 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave(SND_LLUVIAOUT, 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    Else
        If frmMain.IsPlaying <> PlayLoop.plNone Then
            Call Audio.StopWave(RainBufferIndex)
            RainBufferIndex = 0
            frmMain.IsPlaying = PlayLoop.plNone
        End If
        
        
    End If
    
    Call ReproducirSonidosDeAmbiente
    DoFogataFx
    
End If
End Function

Function HayUserAbajo(ByVal x As Integer, ByVal y As Integer, ByVal GrhIndex As Integer) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.x >= x - (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.x <= x + (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.y >= y - (GrhData(GrhIndex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.y <= y
    End If
End Function

Sub LoadGraphics()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero - complete rewrite
'Last Modify Date: 11/03/2006
'Initializes the SurfaceDB and sets up the rain rects
'**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(D3DDevice, D3DX, ClientSetup.bUseVideo, DirRecursos & "Graphics.AO", ClientSetup.byMemory)
    
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128
    
    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
End Sub

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal EngineSpeed As Single) As Boolean
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************
    
    IniPath = App.path & "\Init\"
    
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = Round(setWindowTileHeight \ 2, 0)
    HalfWindowTileWidth = Round(setWindowTileWidth \ 2, 0)
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = EngineSpeed
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    'MinXBorder = 1 + (WindowTileWidth \ 2)
    'MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    'MinYBorder = 1 + (WindowTileHeight \ 2)
    'MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
    
    'Resize mapdata array
    'ReDim MapData(1 To XMaxMapSize, 1 To YMaxMapSize, 1 To 2) As MapBlock
    
    'Set intial user position
    UserPos.x = 1
    UserPos.y = 1
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    'Set the view rect
    With MainViewRect
        .Left = MainViewLeft
        .Top = MainViewTop
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    'Set the dest rect
    With MainDestRect
        .Left = TilePixelWidth * TileBufferSize - TilePixelWidth
        .Top = TilePixelHeight * TileBufferSize - TilePixelHeight
        .Right = .Left + MainViewWidth
        .Bottom = .Top + MainViewHeight
    End With
    
    
IniciarD3D

    Call CargarFont
    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    'Call CargarParticulas
    
    'Call General_Particle_Create(1, 150, 800, -1, 20, -15)
    
    Set TestPart = New clsParticulas
    TestPart.Texture = 14386
    TestPart.ParticleCounts = 35
    TestPart.ReLocate 400, 400
    TestPart.Begin
    
    'Actual = 2
    'Particle_Group_Make Actual, 1, 150, 850, Particula(Actual).VarZ, Particula(Actual).VarX, Particula(Actual).VarY, Particula(Actual).AlphaInicial, Particula(Actual).RedInicial, Particula(Actual).GreenInicial, _
    'Particula(Actual).BlueInicial, Particula(Actual).AlphaFinal, Particula(Actual).RedFinal, Particula(Actual).GreenFinal, Particula(Actual).BlueFinal, Particula(Actual).NumOfParticles, Particula(Actual).gravity, Particula(Actual).Texture, Particula(Actual).Zize, Particula(Actual).Life
    
    
    LTLluvia(0) = 224
    LTLluvia(1) = 352
    LTLluvia(2) = 480
    LTLluvia(3) = 608
    LTLluvia(4) = 736
    LTLluvia(5) = 864
    LTLluvia(6) = 992
    LTLluvia(7) = 1120
    Call LoadGraphics
    
    InitTileEngine = True
End Function


Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer, Optional ByVal Update As Boolean = False, Optional ByVal x As Integer = 0, Optional ByVal y As Integer = 0)
'***************************************************
'Author: Arron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Updates the game's model and renders everything.
'***************************************************
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
    
    '****** Set main view rectangle ******
    MainViewRect.Left = (DisplayFormLeft / Screen.TwipsPerPixelX) + MainViewLeft
    MainViewRect.Top = (DisplayFormTop / Screen.TwipsPerPixelY) + MainViewTop
    MainViewRect.Right = MainViewRect.Left + MainViewWidth
    MainViewRect.Bottom = MainViewRect.Top + MainViewHeight
    
    
    If EngineRun Then
    
    If Update = True Then
       ' D3DDevice.BeginScene
        Debug.Print " UPDATE SCREEEN"
        'Call UpdateRenderScreen(X, Y, OffsetCounterX, OffsetCounterY)
        'D3DDevice.EndScene
        Exit Sub
    End If
    
        If UserEmbarcado Then
            OffsetCounterX = -BarcoOffSetX
            OffsetCounterY = -BarcoOffSetY
        ElseIf UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.x <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.x * timerTicksPerFrame * 1.2
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.x) Then
                    OffsetCounterX = 0
                    AddtoUserPos.x = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame * 1.2
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.y = 0
                    UserMoving = False
                End If
            End If
        End If
        
        'Update mouse position within view area
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
           
        If GoingHome = 1 Or GoingHome = 2 Then
            BlurIntensity = 5
        Else
            BlurIntensity = 0
        End If
        
        'Set the motion blur if needed
        If UseMotionBlur Then
            If BlurIntensity > 0 And BlurIntensity < 255 Or ZoomLevel > 0 Then
                FrameUseMotionBlur = True
                D3DDevice.SetRenderTarget BlurSurf, Nothing, ByVal 0
            Else
                FrameUseMotionBlur = False
            End If
        End If
        
        If UseMotionBlur Then
            If BlurIntensity < 255 Then
                BlurIntensity = BlurIntensity + (timerElapsedTime * 0.01)
                If BlurIntensity > 255 Then BlurIntensity = 255
            End If
        End If
                
        D3DDevice.BeginScene
        
        'Clear the screen with a solid color (to prevent artifacts)
        If Not FrameUseMotionBlur Then
            D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        End If
        
        '****** Update screen ******
        If Conectar Then
            Call RenderConectar
        'ElseIf UserCiego Then
        '    Call CleanViewPort
        Else
            Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
        End If
        
    

        'End the rendering (scene)
        D3DDevice.EndScene
                                        
               
        'Flip the backbuffer to the screen
        If Conectar Then
            D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
        Else
            D3DDevice.Present RectJuego, ByVal 0, 0, ByVal 0
        End If

        'Screen
        If ScreenShooterCapturePending Then
            DoEvents
            Call ScreenCapture(True)
            ScreenShooterCapturePending = False
        End If
    
    
        'Limit FPS to 60 (an easy number higher than monitor's vertical refresh rates)
        'While General_Get_Elapsed_Time2() < 15.5
        '    DoEvents
        'Wend
        
        'timer_ticks_per_frame = General_Get_Elapsed_Time() * 0.029
        
        'FPS update
        If fpsLastCheck + 1000 < (GetTickCount() And &H7FFFFFFF) Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            fpsLastCheck = (GetTickCount() And &H7FFFFFFF)
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
        
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        If timerElapsedTime <= 0 Then timerElapsedTime = 1
        timerTicksPerFrame = timerElapsedTime * EngineSpeed()
    End If
End Sub

Public Sub RenderText(ByVal lngXPos As Integer, ByVal lngYPos As Integer, ByRef strText As String, ByVal lngColor As Long)
    If strText <> "" Then
        Call DrawFont(strText, lngXPos, lngYPos, lngColor)
    End If
End Sub

Public Sub RenderTextCentered(ByVal lngXPos As Integer, ByVal lngYPos As Integer, ByRef strText As String, ByVal lngColor As Long)
    If strText <> "" Then
        Call DrawFont(strText, lngXPos, lngYPos, lngColor, True)
    End If
End Sub

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Public Sub CharRender(ByRef rChar As Char, ByVal CharIndex As Integer, ByVal PixelOffSetX As Integer, ByVal PixelOffSetY As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
    Dim VelChar As Single
    Dim ColorPj As Long
    Dim ShowPJ As Byte
    Dim ShowPJ_Alpha As Byte
    
    With rChar
        If .Moving Then
            If .nombre = "" Then
                VelChar = 0.75
            ElseIf Left(.nombre, 1) = "!" Then
                VelChar = 0.75
            Else
                VelChar = 1.2
            End If
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame * VelChar
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame * VelChar
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        

        
        'If done moving stop animation
        If Not moved And True Then
            'Stop animations
            .Quieto = .Quieto + 1
            If .Quieto >= FPS / 35 Then 'Esto es para que las animacion sean continuas mientras se camine, por ejemplo sin esto el andar del golum se ve feo
            .Quieto = 0
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            If .Arma.WeaponAttack = 0 Then
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 1
            Else
                If .Arma.WeaponWalk(.Heading).Started = 0 Then
                    .Arma.WeaponAttack = 0
                    .Arma.WeaponWalk(.Heading).FrameCounter = 1
                End If
            End If
            
            If .Escudo.ShieldAttack = 0 Then
                .Escudo.ShieldWalk(.Heading).Started = 0
                .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            Else
                If .Escudo.ShieldWalk(.Heading).Started = 0 Then
                    .Escudo.ShieldAttack = 0
                    .Escudo.ShieldWalk(.Heading).FrameCounter = 1
                End If
            End If
            End If
            
            .Moving = False
        Else
            .Quieto = 0
        End If
                
        PixelOffSetX = PixelOffSetX + .MoveOffsetX
        PixelOffSetY = PixelOffSetY + .MoveOffsetY
        
        'Verificamos si vamos a mostrar el PJ
        ShowPJ = 0
        ShowPJ_Alpha = 0
        
        If Not .invisible Then
            ShowPJ = 1
            ShowPJ_Alpha = 255
        ElseIf UserCharIndex = CharIndex Then
            ShowPJ = 2
            ShowPJ_Alpha = 120
        ElseIf .invisible = True Then
            If charEsClan(CharIndex) Then
                ShowPJ = 3
                ShowPJ_Alpha = 120
            ElseIf .oculto = True Then
                ShowPJ = 0
                ShowPJ_Alpha = 0
            End If
        End If
    
        ColorPj = D3DColorRGBA(IluRGB.r, IluRGB.G, IluRGB.B, .Alpha)
        If .Head.Head(.Heading).GrhIndex Then
                If .invisible Then
                    If ShowPJ = 2 Or ShowPJ = 3 Then
                        .Alpha = ShowPJ_Alpha
                    ElseIf .ContadorInvi > 0 Then
                            If .iTick < (GetTickCount() And &H7FFFFFFF) - 35 Then
                                If .ContadorInvi > 30 And .ContadorInvi <= 60 And .Alpha < 255 Then
                                    .Alpha = .Alpha + 5
                                ElseIf .ContadorInvi <= 30 And .Alpha > 0 Then
                                    .Alpha = .Alpha - 5
                                End If
                                .ContadorInvi = .ContadorInvi - 1
                                .iTick = (GetTickCount() And &H7FFFFFFF)
                            End If
                    Else
                        .ContadorInvi = INTERVALO_INVI
                    End If
                End If
            If .Alpha > 0 Then
                If .priv = 9 Then
                    ColorPj = D3DColorRGBA(10, 10, 10, 255)
                Else
                    ColorPj = D3DColorRGBA(IluRGB.r, IluRGB.G, IluRGB.B, .Alpha)
                    If .congelado = True Then
                        ColorPj = D3DColorRGBA(0, 175, 255, 225)
                    End If
                End If
                Dim Sombra As Boolean
                If ZonaActual > 0 Then
                    Sombra = .invisible Or .muerto Or Zonas(ZonaActual).Terreno = eTerreno.Dungeon Or .priv = 10
                End If
                                                
                Dim TempBodyOffsetX As Integer
                Dim TempBodyOffsetY As Integer
                Dim TempHeadOffsetY As Integer
                Dim TempHeadOffsetX As Integer
               
                TempHeadOffsetY = .Body.HeadOffset.y
                TempHeadOffsetX = .Body.HeadOffset.x
                TempBodyOffsetY = PixelOffSetY
                TempBodyOffsetX = PixelOffSetX
                
                If .Chiquito = True Then 'CHIQUITOLIN
                    TempHeadOffsetY = TempHeadOffsetY + 3
                    TempHeadOffsetX = TempHeadOffsetX - 2
                    TempBodyOffsetY = TempBodyOffsetY + 10
                    TempBodyOffsetX = TempBodyOffsetX + 3
                End If
                
'                If .equitando = True Then 'CHIQUITOLIN
'                    TempHeadOffsetY = TempHeadOffsetY - 67
'                    TempHeadOffsetX = TempHeadOffsetX + 1
'                End If

                'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then _
                    Call DrawGrhShadow(.Body.Walk(.Heading), TempBodyOffsetX, TempBodyOffsetY, 1, 0.5, IIf(Sombra, 0, 1), ColorPj, 255, .Chiquito)
            
                'Draw Head
                Call DrawGrhShadow(.Head.Head(.Heading), TempBodyOffsetX + TempHeadOffsetX, TempBodyOffsetY + TempHeadOffsetY, 1, 0, IIf(Sombra, 0, 2), ColorPj, 255, .Chiquito)
                    
                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then _
                        Call DrawGrhShadow(.Casco.Head(.Heading), TempBodyOffsetX + TempHeadOffsetX, TempBodyOffsetY + TempHeadOffsetY, 1, 0, IIf(Sombra, 0, 2), ColorPj, 255, .Chiquito)
                                                     
                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                    Call DrawGrhShadow(.Arma.WeaponWalk(.Heading), TempBodyOffsetX, TempBodyOffsetY, 1, 0.5, IIf(Sombra, 0, 1), ColorPj, 255, .Chiquito)
                    
                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                    Call DrawGrhShadow(.Escudo.ShieldWalk(.Heading), TempBodyOffsetX, TempBodyOffsetY, 1, 0.5, IIf(Sombra, 0, 1), ColorPj, 255, .Chiquito)
                                                  
                'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then _
                    Call DrawGrhShadowOff(.Body.Walk(.Heading), TempBodyOffsetX, TempBodyOffsetY, 1, 0.5, ColorPj, .Chiquito)
            
                'Draw Head
                Call DrawGrhShadowOff(.Head.Head(.Heading), TempBodyOffsetX + TempHeadOffsetX, TempBodyOffsetY + TempHeadOffsetY, 1, 0, ColorPj, .Chiquito)
                                  
                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then _
                    Call DrawGrhShadowOff(.Casco.Head(.Heading), TempBodyOffsetX + TempHeadOffsetX, TempBodyOffsetY + TempHeadOffsetY, 1, 0, ColorPj, .Chiquito)
                    
                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                    Call DrawGrhShadowOff(.Arma.WeaponWalk(.Heading), TempBodyOffsetX, TempBodyOffsetY, 1, 0.5, ColorPj, .Chiquito)
                    
                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                    Call DrawGrhShadowOff(.Escudo.ShieldWalk(.Heading), TempBodyOffsetX, TempBodyOffsetY, 1, 0.5, ColorPj, .Chiquito)
                
                
                    'Draw name over head
                    If LenB(.nombre) > 0 And (ShowPJ = 2 Or ShowPJ = 1) And .priv <> 10 Then
                        If Nombres Then
                            Pos = InStr(.nombre, "<")
                            If Pos = 0 Then Pos = Len(.nombre) + 2
                            
                            If .invisible = True Then
                                color = D3DColorRGBA(200, 200, 200, ShowPJ_Alpha)
                            ElseIf .priv = 0 Then
                                If .Criminal Then
                                    color = D3DColorRGBA(ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).B, 200)
                                Else
                                    color = D3DColorRGBA(ColoresPJ(49).r, ColoresPJ(49).G, ColoresPJ(49).B, 200)
                                End If
                            Else
                                color = D3DColorRGBA(ColoresPJ(.priv).r, ColoresPJ(.priv).G, ColoresPJ(.priv).B, 200)
                            End If
                            
                            'Nick
                            line = Left$(.nombre, Pos - 2)
                            If Left(line, 1) = "!" Then
                                line = Right(line, Len(line) - 1)
                                Pos = Pos - 1
                            End If
                            Call RenderTextCentered(PixelOffSetX + TilePixelWidth \ 2, PixelOffSetY + 30, line, color)
                            
                            'Clan
                            line = mid$(.nombre, Pos)
                            Call RenderTextCentered(PixelOffSetX + TilePixelWidth \ 2, PixelOffSetY + 45, line, color)
                            
'                            If .logged Then
'                                color = D3DColorRGBA(10, 200, 10, 200)
'                                Call RenderTextCentered(PixelOffSetX + TilePixelWidth \ 2, PixelOffSetY + 45, "(Online)", color)
'                            End If
'
                        End If
                    End If
            End If
        Else
            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then _
                Call DrawGrh(.Body.Walk(.Heading), PixelOffSetX, PixelOffSetY, 1, VelChar, IIf(Sombra, 0, 1), ColorPj)
        End If

        
        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffSetX + .Body.HeadOffset.x + 16, PixelOffSetY + .Body.HeadOffset.y, CharIndex)
        
        'Draw FX
        If .FxIndex <> 0 Then
            D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
                        Call DrawGrh(.fX, PixelOffSetX + FxData(.FxIndex).OFFSETX, PixelOffSetY + FxData(.FxIndex).OFFSETY, 1, 1, 0, D3DColorRGBA(255, 255, 255, 170))
            D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
            'Check if animation is over
            If .fX.Started = 0 Then
                .FxIndex = 0
            End If
        End If
        
    End With
End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charlist(CharIndex)
        
        
        'If fX > 0 Then
            'If CharIndex = UserCharIndex Then  ' And Not UserMeditar
            .FxIndex = fX
            If fX > 0 Then
                Call InitGrh(.fX, FxData(fX).Animacion)
        
                .fX.Loops = Loops
            End If
            'End If
        'End If
    End With
End Sub

Public Sub SetAreaFx(ByVal x As Integer, ByVal y As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    
    If fX > 0 Then
        Call InitGrh(MapData(x, y).fXGrh, FxData(fX).Animacion)
        MapData(x, y).fX = fX
        MapData(x, y).fXGrh.Loops = Loops
    End If
           
 
End Sub

Private Sub CleanViewPort()
'Limpiar
End Sub

Public Function Char_Pos_Get(ByVal CharIndex As Integer, ByRef x As Integer, ByRef y As Integer)
    
    If CharIndex < 1 Then Exit Function
    With charlist(CharIndex)
        x = .Pos.x
        y = .Pos.y
        
        If x > 0 And y > 0 Then
            Char_Pos_Get = True
        Else
            Char_Pos_Get = False
        End If
    End With

End Function

Public Function charEsClan(ByVal Char As Integer) As Boolean
charEsClan = False
If Char > 0 Then
    Dim tempTag As String
    Dim tempPos As Integer
    Dim miTag As String
    Dim miTempPos As Integer
    With charlist(Char)
        miTempPos = getTagPosition(charlist(UserCharIndex).nombre)
        miTag = mid$(charlist(UserCharIndex).nombre, miTempPos)
        tempPos = getTagPosition(.nombre)
        tempTag = mid$(.nombre, tempPos)
        If tempTag = miTag And miTag <> "" And tempTag <> "" Then
            charEsClan = True
            Exit Function
        End If
        
    End With
End If
End Function

