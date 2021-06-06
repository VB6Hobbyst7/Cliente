VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBarco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TIEMPO_EN_PUERTO As Long = 15

Private Const ACELERACION      As Byte = 60

Private Const VELOCIDAD        As Single = 4

Dim RutaBarco()                As Position

Dim Paso                       As Byte

Dim Pos                        As Position

Dim Direccion                  As Byte

Dim UPasos                     As Byte

Dim GrhBarco                   As Grh

Dim GrhIBarco                  As Integer

Dim Sentido                    As Byte

Dim BOffSetX                   As Single

Dim BOffSetY                   As Single

Dim pEmbarcado                 As Boolean

Dim EnMovimiento               As Boolean

Public TickPuerto              As Long

Dim Puerto                     As Byte

Dim CapitanIndex               As Integer

Dim Pasajeros(0 To 4)          As Integer

Dim FactorBarco                As Single

''''DEBUGING''''
Dim PasoCount                  As Integer

Dim PasoTick                   As Long

Dim InitTick                   As Long

Public Property Let Embarcado(ByVal Value As Boolean)

    If Value Then
        UserEmbarcado = True
        UserPos.X = Pos.X
        UserPos.Y = Pos.Y

    End If

    pEmbarcado = Value

End Property

Public Property Get Embarcado() As Boolean
    Embarcado = pEmbarcado

End Property

Public Property Get X() As Integer
    X = Pos.X

End Property

Public Property Get Y() As Integer
    Y = Pos.Y

End Property

Private Function NextPuerto() As Integer

    Dim p As Integer

    If Sentido = 0 Then
        p = Puerto + 1

        If p > NUM_PUERTOS Then p = 1
    Else
        p = Puerto - 1

        If p < 1 Then p = NUM_PUERTOS

    End If

    NextPuerto = p

End Function

Public Sub Init(ByVal Ruta As String, _
                ByVal mPaso As Byte, _
                ByVal X As Integer, _
                ByVal Y As Integer, _
                ByVal TiempoPuerto As Long, _
                ByVal mSentido As Byte, _
                ByRef mPasajeros() As Integer)

    Dim Rutas() As String

    Rutas = Split(Ruta, ";")

    Dim I As Integer

    UPasos = UBound(Rutas)
    ReDim RutaBarco(0 To UPasos) As Position

    For I = 0 To UPasos
        RutaBarco(I).X = Val(ReadField(1, Rutas(I), 44))
        RutaBarco(I).Y = Val(ReadField(2, Rutas(I), 44))
    Next I

    Paso = mPaso
    Pos.X = X
    Pos.Y = Y
    Sentido = mSentido

    For I = IIf(Sentido = 0, 1, NUM_PUERTOS) To IIf(Sentido = 0, NUM_PUERTOS, 1) Step IIf(Sentido = 0, 1, -1)

        If Paso > Puertos(I).Paso(Sentido) Then
            Puerto = I

        End If

    Next I

    Puerto = NextPuerto

    If TiempoPuerto > 0 Then
        TickPuerto = (GetTickCount() And &H7FFFFFFF) - TiempoPuerto
        EnMovimiento = False
    Else
        EnMovimiento = True

    End If

    'Capitan
    CapitanIndex = 10000 - Sentido
    charlist(CapitanIndex).Body = BodyData(428)
    charlist(CapitanIndex).Head = HeadData(502)
    charlist(CapitanIndex).Alpha = 255

    Pasajeros(0) = CapitanIndex

    For I = 1 To 4
        Pasajeros(I) = mPasajeros(I)

        If Pasajeros(I) = UserCharIndex And UserCharIndex > 0 Then
            Embarcado = True

        End If

    Next I

    If Pos.X > RutaBarco(Paso).X Then
        Direccion = west
        Call InitGrh(GrhBarco, 10384)
    ElseIf Pos.X < RutaBarco(Paso).X Then
        Direccion = east
        Call InitGrh(GrhBarco, 10388)
    ElseIf Pos.Y > RutaBarco(Paso).Y Then
        Direccion = north
        Call InitGrh(GrhBarco, 10396)
    Else 'If Pos.Y < RutaBarco(Paso).Y Then '''REVISAR
        Direccion = south
        Call InitGrh(GrhBarco, 10392)

    End If

    Call CalcularFactorBarco

    InitTick = (GetTickCount() And &H7FFFFFFF)

End Sub

Public Function Actualizar(ByVal X As Integer, ByVal Y As Integer) As Boolean
    Actualizar = Abs(X - Pos.X) > 20 Or Abs(Y - Pos.Y) > 20

End Function

Public Sub AgregarPasajero(ByVal Num As Byte, ByVal CharIndex As Integer)
    Pasajeros(Num) = CharIndex

    If CharIndex = UserCharIndex And UserCharIndex > 0 Then
        Embarcado = True
    
        UserPos.X = Pos.X
        UserPos.Y = Pos.Y
    
        Call frmMain.RefreshMiniMap
    
    End If

End Sub

Public Sub QuitarPasajero(ByVal Num As Byte)

    If Pasajeros(Num) = UserCharIndex Then
        Embarcado = False
        UserEmbarcado = False

    End If

    Pasajeros(Num) = 0

End Sub

Private Sub CalcularFactorBarco()

    Dim NP   As Byte

    Dim Dist As Integer, Dist2 As Integer

    NP = Puertos(Puerto).Paso(Sentido)
    Dist = Abs(RutaBarco(NP).X - Pos.X) + Abs(RutaBarco(NP).Y - Pos.Y)

    If Sentido = 0 Then
        If Puerto - 1 = 0 Then
            NP = Puertos(NUM_PUERTOS).Paso(Sentido)
        Else
            NP = Puertos(Puerto - 1).Paso(Sentido)

        End If

    Else

        If Puerto + 1 > NUM_PUERTOS Then
            NP = Puertos(1).Paso(Sentido)
        Else
            NP = Puertos(Puerto + 1).Paso(Sentido)

        End If

    End If

    'Sistema de aceleracion, era muy brusco sino cuando el barco llegaba ;)
    Dist2 = Abs(RutaBarco(NP).X - Pos.X) + Abs(RutaBarco(NP).Y - Pos.Y)

    If Dist <= ACELERACION Then
        FactorBarco = (Dist / ACELERACION) * (VELOCIDAD - 0.5) + 0.5
    ElseIf Dist2 <= ACELERACION Then
        FactorBarco = (Dist2 / ACELERACION) * (VELOCIDAD - 0.5) + 0.5
    Else
        FactorBarco = VELOCIDAD

    End If

End Sub

Public Sub Avanzar()

    Dim NP As Integer

    Select Case Direccion

        Case north
            Pos.Y = Pos.Y - 1

        Case south
            Pos.Y = Pos.Y + 1

        Case west
            Pos.X = Pos.X - 1

        Case east
            Pos.X = Pos.X + 1

    End Select

    If Embarcado Then
        UserPos.X = Pos.X
        UserPos.Y = Pos.Y

    End If

    Call CalcularFactorBarco

    If Pos.X = RutaBarco(Paso).X And Pos.Y = RutaBarco(Paso).Y Then
        If EsPuerto(Paso) Then
            TickPuerto = (GetTickCount() And &H7FFFFFFF)
            NP = NextPuerto
        
            Call Audio.PlayWave(215, Pos.X, Pos.Y)

            If Embarcado Then
                Call Dialogos.CreateDialog("¡Hemos llegado al puerto de " & Puertos(Puerto).nombre & "! Pueden descender aquí o seguir rumbo a " & Puertos(NP).nombre, CapitanIndex, 255, 255, 255)
            Else
                Call Dialogos.CreateDialog("¡Todos a bordo! Zarparemos hacia el puerto de " & Puertos(NP).nombre & " en unos instantes.", CapitanIndex, 255, 255, 255)

            End If

            Puerto = NP
        
        End If

        Paso = NextPaso
    
        If Pos.X > RutaBarco(Paso).X Then
            Direccion = west
            Call InitGrh(GrhBarco, 10384)
        ElseIf Pos.X < RutaBarco(Paso).X Then
            Direccion = east
            Call InitGrh(GrhBarco, 10388)
        ElseIf Pos.Y > RutaBarco(Paso).Y Then
            Direccion = north
            Call InitGrh(GrhBarco, 10396)
        ElseIf Pos.Y < RutaBarco(Paso).Y Then
            Direccion = south
            Call InitGrh(GrhBarco, 10392)

        End If

    End If

    'If Embarcado And EnMovimiento And TickPuerto = 0 Then MoveScreen (Direccion)

    'Debug.Print Pos.X & "  -  " & Pos.Y

    Call frmMain.RefreshMiniMap

End Sub

Private Function NextPaso() As Byte

    If Paso + 1 > UPasos Then
        NextPaso = 0
    Else
        NextPaso = Paso + 1

    End If

End Function

Private Function EsPuerto(ByVal Paso As Byte) As Boolean

    Dim I As Integer

    For I = 1 To NUM_PUERTOS

        If Puertos(I).Paso(Sentido) = Paso Then
            EsPuerto = True
            Exit For

        End If

    Next I

End Function

Public Sub Calcular()

    Dim Factor As Double

    'Debuging
    'If PasoTick = 0 Then
    '    Debug.Print Paso & " - " & Pos.X & ", " & Pos.Y & " - " & IIf(TickPuerto > 0, (GetTickCount() And &H7FFFFFFF) - TickPuerto, 0)
    '    PasoTick = (GetTickCount() And &H7FFFFFFF)
    'ElseIf (GetTickCount() And &H7FFFFFFF) - PasoTick >= 57000 Then
    '    PasoTick = 0
    'End If
    'If Paso = 1 And (GetTickCount() And &H7FFFFFFF) - InitTick > 60000 Then
    '    Debug.Print (GetTickCount() And &H7FFFFFFF) - InitTick
    '    End
    'End If
    If Abs(Pos.X - UserPos.X) > 25 Or Abs(Pos.Y - UserPos.Y) > 25 Then
        Set Barco(Sentido) = Nothing
        Exit Sub

    End If

    If TickPuerto > 0 Then 'Ahora el barco no zarpa hasta que el servidor le diga.
        EnMovimiento = False
        'If (GetTickCount() And &H7FFFFFFF) - TickPuerto < TIEMPO_EN_PUERTO * 1000 Then
        '    EnMovimiento = False
        'Else
        '    TickPuerto = 0
        '    EnMovimiento = True
        'End If
    Else
        EnMovimiento = True

    End If

    If EnMovimiento Then
        Factor = timerTicksPerFrame * FactorBarco * 8
    
        If Abs(BOffSetX) >= 32 Then
            BOffSetX = BOffSetX - Sgn(BOffSetX) * 32
            BOffSetY = 0
            Avanzar
        ElseIf Abs(BOffSetY) >= 32 Then
            BOffSetY = BOffSetY - Sgn(BOffSetY) * 32
            BOffSetX = 0
            Avanzar

        End If

        Select Case Direccion

            Case north
                BOffSetY = BOffSetY - Factor

            Case south
                BOffSetY = BOffSetY + Factor

            Case west
                BOffSetX = BOffSetX - Factor

            Case east
                BOffSetX = BOffSetX + Factor

        End Select
    
        If Embarcado Then
            BarcoOffSetX = BOffSetX
            BarcoOffSetY = BOffSetY

        End If

    End If

End Sub

Public Sub Render(ByVal TileX As Integer, _
                  ByVal TileY As Integer, _
                  ByVal PixelOffSetX As Single, _
                  ByVal PixelOffSetY As Single)

    Dim X    As Integer

    Dim Y    As Integer

    Dim OffX As Integer, AddX As Integer

    Dim OffY As Integer, AddY As Integer

    X = (Pos.X - TileX + 11) * 32 + PixelOffSetX + 256 + BOffSetX
    Y = (Pos.Y - TileY + 8) * 32 + PixelOffSetY + 256 + BOffSetY + 20

    'Call DrawGrh(GrhBarco, X, Y, 1, 1, 0)
    Select Case Direccion

        Case east
            Call DrawGrhIndex(5447, X, Y, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 255)) 'derecha
            

        Case west
            Call DrawGrhIndex(5446, X, Y, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 255)) 'isq
            

        Case north
            Call DrawGrhIndex(5497, X, Y, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 255)) 'arriba
            

        Case south
            Call DrawGrhIndex(5496, X, Y, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 255)) 'abajo
            

    End Select

    OffX = X
    OffY = Y
    'Dibujamos a los pasajeros y capitan

    Select Case Direccion

        Case north
            OffY = OffY - 155
            AddY = 20

        Case south
            OffY = OffY - 45
            AddY = -20

        Case east
            OffX = OffX + 67
            OffY = OffY - 44
            AddX = -20

        Case west
            OffX = OffX - 67
            OffY = OffY - 44
            AddX = 20

    End Select

    Dim I     As Integer

    Dim index As Integer

    For I = IIf(Direccion = north, 0, 4) To IIf(Direccion = north, 4, 0) Step IIf(Direccion = north, 1, -1)
        index = Pasajeros(I)

        If index > 0 Then
            charlist(index).Heading = Direccion
            charlist(index).priv = 10
            Call CharRender(charlist(index), index, OffX + AddX * IIf(I > 0 And (Direccion = east Or Direccion = west), 9.5 - I, 0), OffY + AddY * IIf(I > 0 And (Direccion = north Or Direccion = south), 6 - I, 0))

        End If

    Next I

    Select Case Direccion

        Case east
            Call DrawGrhIndex(5449, X - 108, Y - 29, 0, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 255))

        Case west
            Call DrawGrhIndex(5448, X + 56, Y - 29, 0, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 255))

        Case north
            Call DrawGrhIndex(5499, X, Y, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 220))

        Case south
            Call DrawGrhIndex(5498, X + 2, Y, 1, D3DColorRGBA(IluRGB.R, IluRGB.G, IluRGB.b, 220))
            index = Pasajeros(0)
            Call CharRender(charlist(index), index, OffX, OffY)

    End Select

End Sub
