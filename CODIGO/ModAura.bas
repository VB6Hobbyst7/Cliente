Attribute VB_Name = "ModAura"
Option Explicit
Public Enum UpdateAuras
    Arma
    Armadura
    Escudo
    Casco
    Anillo
    Sets
End Enum

Public Type tAuras
    R As Byte
    G As Byte
    b As Byte
    AuraGrh As Integer
    Giratoria As Boolean
    color As Long
    OffSetX As Integer
    OffSetY As Integer
End Type

Public MaxAuras As Byte
Public Auras() As tAuras

Public Sub CargarAuras()

    Dim AuraPath As String, loopc As Long, Gira As Byte
    AuraPath = App.path & "\INIT\Auras.ini"
    MaxAuras = Val(GetVar(AuraPath, "INIT", "MaxAuras"))
 
    If MaxAuras > 0 Then
        ReDim Auras(1 To MaxAuras) As tAuras
     
        For loopc = 1 To MaxAuras
            With Auras(loopc)
                .R = Val(GetVar(AuraPath, "AURA" & loopc, "R"))
                .G = Val(GetVar(AuraPath, "AURA" & loopc, "G"))
                .b = Val(GetVar(AuraPath, "AURA" & loopc, "B"))
             
                .AuraGrh = Val(GetVar(AuraPath, "AURA" & loopc, "GRH"))
             
                .OffSetX = Val(GetVar(AuraPath, "AURA" & loopc, "OffSetX"))
                .OffSetY = Val(GetVar(AuraPath, "AURA" & loopc, "OffSetY"))
             
                Gira = Val(GetVar(AuraPath, "AURA" & loopc, "GIRATORIA"))
                If Gira <> 0 Then
                    .Giratoria = True
                Else
                    .Giratoria = False
                End If
            End With
        Next loopc
    End If 'Maxauras > 0
End Sub

