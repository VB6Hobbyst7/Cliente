Attribute VB_Name = "ModAgregadoParaQuest"
Public ObjData()          As ObjDatas
Public Type tQuestNpc

    NpcIndex As Integer
    Amount As Integer

End Type

Public QuestList() As tQuest
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

Public Type tQuest

    nombre As String
    desc As String
    NextQuest As String
    DescFinal As String
    RequiredLevel As Byte
    
    RequiredQuest As Byte
    
    RequiredOBJs As Byte
    RequiredOBJ() As Obj
    
    RequiredNPCs As Byte
    RequiredNPC() As tQuestNpc
    
    RewardGLD As Long
    RewardEXP As Long
    
    RewardOBJs As Byte
    RewardOBJ() As Obj
    Repetible As Byte

End Type


Public NpcData()          As NpcDatas
Public Type NpcDatas

    Name As String
    desc As String
    Body As Integer
    Hp As Long
    exp As Long
    oro As Long
    MinHit As Integer
    MaxHit As Integer
    Head As Integer
    NumQuiza As Byte
    QuizaDropea() As Integer
    ExpClan As Long
    
End Type
Dim Leer As New clsIniManager
Public NumQuest  As Integer
Public PosMap()           As Integer
Public Type ObjDatas

    GrhIndex As Long ' Indice del grafico que representa el obj
    Name As String
    MinDef As Integer
    MaxDef As Integer
    MinHit As Integer
    MaxHit As Integer
    ObjType As Byte
    Texto As String
    info As String
    CreaGRH As String
    CreaLuz As String
    CreaParticulaPiso As Integer
    proyectil As Byte
    Raices As Integer
    Madera As Integer
    PielLobo As Integer
    PielOsoPardo As Integer
    PielOsoPolar As Integer
    LingH As Integer
    LingP As Integer
    LingO As Integer
    Destruye As Byte
    SkHerreria As Byte
    SkPociones As Byte
    Sksastreria As Byte
    Valor As Long

End Type


Public Sub CargarNpc()

ObjFile = "C:\Users\waalter\Documents\GitHub\Recursos\Dat\NPCs.dat"
Call Leer.Initialize(ObjFile)
NumNpcs = Val(Leer.GetValue("INIT", "NumNPCs"))

    ReDim NpcData(0 To NumNpcs) As NpcDatas
'NumQuest = Val(Leer.GetValue("INIT", "NUMQUESTS"))
For npc = 1 To NumNpcs
        DoEvents
        
        NpcData(npc).Name = Leer.GetValue("NPC" & npc, "Name")

        If NpcData(npc).Name = "" Then
            NpcData(npc).Name = "Vacio"

        End If

        NpcData(npc).desc = Leer.GetValue("npc" & npc, "desc")
      
        NpcData(npc).exp = Val(Leer.GetValue("npc" & npc, "exp"))
        NpcData(npc).Head = Val(Leer.GetValue("npc" & npc, "Head"))
        NpcData(npc).Hp = Val(Leer.GetValue("npc" & npc, "Hp"))
        NpcData(npc).MaxHit = Val(Leer.GetValue("npc" & npc, "MaxHit"))
        NpcData(npc).MinHit = Val(Leer.GetValue("npc" & npc, "MinHit"))
        NpcData(npc).oro = Val(Leer.GetValue("npc" & npc, "oro"))
        
        NpcData(npc).ExpClan = Val(Leer.GetValue("npc" & npc, "GiveEXPClan"))
       
        aux = Val(Leer.GetValue("npc" & npc, "NumQuiza"))

        If aux = 0 Then
            NpcData(npc).NumQuiza = 0
        Else
            NpcData(npc).NumQuiza = Val(aux)
            ReDim NpcData(npc).QuizaDropea(1 To NpcData(npc).NumQuiza) As Integer

            For loopc = 1 To NpcData(npc).NumQuiza
               
                NpcData(npc).QuizaDropea(loopc) = Val(Leer.GetValue("npc" & npc, "QuizaDropea" & loopc))
                ' Debug.Print NpcData(Npc).QuizaDropea(loopc)
            Next loopc

        End If

    Next npc

End Sub

Public Sub CargarQuests()
   ObjFile = "C:\Users\waalter\Documents\GitHub\Recursos\Dat\QUESTS.dat"
Call Leer.Initialize(ObjFile)
  NumQuest = Val(Leer.GetValue("INIT", "NumQuests"))
  ReDim QuestList(1 To NumQuest)
  ReDim PosMap(1 To NumQuest) As Integer
  For Nquest = 1 To NumQuest
        DoEvents
        
        QuestList(Nquest).nombre = Leer.GetValue("QUEST" & Nquest, "Nombre")
        
        QuestList(Nquest).desc = Leer.GetValue("QUEST" & Nquest, "Desc")
        QuestList(Nquest).NextQuest = Leer.GetValue("QUEST" & Nquest, "NextQuest")
        QuestList(Nquest).DescFinal = Leer.GetValue("QUEST" & Nquest, "DescFinal")
        QuestList(Nquest).RequiredLevel = Val(Leer.GetValue("QUEST" & Nquest, "RequiredLevel"))
        PosMap(Nquest) = Leer.GetValue("QUEST" & Nquest, "PosMap")
    Next Nquest
End Sub


Public Sub CargarObjetos()
ObjFile = "C:\Users\waalter\Documents\GitHub\Recursos\Dat\obj.dat"
Call Leer.Initialize(ObjFile)
NumOBJs = Val(Leer.GetValue("INIT", "NumObjs"))
ReDim ObjData(0 To NumOBJs) As ObjDatas
For Obj = 1 To NumOBJs
        DoEvents
        ObjData(Obj).GrhIndex = Val(Leer.GetValue("OBJ" & Obj, "grhindex"))
        ObjData(Obj).Name = Leer.GetValue("OBJ" & Obj, "Name")
        ObjData(Obj).MinDef = Val(Leer.GetValue("OBJ" & Obj, "MinDef"))
        ObjData(Obj).MaxDef = Val(Leer.GetValue("OBJ" & Obj, "MaxDef"))
        ObjData(Obj).MinHit = Val(Leer.GetValue("OBJ" & Obj, "MinHit"))
        ObjData(Obj).MaxHit = Val(Leer.GetValue("OBJ" & Obj, "MaxHit"))
        ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
        ObjData(Obj).info = Leer.GetValue("OBJ" & Obj, "Info")
        ObjData(Obj).Texto = Leer.GetValue("OBJ" & Obj, "Texto")
        ObjData(Obj).CreaGRH = Leer.GetValue("OBJ" & Obj, "CreaGRH")
        ObjData(Obj).CreaLuz = Leer.GetValue("OBJ" & Obj, "CreaLuz")
        ObjData(Obj).CreaParticulaPiso = Val(Leer.GetValue("OBJ" & Obj, "CreaParticulaPiso"))
        ObjData(Obj).proyectil = Val(Leer.GetValue("OBJ" & Obj, "proyectil"))
        ObjData(Obj).Raices = Val(Leer.GetValue("OBJ" & Obj, "Raices"))
        ObjData(Obj).Madera = Val(Leer.GetValue("OBJ" & Obj, "Madera"))
        ObjData(Obj).PielLobo = Val(Leer.GetValue("OBJ" & Obj, "PielLobo"))
        ObjData(Obj).PielOsoPardo = Val(Leer.GetValue("OBJ" & Obj, "PielOsoPardo"))
        ObjData(Obj).PielOsoPolar = Val(Leer.GetValue("OBJ" & Obj, "PielOsoPolar"))
        ObjData(Obj).LingH = Val(Leer.GetValue("OBJ" & Obj, "LingH"))
        ObjData(Obj).LingP = Val(Leer.GetValue("OBJ" & Obj, "LingP"))
        ObjData(Obj).LingO = Val(Leer.GetValue("OBJ" & Obj, "LingO"))
        ObjData(Obj).Destruye = Val(Leer.GetValue("OBJ" & Obj, "Destruye"))
        'ObjData(Obj).SkHerreria = Val(Leer.GetValue("OBJ" & Obj, "SkHerreria"))
        ObjData(Obj).SkPociones = Val(Leer.GetValue("OBJ" & Obj, "SkPociones"))
        ObjData(Obj).Sksastreria = Val(Leer.GetValue("OBJ" & Obj, "Sksastreria"))
        ObjData(Obj).Valor = Val(Leer.GetValue("OBJ" & Obj, "Valor"))
        
    Next Obj
  
End Sub

