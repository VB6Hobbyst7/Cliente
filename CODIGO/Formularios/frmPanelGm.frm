VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPanelGm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel GM"
   ClientHeight    =   8760
   ClientLeft      =   17580
   ClientTop       =   3735
   ClientWidth     =   4335
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   8280
      Width           =   4095
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton Command34 
         Caption         =   "Muelle de Nix"
         Height          =   255
         Left            =   120
         TabIndex        =   155
         Top             =   7320
         Width           =   1815
      End
      Begin VB.CommandButton Command32 
         Caption         =   "Muelle de Arkein"
         Height          =   255
         Left            =   2040
         TabIndex        =   154
         Top             =   6960
         Width           =   1815
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Muelle Lindos Oeste"
         Height          =   255
         Left            =   120
         TabIndex        =   153
         Top             =   6960
         Width           =   1815
      End
      Begin VB.CommandButton Command30 
         Caption         =   "Muelle de Arghal"
         Height          =   255
         Left            =   2040
         TabIndex        =   152
         Top             =   6600
         Width           =   1815
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Muelle Banderbille"
         Height          =   255
         Left            =   120
         TabIndex        =   151
         Top             =   6600
         Width           =   1815
      End
      Begin VB.CommandButton Command33 
         Caption         =   "Desierto"
         Height          =   255
         Left            =   2040
         TabIndex        =   150
         Top             =   4920
         Width           =   1815
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Fuerte Pretoriano"
         Height          =   255
         Left            =   120
         TabIndex        =   149
         Top             =   6120
         Width           =   1815
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Batallon de Ankon"
         Height          =   255
         Left            =   120
         TabIndex        =   148
         Top             =   5760
         Width           =   1815
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Dungeon Marabel"
         Height          =   255
         Left            =   2040
         TabIndex        =   147
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Polo Norte"
         Height          =   255
         Left            =   2040
         TabIndex        =   146
         Top             =   6120
         Width           =   1815
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Minas de Oro"
         Height          =   255
         Left            =   120
         TabIndex        =   145
         Top             =   4920
         Width           =   1815
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Minas de Plata"
         Height          =   255
         Left            =   120
         TabIndex        =   144
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Minas de Hierro"
         Height          =   255
         Left            =   2040
         TabIndex        =   143
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Dungeon Inferno"
         Height          =   255
         Left            =   120
         TabIndex        =   142
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Bundeon Dragon"
         Height          =   255
         Left            =   120
         TabIndex        =   141
         Top             =   4200
         Width           =   1815
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Dungeon Veril"
         Height          =   255
         Left            =   2040
         TabIndex        =   140
         Top             =   4200
         Width           =   1815
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Fortaleza Oeste"
         Height          =   255
         Left            =   120
         TabIndex        =   139
         Top             =   5400
         Width           =   1815
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Bosque Dorck"
         Height          =   255
         Left            =   120
         TabIndex        =   138
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Puente de los Caidos"
         Height          =   255
         Left            =   2040
         TabIndex        =   137
         Top             =   5760
         Width           =   1815
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Bosque Negro"
         Height          =   255
         Left            =   2040
         TabIndex        =   136
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Fortaleza Este"
         Height          =   255
         Left            =   2040
         TabIndex        =   135
         Top             =   5400
         Width           =   1815
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Bosque Arkhein"
         Height          =   255
         Left            =   2040
         TabIndex        =   134
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Bosque Elfico"
         Height          =   255
         Left            =   120
         TabIndex        =   133
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Isla Morgolock"
         Height          =   255
         Left            =   2040
         TabIndex        =   132
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Isla Euclides"
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Isla Veleta"
         Height          =   255
         Left            =   2040
         TabIndex        =   130
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Isla Pirata"
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CommandButton cmdIra1 
         Caption         =   "Isla Victoria"
         Height          =   255
         Left            =   120
         TabIndex        =   127
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton cmdSHOWNAME 
         Caption         =   "SHOWNAME"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   61
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdREM 
         Caption         =   "DEJAR COMENTARIO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   60
         Top             =   1200
         Width           =   3735
      End
      Begin VB.CommandButton cmdINVISIBLE 
         Caption         =   "INVISIBLE"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   59
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdSETDESC 
         Caption         =   "DESCRIPCION"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   58
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdNAVE 
         Caption         =   "NAVEGACION"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   57
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdCHATCOLOR 
         Caption         =   "CHATCOLOR"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdIGNORADO 
         Caption         =   "IGNORADO"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   55
         Top             =   120
         Width           =   1815
      End
      Begin VB.Line Line7 
         X1              =   120
         X2              =   3960
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line Line6 
         X1              =   120
         X2              =   3840
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line5 
         X1              =   240
         X2              =   4080
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   3960
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label10 
         Caption         =   "Telep yo"
         Height          =   255
         Left            =   1680
         TabIndex        =   128
         Top             =   1560
         Width           =   855
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   6
      Left            =   120
      TabIndex        =   54
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdSHOWCMSG 
         Caption         =   "Escuchar a Clan"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   78
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdBANCLAN 
         Caption         =   "/BANNEA AL CLAN"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   77
         Top             =   1440
         Width           =   3495
      End
      Begin VB.CommandButton cmdMIEMBROSCLAN 
         Caption         =   "Mienbros del Clan"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   76
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdBANIPRELOAD 
         Caption         =   "/BANIPRELOAD"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   75
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdBANIPLIST 
         Caption         =   "/BANIPLIST"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   74
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdIP2NICK 
         Caption         =   "BAN X IP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   73
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdBANIP 
         Caption         =   "BAN IP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1440
         TabIndex        =   72
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdUNBANIP 
         Caption         =   "Sacar BAN IP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   71
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   7695
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   360
      Width           =   3855
      Begin VB.CommandButton cmdConsulta 
         Caption         =   "Consulta"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   85
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdNOREAL 
         Caption         =   "Explulsar de la Armada"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   84
         Top             =   6360
         Width           =   1815
      End
      Begin VB.CommandButton cmdNOCAOS 
         Caption         =   "Expulsar de Caos"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   83
         Top             =   6360
         Width           =   1815
      End
      Begin VB.CommandButton cmdKICKCONSE 
         Caption         =   "Degradar Consejero"
         CausesValidation=   0   'False
         Height          =   915
         Left            =   2400
         TabIndex        =   82
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton cmdACEPTCONSECAOS 
         Caption         =   "Ascender a consejero del Caos"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   0
         TabIndex        =   81
         Top             =   7200
         Width           =   2295
      End
      Begin VB.CommandButton cmdACEPTCONSE 
         Caption         =   "Ascender a Consejero Real"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   0
         TabIndex        =   80
         Top             =   6720
         Width           =   2295
      End
      Begin VB.ComboBox cboListaUsus 
         Height          =   315
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   720
         Width           =   3795
      End
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "Actualizar"
         CausesValidation=   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   51
         Top             =   120
         Width           =   3675
      End
      Begin VB.CommandButton cmdIRCERCA 
         Caption         =   "Ir Cerca"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   50
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdDONDE 
         Caption         =   "Ubicar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   49
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdPENAS 
         Caption         =   "Pena"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdTELEP 
         Caption         =   "Mandar User"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   47
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdSILENCIAR 
         Caption         =   "Silenciar"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   1200
         TabIndex        =   46
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdIRA 
         Caption         =   "Ir al User"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   45
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdCARCEL 
         Caption         =   "Carcel"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   44
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdADVERTENCIA 
         Caption         =   "Advertencia"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   43
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdINFO 
         Caption         =   "Informacion"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   42
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdSTAT 
         Caption         =   "Start"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   41
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdBAL 
         Caption         =   "Oro"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   40
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmdINV 
         Caption         =   "Inventario"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   39
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdBOV 
         Caption         =   "Boveda"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   38
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdSKILLS 
         Caption         =   "Skills"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   37
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton cmdREVIVIR 
         Caption         =   "Revivir User"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   36
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdPERDON 
         Caption         =   "Perdonar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   35
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdECHAR 
         Caption         =   "Echar"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   0
         TabIndex        =   34
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdEJECUTAR 
         Caption         =   "Ejecutar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   33
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmdBAN 
         Caption         =   "Bannear"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   32
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton cmdUNBAN 
         Caption         =   "Sacar Ban"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   31
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton cmdSUM 
         Caption         =   "Traer"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   30
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdNICK2IP 
         Caption         =   "Nick del IP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   29
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdESTUPIDO 
         Caption         =   "Estudides al User"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   28
         Top             =   5760
         Width           =   1815
      End
      Begin VB.CommandButton cmdNOESTUPIDO 
         Caption         =   "Sacar la estupides"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1920
         TabIndex        =   27
         Top             =   5760
         Width           =   1815
      End
      Begin VB.CommandButton cmdBORRARPENA 
         Caption         =   "Modificar condena"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   2400
         TabIndex        =   26
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton cmdLASTIP 
         Caption         =   "Ultimo IP"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   25
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdCONDEN 
         Caption         =   "Condenar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   0
         TabIndex        =   24
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdRAJAR 
         Caption         =   "Sacar Faccion"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   23
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdRAJARCLAN 
         Caption         =   "Dejar sin Clan"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   22
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton cmdLASTEMAIL 
         Caption         =   "Ultimo Mail"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   21
         Top             =   2880
         Width           =   1335
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   7695
      Index           =   5
      Left            =   120
      TabIndex        =   53
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton Command6 
         Caption         =   "Mapa sin Invocacion"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   126
         Top             =   7200
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cambiar Triggers"
         Height          =   315
         Left            =   2040
         TabIndex        =   125
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cndPK1 
         Caption         =   "Mapa Inseguro"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   124
         Top             =   5280
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Mapa prohibe Robar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   123
         Top             =   6720
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Mapa con Magia"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   122
         Top             =   6240
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Mapa sin Backup"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   121
         Top             =   5760
         Width           =   1695
      End
      Begin VB.CommandButton cmdBacup 
         Caption         =   "Mapa con BackUp"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   120
         Top             =   5760
         Width           =   1815
      End
      Begin VB.CommandButton cmdMagiaNO 
         Caption         =   "Maspa sin Magia"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   119
         Top             =   6240
         Width           =   1695
      End
      Begin VB.CommandButton cmdRoboSi 
         Caption         =   "Mapa perimite Robar"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   118
         Top             =   6720
         Width           =   1815
      End
      Begin VB.CommandButton cmdInvocaSI 
         Caption         =   "Mapa con Invocacion"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   117
         Top             =   7200
         Width           =   1815
      End
      Begin VB.CommandButton cmdPK0 
         Caption         =   "Mapa Seguro"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   115
         Top             =   5280
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Crear NPCs sin Respawn"
         Height          =   435
         Left            =   240
         TabIndex        =   114
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton cmdNPCsConRespawn 
         Caption         =   "Crear NPC con Respawn"
         Height          =   435
         Left            =   240
         TabIndex        =   113
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CommandButton cmdResetInv 
         Caption         =   "Resetear Inventario"
         Height          =   315
         Left            =   240
         TabIndex        =   112
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdMataconRepawn 
         Caption         =   "      Matar criatura       deja respawn"
         Height          =   435
         Left            =   2040
         TabIndex        =   111
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CommandButton cmdMata 
         Caption         =   "Matar criatura"
         Height          =   435
         Left            =   2040
         TabIndex        =   110
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton cmddestBloq 
         Caption         =   "Quitar/Poner Bloqueo"
         Height          =   315
         Left            =   240
         TabIndex        =   109
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cmdDE 
         Caption         =   "Destruir exit"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   108
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdCC 
         Caption         =   "Crear NPCs"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   240
         TabIndex        =   70
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton cmdLIMPIAR 
         Caption         =   "Limpiar Mundo"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   69
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdCT 
         Caption         =   "Crear Telepor"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   68
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdDT 
         Caption         =   "Destruir Teleport"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   67
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdLLUVIA 
         Caption         =   "Lluvia - Si / No"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   66
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdMASSDEST 
         Caption         =   "Dest Item en Mapa"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   65
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cmdPISO 
         Caption         =   "Informe del Piso"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   64
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cmdCI 
         Caption         =   "Crear Item"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   63
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton cmdDEST 
         Caption         =   "Destruir item"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   62
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   3720
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label lblMapa 
         Caption         =   "Modificar opciones del Mapa"
         Height          =   375
         Left            =   840
         TabIndex        =   116
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   3960
         Y1              =   4800
         Y2              =   4800
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   7395
      Index           =   7
      Left            =   120
      TabIndex        =   86
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ACTUALIZAR"
         Height          =   495
         Left            =   2160
         TabIndex        =   107
         Top             =   2100
         Width           =   1695
      End
      Begin VB.TextBox txtNuevaDescrip 
         Height          =   765
         Left            =   120
         MaxLength       =   40
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   105
         Top             =   6120
         Width           =   3735
      End
      Begin VB.CommandButton cmdAddFollow 
         Caption         =   "Agregar Seguimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   103
         Top             =   6960
         Width           =   3735
      End
      Begin VB.TextBox txtNuevoUsuario 
         Height          =   285
         Left            =   120
         TabIndex        =   102
         Top             =   5580
         Width           =   3735
      End
      Begin VB.CommandButton cmdAddObs 
         Caption         =   "Agregar Observacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   100
         Top             =   4800
         Width           =   3735
      End
      Begin VB.TextBox txtObs 
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   99
         Top             =   3780
         Width           =   3735
      End
      Begin VB.TextBox txtDescrip 
         Height          =   675
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   40
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   97
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox txtCreador 
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   1620
         Width           =   1695
      End
      Begin VB.TextBox txtTimeOn 
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtIP 
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   540
         Width           =   1695
      End
      Begin VB.ListBox lstUsers 
         Height          =   2400
         Left            =   120
         TabIndex        =   87
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   106
         Top             =   60
         Width           =   660
      End
      Begin VB.Label Label9 
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4200
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Label Label8 
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   101
         Top             =   5340
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   98
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   96
         Top             =   2700
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Creador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   94
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Logueado Hace:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   92
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "IP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   90
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblEstado 
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Online"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   2880
         TabIndex        =   89
         Top             =   60
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Usuarios Marcados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdGMSG 
         Caption         =   "Mensaje Consola"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdHORA 
         Caption         =   "Hora"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdRMSG 
         Caption         =   "     Mensaje      Rol Master"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   1320
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdREALMSG 
         Caption         =   "Mensajes a Reales"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdCAOSMSG 
         Caption         =   "Mensajes a Caos"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   1320
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdCIUMSG 
         Caption         =   "Mensajes a Ciudadanos"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   2640
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdTALKAS 
         Caption         =   "Hablar por NPC"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   2640
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdMOTDCAMBIA 
         Caption         =   "Carbiar Motd"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   1320
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdSMSG 
         Caption         =   "Mensaje por Sistema"
         CausesValidation=   0   'False
         Height          =   555
         Left            =   2640
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3975
      Begin VB.CommandButton cmdONLINEREAL 
         Caption         =   "Reales Online"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdONLINECAOS 
         Caption         =   "Caos Online"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdOCULTANDO 
         Caption         =   "Ocultos"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdONLINEGM 
         Caption         =   "GMs Online"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CommandButton cmdBORRAR_SOS 
         Caption         =   "Borrar S:O:S"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdTRABAJANDO 
         Caption         =   "Trabajando"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   16
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdSHOW_SOS 
         Caption         =   "Ver S.O.S"
         CausesValidation=   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   17
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSComctlLib.TabStrip TabStrip 
      CausesValidation=   0   'False
      Height          =   1935
      Left            =   120
      TabIndex        =   79
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3413
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Message"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Player"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Me"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "World"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Admin"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seguimientos"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSeguimientos 
      Caption         =   "Seguimientos"
      Begin VB.Menu mnuIra 
         Caption         =   "Ir Cerca"
      End
      Begin VB.Menu mnuSum 
         Caption         =   "Sumonear"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Eliminar Seguimiento"
      End
   End
   Begin VB.Menu PEventos 
      Caption         =   "Eventos"
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmPanelGm.frm
'
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

''
' IMPORTANT!!!
' To prevent the combo list of usernames from closing when a conole message arrives, the Validate event allways
' sets the Cancel arg to True. This, combined with setting the CausesValidation of the RichTextBox to True
' makes the trick. However, in order to be able to use other commands, ALL OTHER controls in this form must have the
' CuasesValidation parameter set to false (unless you want to code your custom flag system to know when to allow or not the loose of focus).

Private Sub cboListaUsus_Validate(Cancel As Boolean)
    Cancel = True

End Sub

Private Sub cmdACEPTCONSE_Click()

    '/ACEPTCONSE
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea aceptar a " & Nick & " como consejero real?", vbYesNo, "Atencion!") = vbYes Then Call WriteAcceptRoyalCouncilMember(Nick)
    frmMain.Show

End Sub

Private Sub cmdACEPTCONSECAOS_Click()

    '/ACEPTCONSECAOS
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea aceptar a " & Nick & " como consejero del caos?", vbYesNo, "Atencion!") = vbYes Then Call WriteAcceptChaosCouncilMember(Nick)
    frmMain.Show

End Sub

Private Sub cmdAddFollow_Click()

    'Dim i As Long
    '
    '    For i = 0 To lstUsers.ListCount
    '        If UCase$(lstUsers.List(i)) = UCase$(txtNuevoUsuario.Text) Then
    '            Call MsgBox("El usuario ya esta en la lista!", vbOKOnly + vbExclamation)
    '            Exit Sub
    '        End If
    '    Next i
    '
    '    If LenB(txtNuevoUsuario.Text) = 0 Then
    '        Call MsgBox("Escribe el nombre de un usuario!", vbOKOnly + vbExclamation)
    '        Exit Sub
    '    End If
    '
    '    If LenB(txtNuevaDescrip.Text) = 0 Then
    '        Call MsgBox("Escribe el motivo del seguimiento!", vbOKOnly + vbExclamation)
    '        Exit Sub
    '    End If
    '
    '    Call WriteRecordAdd(txtNuevoUsuario.Text, txtNuevaDescrip.Text)
    '
    '    txtNuevoUsuario.Text = vbNullString
    '    txtNuevaDescrip.Text = vbNullString
End Sub

Private Sub cmdAddObs_Click()

    'Dim Obs As String
    '
    '    Obs = InputBox("Ingrese la observacion", "Nueva Observacion")
    '
    '    If LenB(Obs) = 0 Then
    '        Call MsgBox("Escribe una observacion!", vbOKOnly + vbExclamation)
    '        Exit Sub
    '    End If
    '
    '    If lstUsers.ListIndex = -1 Then
    '        Call MsgBox("Seleccione un seguimiento!", vbOKOnly + vbExclamation)
    '        Exit Sub
    '    End If
    '
    '    Call WriteRecordAddObs(lstUsers.ListIndex + 1, Obs)
End Sub

Private Sub cmdADVERTENCIA_Click()

    '/ADVERTENCIA
    Dim tStr As String

    Dim Nick As String

    Nick = cboListaUsus.Text
        
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Escriba el motivo de la advertencia.", "Advertir a " & Nick)
                
        If LenB(tStr) <> 0 Then
            'We use the Parser to control the command format
            Call ParseUserCommand("/ADVERTENCIA " & Nick & "@" & tStr)

        End If

    End If

    frmMain.Show

End Sub

Private Sub cmdBackUPNo_Click()

    '/BACKUP NO Agregado Por ReyarB 21/05/2020
    If MsgBox("Seguro desea hacer que el Mapa NO tenga Backup hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/MODMAPINFO BACKUP 0") 'We use the Parser to control the command format

End Sub

Private Sub cmdBacup_Click()

    '/BACKUP SI Agregado Por ReyarB 21/05/2020
    If MsgBox("Seguro desea hacer que el Mapa tenga Backup hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/MODMAPINFO BACKUP 1") 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub cmdBAL_Click()

    '/BAL
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharGold(Nick)
    frmMain.Show

End Sub

Private Sub cmdBAN_Click()

    '/BAN
    Dim tStr As String

    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Escriba el motivo del ban.", "BAN a " & Nick)
                
        If LenB(tStr) <> 0 Then If MsgBox("Seguro desea banear a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteBanChar(Nick, tStr)

    End If

    frmMain.Show

End Sub

Private Sub cmdBANCLAN_Click()

    '/BANCLAN
    Dim tStr As String
    
    tStr = InputBox("Escriba el nombre del clan.", "Banear clan")

    If LenB(tStr) <> 0 Then If MsgBox("Seguro desea banear al clan " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteGuildBan(tStr)
    frmMain.Show

End Sub

Private Sub cmdBANIP_Click()

    '/BANIP
    Dim tStr   As String

    Dim reason As String
    
    tStr = InputBox("Escriba el ip o el nick del PJ.", "Banear IP")
    
    reason = InputBox("Escriba el motivo del ban.", "Banear IP")
    
    If LenB(tStr) <> 0 Then If MsgBox("Seguro desea banear la ip " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/BANIP " & tStr & " " & reason) 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub cmdBANIPLIST_Click()
    '/BANIPLIST
    Call WriteBannedIPList
    frmMain.Show

End Sub

Private Sub cmdBANIPRELOAD_Click()
    '/BANIPRELOAD
    Call WriteBannedIPReload
    frmMain.Show

End Sub

Private Sub cmdBORRAR_SOS_Click()

    '/BORRAR SOS
    If MsgBox("Seguro desea borrar el SOS?", vbYesNo, "Atencion!") = vbYes Then Call WriteCleanSOS
    frmMain.Show

End Sub

Private Sub cmdBORRARPENA_Click()

    '/BORRARPENA
    Dim tStr As String

    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Indique el numero de la pena a borrar.", "Borrar pena")

        If LenB(tStr) <> 0 Then If MsgBox("Seguro desea borrar la pena " & tStr & " a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/BORRARPENA " & Nick & "@" & tStr) 'We use the Parser to control the command format

    End If

    frmMain.Show

End Sub

Private Sub cmdBOV_Click()

    '/BOV
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharBank(Nick)
    frmMain.Show

End Sub

Private Sub cmdCAOSMSG_Click()

    '/CAOSMSG
    Dim tStr As String
    
    tStr = InputBox("Ingrese el TEXTO", "Mensaje por consola LegionOscura")

    If LenB(tStr) <> 0 Then Call WriteChaosLegionMessage(tStr)

End Sub

Private Sub cmdCARCEL_Click()

    '/CARCEL
    Dim tStr As String

    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Escriba el motivo de la pena.", "Carcel a " & Nick)
                
        If LenB(tStr) <> 0 Then
            tStr = tStr & "@" & InputBox("Indique el tiempo de condena (entre 0 y 60 minutos).", "Carcel a " & Nick)
            'We use the Parser to control the command format
            Call ParseUserCommand("/CARCEL " & Nick & "@" & tStr)

        End If

    End If

    frmMain.Show

End Sub

Private Sub cmdCC_Click()
    '/CC
    Call WriteSpawnListRequest
    frmMain.Show

End Sub

Private Sub cmdCHATCOLOR_Click()

    '/CHATCOLOR
    Dim tStr As String

    tStr = InputBox("Defina el color (R G B). Deje en blanco para usar el default.", "Cambiar color del chat")
    Call ParseUserCommand("/CHATCOLOR " & tStr) 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub cmdCI_Click()

    '/CI
    Dim tStr As String
    
    tStr = InputBox("Indique el numero del objeto a crear y la cantidad.", "Crear Objeto")

    If LenB(tStr) <> 0 Then If MsgBox("Seguro desea crear el objeto " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/CI " & tStr) 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub cmdCIUMSG_Click()

    '/CIUMSG
    Dim tStr As String
    
    tStr = InputBox(" Ingrese el TEXTO", "Mensaje por consola Ciudadanos")

    If LenB(tStr) <> 0 Then Call WriteCitizenMessage(tStr)

End Sub

Private Sub cmdCONDEN_Click()

    '/CONDEN
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea volver criminal a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteTurnCriminal(Nick)
    frmMain.Show

End Sub

Private Sub cmdConsulta_Click()

    '    WriteConsultation
    '    frmMain.Show
End Sub

Private Sub cmdCT_Click()

    '/CT
    Dim tStr As String
    
    tStr = InputBox("Indique la posicion donde lleva el portal (MAPA X Y).", "Crear Portal")

    If LenB(tStr) <> 0 Then Call ParseUserCommand("/CT " & tStr) 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub cmdDE_Click()

    ''/DE
    '    If MsgBox("Seguro desea destruir el Tile Exit?", vbYesNo, "Atencion!") = vbYes Then _
    '        Call WriteExitDestroy
    '        frmMain.Show
End Sub

Private Sub cmdDEST_Click()

    '/DEST
    If MsgBox("Seguro desea destruir el objeto sobre el que esta parado?", vbYesNo, "Atencion!") = vbYes Then Call WriteDestroyItems
    frmMain.Show

End Sub

Private Sub cmddestBloq_Click()

    If MsgBox("Seguro desea el bloqueo en su ubicaci??n ? ", vbYesNo, "Atencion!") = vbYes Then Call WriteTileBlockedToggle
    frmMain.Show

End Sub

Private Sub cmdDONDE_Click()

    '/DONDE
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteWhere(Nick)
    frmMain.Show

End Sub

Private Sub cmdDT_Click()

    '    'DT
    '    If MsgBox("Seguro desea destruir el portal?", vbYesNo, "Atencion!") = vbYes Then _
    '        Call WriteTeleportDestroy
    '        Call WriteExitDestroy
    '        frmMain.Show
End Sub

Private Sub cmdECHAR_Click()

    '/ECHAR
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteKick(Nick)
    frmMain.Show

End Sub

Private Sub cmdEJECUTAR_Click()

    '/EJECUTAR
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea ejecutar a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteExecute(Nick)
    frmMain.Show

End Sub

Private Sub cmdESTUPIDO_Click()

    '/ESTUPIDO
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteMakeDumb(Nick)
    frmMain.Show

End Sub

Private Sub cmdGMSG_Click()

    '/GMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de GM")

    If LenB(tStr) <> 0 Then Call WriteGMMessage(tStr)

End Sub

Private Sub cmdHORA_Click()
    '/HORA
    Call Protocol.WriteServerTime

End Sub

Private Sub cmdIGNORADO_Click()
    '/IGNORADO
    Call WriteIgnored
    frmMain.Show

End Sub

Private Sub cmdINFO_Click()

    '/INFO
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharInfo(Nick)
    frmMain.Show

End Sub

Private Sub cmdINV_Click()

    '/INV
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharInventory(Nick)
    frmMain.Show

End Sub

Private Sub cmdINVISIBLE_Click()
    '/INVISIBLE
    Call WriteInvisible
    frmMain.Show

End Sub

Private Sub cmdInvocaNO_Click()

    '/Sin Invocaci??n Agregado Por ReyarB 21/05/2020
    If MsgBox("Seguro desea hacer que el Mapa NO puedan Invocar hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/MODMAPINFO INVOCARSINEFECTO 0") 'We use the Parser to control the command format

End Sub

Private Sub cmdInvocaSI_Click()

    '/Con Invocaci??n Agregado Por ReyarB 21/05/2020
    If MsgBox("Seguro desea hacer que el Mapa puedan Invocar hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/MODMAPINFO INVOCARSINEFECTO 1") 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub cmdIP2NICK_Click()

    '/IP2NICK
    Dim tStr As String
    
    tStr = InputBox("Escriba la ip.", "IP to Nick")

    If LenB(tStr) <> 0 Then Call ParseUserCommand("/IP2NICK " & tStr) 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub cmdIRA_Click()

    '/IRA
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteGoToChar(Nick)
    frmMain.Show

End Sub

Private Sub cmdIra1_Click()
    '/Telep yo Isla victoria

    Call ParseUserCommand("/TELEP YO 1 926 627") 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub cmdIRCERCA_Click()

    '/IRCERCA
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteGoNearby(Nick)
    frmMain.Show

End Sub

Private Sub cmdKICKCONSE_Click()

    'KICKCONSE
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea destituir a " & Nick & " de su cargo de consejero?", vbYesNo, "Atencion!") = vbYes Then Call WriteCouncilKick(Nick)
    frmMain.Show

End Sub

Private Sub cmdLASTEMAIL_Click()

    '/LASTEMAIL
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharMail(Nick)
    frmMain.Show

End Sub

Private Sub cmdLASTIP_Click()

    '/LASTIP
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteLastIP(Nick)
    frmMain.Show

End Sub

Private Sub cmdLIMPIAR_Click()

    '    '/LIMPIARMUNDO
    '    Call WriteLimpiarMundo
    '    frmMain.Show
End Sub

Private Sub cmdLLUVIA_Click()
    '/LLUVIA
    Call WriteRainToggle
    frmMain.Show

End Sub

Private Sub cmdMagiaNO_Click()

    '/Sin Magia Agregado Por ReyarB 21/05/2020
    If MsgBox("Seguro desea hacer que el Mapa NO tenga Magia hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/MODMAPINFO MAGIASINEFECTO 0") 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub cmdMagiaSI_Click()

    '/Con Magia Agregado Por ReyarB 21/05/2020
    If MsgBox("Seguro desea hacer que el Mapa tenga Magia hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/MODMAPINFO MAGIASINEFECTO 1") 'We use the Parser to control the command format

End Sub

Private Sub cmdMASSDEST_Click()

    '/MASSDEST
    If MsgBox("Seguro desea destruir todos los items a la vista?", vbYesNo, "Atencion!") = vbYes Then Call WriteDestroyAllItemsInArea
    frmMain.Show

End Sub

Private Sub cmdMata_Click()
    Call WriteKillNPCNoRespawn
    frmMain.Show

End Sub

Private Sub cmdMataconRepawn_Click()
    Call WriteKillNPC
    frmMain.Show

End Sub

Private Sub cmdMIEMBROSCLAN_Click()

    '/MIEMBROSCLAN
    Dim tStr As String
    
    tStr = InputBox("Escriba el nombre del clan.", "Lista de miembros del clan")

    If LenB(tStr) <> 0 Then Call WriteGuildMemberList(tStr)
    frmMain.Show

End Sub

Private Sub cmdMOTDCAMBIA_Click()
    '/MOTDCAMBIA
    Call WriteChangeMOTD

End Sub

Private Sub cmdNAVE_Click()
    '/NAVE
    Call WriteNavigateToggle
    frmMain.Show

End Sub

Private Sub cmdNENE_Click()

    '    '/NENE
    '    Dim tStr As String
    '
    '    tStr = InputBox("Indique el mapa.", "Numero de NPCs enemigos.")
    '    If LenB(tStr) <> 0 Then _
    '        Call ParseUserCommand("/NENE " & tStr) 'We use the Parser to control the command format
    '        frmMain.Show
End Sub

Private Sub cmdNICK2IP_Click()

    '/NICK2IP
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteNickToIP(Nick)
    frmMain.Show

End Sub

Private Sub cmdNOCAOS_Click()

    '/NOCAOS
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea expulsar a " & Nick & " de la legion oscura?", vbYesNo, "Atencion!") = vbYes Then Call WriteChaosLegionKick(Nick)
    frmMain.Show

End Sub

Private Sub cmdNOESTUPIDO_Click()

    '/NOESTUPIDO
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteMakeDumbNoMore(Nick)
    frmMain.Show

End Sub

Private Sub cmdNOREAL_Click()

    '/NOREAL
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea expulsar a " & Nick & " de la armada real?", vbYesNo, "Atencion!") = vbYes Then Call WriteRoyalArmyKick(Nick)
    frmMain.Show

End Sub

Private Sub cmdNPCsConRespawn_Click()

    '/RACC
    Dim tStr As String
    
    tStr = InputBox("Indique el numero del NPC a crear.", "Crear NPC con Respawn")

    If LenB(tStr) <> 0 Then If MsgBox("Seguro desea crear el NPC " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/RACC " & tStr) 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub cmdOCULTANDO_Click()
    '/OCULTANDO
    Call WriteHiding
    frmMain.Show

End Sub

Private Sub cmdONLINECAOS_Click()
    '/ONLINECAOS
    Call WriteOnlineChaosLegion
    frmMain.Show

End Sub

Private Sub cmdONLINEGM_Click()
    '/ONLINEGM
    Call WriteOnlineGM
    frmMain.Show

End Sub

Private Sub cmdONLINEMAP_Click()

    '    '/ONLINEMAP
    '    Call WriteOnlineMap(UserMap)
    '    frmMain.Show
End Sub

Private Sub cmdONLINEREAL_Click()
    '/ONLINEREAL
    Call WriteOnlineRoyalArmy
    frmMain.Show

End Sub

Private Sub cmdPENAS_Click()

    '/PENAS
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WritePunishments(Nick)
    frmMain.Show

End Sub

Private Sub cmdPERDON_Click()

    '/PERDON
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteForgive(Nick)
    frmMain.Show

End Sub

Private Sub cmdPISO_Click()
    '/PISO
    Call WriteItemsInTheFloor
    frmMain.Show

End Sub

Private Sub cmdPK0_Click()

    '/PK Seguro Agregado Por ReyarB 21/05/2020
    If MsgBox("Seguro desea hacer el Mapa Seguro?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/MODMAPINFO PK 0") 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub cmdRAJAR_Click()

    '/RAJAR
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea resetear la faccion de " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteResetFactions(Nick)
    frmMain.Show

End Sub

Private Sub cmdRAJARCLAN_Click()

    '/RAJARCLAN
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea expulsar a " & Nick & " de su clan?", vbYesNo, "Atencion!") = vbYes Then Call WriteRemoveCharFromGuild(Nick)
    frmMain.Show

End Sub

Private Sub cmdREALMSG_Click()

    '/REALMSG
    Dim tStr As String
    
    tStr = InputBox("Ingrese el TEXTO", "Mensaje por consola ArmadaReal")

    If LenB(tStr) <> 0 Then Call WriteRoyalArmyMessage(tStr)

End Sub

Private Sub cmdRefresh_Click()
    Call ClearRecordDetails

    'Call WriteRecordListRequest
End Sub

Private Sub cmdREM_Click()

    '/REM
    Dim tStr As String
    
    tStr = InputBox("Escriba el comentario.", "Comentario en el logGM")

    If LenB(tStr) <> 0 Then Call WriteComment(tStr)
    frmMain.Show

End Sub

Private Sub cmdResetInv_Click()
    Call WriteResetNPCInventory
    frmMain.Show

End Sub

Private Sub cmdREVIVIR_Click()

    '/REVIVIR
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteReviveChar(Nick)

    'frmMain.Show
End Sub

Private Sub cmdRMSG_Click()

    '/RMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba el mensaje.", "Mensaje por consola de RoleMaster")

    If LenB(tStr) <> 0 Then Call WriteServerMessage(tStr)

End Sub

Private Sub cmdRoboNO_Click()

    '/NO Robor Agregado Por ReyarB 21/05/2020
    If MsgBox("Seguro desea hacer que el Mapa NO puedan Robar hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/MODMAPINFO ROBONPC 0") 'We use the Parser to control the command format

End Sub

Private Sub cmdRoboSi_Click()

    '/Con Robo Agregado Por ReyarB 21/05/2020
    If MsgBox("Seguro desea hacer que el Mapa puedan Robar hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/MODMAPINFO ROBONPC 1") 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub cmdSETDESC_Click()

    '/SETDESC
    Dim tStr As String
    
    tStr = InputBox("Escriba una DESC.", "Set Description")

    If LenB(tStr) <> 0 Then Call WriteSetCharDescription(tStr)
    frmMain.Show

End Sub

Private Sub cmdSHOW_SOS_Click()
    '/SHOW SOS
    Call WriteSOSShowList
    frmMain.Show

End Sub

Private Sub cmdSHOWCMSG_Click()

    '/SHOWCMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba el nombre del clan que desea escuchar.", "Escuchar los mensajes del clan")

    If LenB(tStr) <> 0 Then Call WriteShowGuildMessages(tStr)
    frmMain.Show

End Sub

Private Sub cmdSHOWNAME_Click()

    '    '/SHOWNAME
    '    Call WriteShowName
    '    frmMain.Show
End Sub

Private Sub cmdSILENCIAR_Click()

    '/SILENCIAR
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteSilence(Nick)
    frmMain.Show

End Sub

Private Sub cmdSKILLS_Click()

    '/SKILLS
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharSkills(Nick)
    frmMain.Show

End Sub

Private Sub cmdSMSG_Click()

    '/SMSG
    Dim tStr As String
    
    tStr = InputBox("Escriba el mensaje.", "Mensaje de sistema")

    If LenB(tStr) <> 0 Then Call WriteSystemMessage(tStr)

End Sub

Private Sub cmdSTAT_Click()

    '/STAT
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteRequestCharStats(Nick)
    frmMain.Show

End Sub

Private Sub cmdSUM_Click()

    '/SUM
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then Call WriteSummonChar(Nick)
    frmMain.Show

End Sub

Private Sub cmdTALKAS_Click()

    '/TALKAS
    Dim tStr As String

    tStr = InputBox("Escriba el mensaje", "Hablar por NPC")

    If LenB(tStr) <> 0 Then Call WriteTalkAsNPC(tStr)

End Sub

Private Sub cmdTELEP_Click()

    '/TELEP
    Dim tStr As String

    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Indique la posicion (MAPA X Y).", "Transportar a " & Nick)

        If LenB(tStr) <> 0 Then Call ParseUserCommand("/TELEP " & Nick & " " & tStr) 'We use the Parser to control the command format

    End If

    frmMain.Show

End Sub

Private Sub cmdTRABAJANDO_Click()
    '/TRABAJANDO
    Call WriteWorking
    frmMain.Show

End Sub

Private Sub cmdUNBAN_Click()

    '/UNBAN
    Dim Nick As String

    Nick = cboListaUsus.Text
    
    If LenB(Nick) <> 0 Then If MsgBox("Seguro desea unbanear a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then Call WriteUnbanChar(Nick)
    frmMain.Show

End Sub

Private Sub cmdUNBANIP_Click()

    '/UNBANIP
    Dim tStr As String
    
    tStr = InputBox("Escriba el ip.", "Unbanear IP")

    If LenB(tStr) <> 0 Then If MsgBox("Seguro desea unbanear la ip " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/UNBANIP " & tStr) 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub cndPK1_Click()

    '/PK Inseguro
    If MsgBox("Seguro desea hacer el Mapa Inseguro?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/MODMAPINFO PK 1") 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub Command1_Click()

    '/BACKUP no Agregado Por ReyarB 21/05/2020
    If MsgBox("Seguro desea hacer que el Mapa NO tenga Backup hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/MODMAPINFO BACKUP 0") 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub Command10_Click()
    '/Telep yo Isla Morgolock

    Call ParseUserCommand("/TELEP YO 1 795 522") 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub Command11_Click()
    '/Telep yo Bosque Elfico

    Call ParseUserCommand("/TELEP YO 1 105 563") 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub Command12_Click()
    '/Telep yo Bosque Arkhein

    Call ParseUserCommand("/TELEP YO 1 471 1310")
    frmMain.Show

End Sub

Private Sub Command13_Click()
    '/Telep yo Fortaleza Este

    Call ParseUserCommand("/TELEP YO 1 1057 1428")
    frmMain.Show

End Sub

Private Sub Command14_Click()
    '/Telep yo Bosque Negro

    Call ParseUserCommand("/TELEP YO 1 177 704")
    frmMain.Show

End Sub

Private Sub Command15_Click()
    '/Telep yo Puente de los Caidos

    Call ParseUserCommand("/TELEP YO 1 545 92")
    frmMain.Show

End Sub

Private Sub Command16_Click()
    '/Telep yo Bosque Dorck

    Call ParseUserCommand("/TELEP YO 1 233 933")
    frmMain.Show

End Sub

Private Sub Command17_Click()
    '/Telep yo Fortaleza Oeste

    Call ParseUserCommand("/TELEP YO 1 42 1465")
    frmMain.Show

End Sub

Private Sub Command19_Click()
    '/Telep yo Dungeon Veril

    Call ParseUserCommand("/TELEP YO 1 69 754")
    frmMain.Show

End Sub

Private Sub Command2_Click()

    '/ACC
    Dim tStr As String
    
    tStr = InputBox("Indique el numero del NPC a crear.", "Crear NPC con Respawn")

    If LenB(tStr) <> 0 Then If MsgBox("Seguro desea crear el NPC " & tStr & "?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/ACC " & tStr) 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub Command20_Click()
    '/Telep yo Dungeon Dragon

    Call ParseUserCommand("/TELEP YO 1 602 1096")
    frmMain.Show

End Sub

Private Sub Command21_Click()
    '/Telep yo Dungeon Inferno

    Call ParseUserCommand("/TELEP YO 1 900 402")
    frmMain.Show

End Sub

Private Sub Command22_Click()
    '/Telep yo Minas Hierro

    Call ParseUserCommand("/TELEP YO 1 373 855")
    frmMain.Show

End Sub

Private Sub Command23_Click()
    '/Telep yo Minas Plata

    Call ParseUserCommand("/TELEP YO 1 173 968")
    frmMain.Show

End Sub

Private Sub Command24_Click()
    '/Telep yo Minas Oro

    Call ParseUserCommand("/TELEP YO 2 257 309")
    frmMain.Show

End Sub

Private Sub Command25_Click()
    '/Telep yo Polo Norte

    Call ParseUserCommand("/TELEP YO 1 993 123")
    frmMain.Show

End Sub

Private Sub Command26_Click()
    '/Telep yo Isla Pirata

    Call ParseUserCommand("/TELEP YO 1 434 624")
    frmMain.Show

End Sub

Private Sub Command27_Click()
    '/Telep yo Batallon alkon

    Call ParseUserCommand("/TELEP YO 1 656 184")
    frmMain.Show

End Sub

Private Sub Command28_Click()
    '/Telep yo Fuerte pretoriano

    Call ParseUserCommand("/TELEP YO 1 410 256")
    frmMain.Show

End Sub

Private Sub Command29_Click()
    '/Telep yo Muelle Bander

    Call ParseUserCommand("/TELEP YO 1 301 60")
    frmMain.Show

End Sub

Private Sub Command3_Click()

    '/Con Magia Agregado Por ReyarB 21/05/2020
    If MsgBox("Seguro desea hacer que el Mapa tenga Magia hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/MODMAPINFO MAGIASINEFECTO 1") 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub Command30_Click()
    '/Telep yo Muelle Arghal

    Call ParseUserCommand("/TELEP YO 1 800 315")
    frmMain.Show

End Sub

Private Sub Command31_Click()
    '/Telep yo muelle lindos oeste

    Call ParseUserCommand("/TELEP YO 1 894 995")
    frmMain.Show

End Sub

Private Sub Command32_Click()
    '/Telep yo muelle Arkeig

    Call ParseUserCommand("/TELEP YO 1 639 1370")
    frmMain.Show

End Sub

Private Sub Command33_Click()
    '/Telep yo Desierto

    Call ParseUserCommand("/TELEP YO 1 724 869")
    frmMain.Show

End Sub

Private Sub Command34_Click()
    '/Telep yo Muelle de Nix

    Call ParseUserCommand("/TELEP YO 1 161 1243")
    frmMain.Show

End Sub

Private Sub Command4_Click()

    '/Sin Robo Agregado Por ReyarB 21/05/2020
    If MsgBox("Seguro desea hacer que el Mapa NO puedan Robar hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/MODMAPINFO ROBONPC 0") 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub Command5_Click()

    Dim tStr As String
    
    tStr = InputBox("Indique el numero del Trigger??s a cambiar.", "Crear NPC con Respawn")

    If LenB(tStr) <> 0 Then If MsgBox("Seguro desea cambiar a " & tStr & " el trigger donde esta parado?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/trigger " & tStr) 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub Command6_Click()

    '/Sin Invocaci??n Agregado Por ReyarB 21/05/2020
    If MsgBox("Seguro desea hacer que el Mapa NO puedan Invocar hasta un reset.?", vbYesNo, "Atencion!") = vbYes Then Call ParseUserCommand("/MODMAPINFO INVOCARSINEFECTO 0") 'We use the Parser to control the command format
    frmMain.Show

End Sub

Private Sub Command7_Click()
    '/Telep yo Isla Pirata

    Call ParseUserCommand("/TELEP YO 1 192 47")
    frmMain.Show

End Sub

Private Sub Command8_Click()
    '/Telep yo Isla Veleta

    Call ParseUserCommand("/TELEP YO 1 594 33")

End Sub

Private Sub Command9_Click()
    '/Telep yo Isla Euclides

    Call ParseUserCommand("/TELEP YO 1 120 1447")

End Sub

Private Sub Form_Load()
    Call showTab(1)
    
    'Actualiza los usuarios online
    Call cmdActualiza_Click
    
    'Actualiza los seguimientos
    Call cmdRefresh_Click
    
    'Oculta el menu usado para el PopUp
    mnuSeguimientos.Visible = False

End Sub

Private Sub cmdActualiza_Click()
    Call WriteRequestUserList
    Call FlushBuffer

End Sub

Private Sub cmdCerrar_Click()
    Unload Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me

End Sub

Private Sub lstUsers_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)

    '    If Button = vbRightButton Then
    '        PopupMenu mnuSeguimientos
    '    Else
    '        If lstUsers.ListIndex <> -1 Then
    '            Call ClearRecordDetails
    '            Call WriteRecordDetailsRequest(lstUsers.ListIndex + 1)
    '        End If
    '    End If
End Sub

Private Sub ClearRecordDetails()
    txtIP.Text = vbNullString
    txtCreador.Text = vbNullString
    txtDescrip.Text = vbNullString
    txtObs.Text = vbNullString
    txtTimeOn.Text = vbNullString
    lblEstado.Caption = vbNullString

End Sub

Private Sub mnuDelete_Click()

    '    With lstUsers
    '        If .ListIndex = -1 Then
    '            Call MsgBox("Seleccione un usuario para remover el seguimiento!", vbOKOnly + vbExclamation)
    '            Exit Sub
    '        End If
    '
    '        If MsgBox("Desea eliminar el seguimiento al personaje " & .List(.ListIndex) & "?", vbYesNo) = vbYes Then
    '            Call WriteRecordRemove(.ListIndex + 1)
    '            Call ClearRecordDetails
    '        End If
    '    End With
End Sub

Private Sub mnuIra_Click()

    With lstUsers

        If .ListIndex <> -1 Then
            Call WriteGoToChar(.List(.ListIndex))

        End If

    End With

End Sub

Private Sub mnuSum_Click()

    With lstUsers

        If .ListIndex <> -1 Then
            Call WriteSummonChar(.List(.ListIndex))

        End If

    End With

End Sub

Private Sub PEventos_Click()
    frmPanelGm.Hide
    frmPanelTorneo.Show vbModal

End Sub

Private Sub TabStrip_Click()
    Call showTab(TabStrip.SelectedItem.Index)

End Sub

Private Sub showTab(TabId As Byte)

    Dim I As Byte
    
    For I = 1 To Frame.UBound
        Frame(I).Visible = (I = TabId)
    Next I
    
    With Frame(TabId)
        frmPanelGm.Height = .Height + 1280
        TabStrip.Height = .Height + 480
        cmdCerrar.Top = .Height + 480

    End With

End Sub

