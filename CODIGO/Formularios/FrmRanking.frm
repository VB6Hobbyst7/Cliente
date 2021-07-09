VERSION 5.00
Begin VB.Form FrmRanking 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   Picture         =   "FrmRanking.frx":0000
   ScaleHeight     =   4455
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer aniTimer 
      Enabled         =   0   'False
      Interval        =   23
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   135
   End
   Begin VB.Image ImgOro 
      Height          =   375
      Left            =   360
      MouseIcon       =   "FrmRanking.frx":5917
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Image ImgFrags 
      Height          =   375
      Left            =   360
      MouseIcon       =   "FrmRanking.frx":65E1
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Image ImgReto 
      Height          =   375
      Left            =   360
      MouseIcon       =   "FrmRanking.frx":72AB
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Image ImgNivel 
      Height          =   375
      Left            =   360
      MouseIcon       =   "FrmRanking.frx":7F75
      MousePointer    =   99  'Custom
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Image ImgTorneo 
      Height          =   375
      Left            =   360
      MouseIcon       =   "FrmRanking.frx":8C3F
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   2295
   End
   Begin VB.Image ImgClan 
      Height          =   255
      Left            =   360
      MouseIcon       =   "FrmRanking.frx":9909
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   2295
   End
End
Attribute VB_Name = "FrmRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type GIFAnimator

    Frames As Long
    Frame As Long
    LoopCount As Long
    Intervals() As Long

End Type

Private myAnimator As GIFAnimator






Private Sub Image1_Click()
Unload Me
frmMain.SetFocus
End Sub




Private Sub aniTimer_Timer()
On Error GoTo err

    aniTimer.Enabled = False

    With myAnimator

        If .Frame = .Frames Then        ' loop occurred
            ' intervals(0) is number of loops before stopping animation. values < 1 indicate infinite looping
            .Frame = 0

            If .Intervals(0) > 0 Then
                .LoopCount = .LoopCount + 1

                If .LoopCount = .Intervals(0) Then
                    .LoopCount = 0 ' if loops terminated, stop on last frame or first frame. your choice
                    Exit Sub

                End If

            End If

        End If

        .Frame = .Frame + 1&

    End With

    Set FrmRanking.Picture = StdPictureEx.SubImage(FrmRanking.Picture, myAnimator.Frame)
    FrmRanking.Refresh  ' note: form/picturebox picture property does not require a refresh; updated automatically
    aniTimer.Interval = myAnimator.Intervals(myAnimator.Frame) * 20
    aniTimer.Enabled = True
err:
End Sub

Private Sub ImgClan_Click()
Call Audio.PlayWave(SND_CLICK)

Call WriteSolicitarRanking(TopClanes)
RankingOro = ""
End Sub


Private Sub ImgFrags_Click()

Call Audio.PlayWave(SND_CLICKNEW)
 Call WriteSolicitarRanking(TopFrags)
   FrmRanking2.Picture = LoadPictureEX("RankingAsesinados.jpg")
    RankingOro = ""
   
End Sub

Private Sub ImgNivel_Click()

Call Audio.PlayWave(SND_CLICKNEW)
Call WriteSolicitarRanking(TopLevel)
RankingOro = ""
End Sub

Private Sub ImgOro_Click()
Call Audio.PlayWave(SND_CLICKNEW)
FrmRanking2.Picture = LoadPictureEX("RankingOro_1.jpg")
Call WriteSolicitarRanking(TopOro)
RankingOro = "$"
End Sub

Private Sub ImgReto_Click()
Call Audio.PlayWave(SND_CLICKNEW)

Call WriteSolicitarRanking(TopRetos)
RankingOro = ""
End Sub

Private Sub ImgTorneo_Click()
Call Audio.PlayWave(SND_CLICKNEW)
 Call WriteSolicitarRanking(TopLevel)
   FrmRanking2.Picture = LoadPictureEX("RankingLevel.jpg")
    RankingOro = ""
End Sub

Private Sub Label1_Click()
Call Audio.PlayWave(SND_CLICKNEW)
Unload Me
frmMain.SetFocus
End Sub


Sub Animacion(imagen2 As Form)

    If aniTimer.Enabled Then
        aniTimer.Enabled = False
    ElseIf myAnimator.Frames = 0& Then

        Set imagen2.Picture = StdPictureEx.LoadPicture("C:\Users\waalter\Desktop\3434.gif", , , , , True)  ' True=can change frames
        myAnimator.Frames = StdPictureEx.SubImageCount(imagen2.Picture)

        If myAnimator.Frames < 2 Or StdPictureEx.PictureType(imagen2.Picture) <> ptcGIF Then
            myAnimator.Frames = -1    ' flag indicating this image is not GIF or can't be animated
            aniTimer.Interval = 0
        Else
            myAnimator.Frame = 1
            Call StdPictureEx.GetGIFAnimationInfo(imagen2.Picture, myAnimator.Intervals)
            aniTimer.Interval = myAnimator.Intervals(1) * 20
            aniTimer.Enabled = True

        End If

    Else
        aniTimer.Enabled = True

    End If

End Sub
