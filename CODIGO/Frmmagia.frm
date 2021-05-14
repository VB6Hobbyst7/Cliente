VERSION 5.00
Begin VB.Form Frmmagia 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   Picture         =   "Frmmagia.frx":0000
   ScaleHeight     =   3765
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   480
      ScaleHeight     =   168
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   0
      Top             =   720
      Width           =   2220
      Begin VB.Label lblitem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   4080
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Frmmagia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InvX                 As Integer

Dim InvY                 As Integer



Private Sub Form_Load()
Skin Frmmagia, vbRed
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Button = vbLeftButton) Then Call Auto_Drag(Frmmagia.hwnd)
End Sub
Private Sub PicInv_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    InvX = X
    InvY = Y

    If Button = 2 And Not Comerciando Then
        If Inventario.GrhIndex(Inventario.SelectedItem) > 0 Then
            DragAndDrop = True
            Me.MouseIcon = GetIcon(Inventario.Grafico(GrhData(Inventario.GrhIndex(Inventario.SelectedItem)).FileNum), 0, 0, Halftone, True, RGB(255, 0, 255))
            Me.MousePointer = 99

        End If

    End If

End Sub
Private Sub picInv_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
If (Button = vbLeftButton) Then Call Auto_Drag(picInv.hwnd)
    If InvX >= Inventario.OffSetX And InvY >= Inventario.OffSetY Then
        Call Audio.PlayWave(SND_CLICK)

    End If

End Sub
Private Sub picInv_DblClick()

    If InvX >= Inventario.OffSetX And InvY >= Inventario.OffSetY Then
        If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
        If Not MainTimer.Check(TimersIndex.PuedeUsarDobleClick) Then Exit Sub
    
        If frmMain.macrotrabajo.Enabled Then frmMain.DesactivarMacroTrabajo
    
        Select Case Inventario.ObjType(Inventario.SelectedItem)
        
            Case eObjType.otcasco
                Call frmMain.EquiparItem
    
            Case eObjType.otArmadura
                Call frmMain.EquiparItem

            Case eObjType.otescudo
                Call frmMain.EquiparItem
        
            Case eObjType.otWeapon
           
                If InStr(Inventario.ItemName(Inventario.SelectedItem), "Arco") > 0 Then
                    If Inventario.Equipped(Inventario.SelectedItem) Then
                        Call frmMain.UsarItem
                    Else
                        Call frmMain.EquiparItem

                    End If

                ElseIf InStr(Inventario.ItemName(Inventario.SelectedItem), "Bala") > 0 Then

                    If Inventario.Equipped(Inventario.SelectedItem) Then
                        Call frmMain.UsarItem
                        UsingSecondSkill = 1
                    Else
                        Call frmMain.EquiparItem

                    End If

                Else
                    Call frmMain.EquiparItem

                End If

            Case eObjType.otAnillo
                Call frmMain.EquiparItem
            
            Case eObjType.otFlechas
                Call frmMain.EquiparItem
        
            Case Else
                Call frmMain.UsarItem
            
        End Select
    
    End If

End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If DragAndDrop Then
        frmMain.MouseIcon = Nothing
        frmMain.MousePointer = 99
        Call SetCursor(General)

    End If

    If Button = 2 And DragAndDrop And Inventario.SelectedItem > 0 And Not Comerciando Then
        If X >= Inventario.OffSetX And Y >= Inventario.OffSetY And X <= picInv.Width And Y <= picInv.Height Then

            Dim NewPosInv As Integer

            NewPosInv = Inventario.ClickItem(X, Y)

            If NewPosInv > 0 Then
                Call WriteIntercambiarInv(Inventario.SelectedItem, NewPosInv, False)
                Call Inventario.Intercambiar(NewPosInv)

            End If
    
        Else

            Dim DropX As Integer, tmpX As Integer

            Dim DropY As Integer, tmpY As Integer

            tmpX = X + 823 - frmMain.pRender.Left
            tmpY = Y + 200 - frmMain.pRender.Top
        
            If tmpX > 0 And tmpX < frmMain.pRender.Width And tmpY > 0 And tmpY < frmMain.pRender.Height Then
                Call ConvertCPtoTP(tmpX, tmpY, DropX, DropY)
        
                'Solo tira a un tilde de distancia...
                If DropX < UserPos.X - 1 Then
                    DropX = UserPos.X - 1
                    DropY = UserPos.Y
                ElseIf DropX > UserPos.X + 1 Then
                    DropX = UserPos.X + 1
                    DropY = UserPos.Y
                ElseIf DropY < UserPos.Y - 1 Then
                    DropY = UserPos.Y - 1
                    DropX = UserPos.X
                ElseIf DropY > UserPos.Y + 1 Then
                    DropY = UserPos.Y + 1
                    DropX = UserPos.X

                End If
            
                If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                    Call WriteDrop(Inventario.SelectedItem, 1, DropX, DropY)
                Else

                    If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                        Inventario.DropX = DropX
                        Inventario.DropY = DropY
                        frmCantidad.Show , frmMain

                    End If

                End If

            End If

        End If

    End If

    DragAndDrop = False

End Sub


