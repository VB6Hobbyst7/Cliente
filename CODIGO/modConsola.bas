Attribute VB_Name = "modConsola"
Option Explicit
 
Const CONSOLE_LINES As Integer = 10
 
Private Type consoleLine
        mString     As String
        lngColour   As Long
End Type
 
Private renderConsole(CONSOLE_LINES - 1) As consoleLine
 
Public Sub setConsoleText(ByRef mText As String, ByVal bRed As Integer, ByVal bGreen As Integer, ByVal bBlue As Integer)
 
    '
    ' @ maTih.-
    
    Dim loopC As Long
   
    Dim tmp(CONSOLE_LINES - 1) As consoleLine
    
    For loopC = 0 To (CONSOLE_LINES - 1)
        tmp(loopC) = renderConsole(loopC)
    Next loopC
    
    For loopC = 1 To (CONSOLE_LINES - 1)
        renderConsole(loopC) = tmp(loopC - 1)
    Next loopC
    
    With renderConsole(0)
         .lngColour = D3DColorRGBA(bRed, bGreen, bBlue, 255)
         .mString = mid$(mText, 1)
    End With
 
End Sub
 
Public Sub renderConsoleText()
 
    '
    ' @ maTih.-
    
    Dim renderX As Integer
    Dim renderY As Integer
    Dim loopC As Long
    
    renderX = 280
    
    For loopC = 0 To (CONSOLE_LINES - 4)
         renderY = 450 - ((loopC + 3) * 15)
         Call Mod_TileEngine.RenderText(renderX, renderY, renderConsole(loopC).mString, renderConsole(loopC).lngColour)
    Next loopC
    
    
    
    
    
    
    
    
    
    
    
 
End Sub
