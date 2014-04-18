Attribute VB_Name = "modText"
Option Explicit

' Text declares
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' Used to set a font for GDI text drawing
Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
    frmMain.Font = Font
    frmMain.FontSize = Size - 5
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetFont", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' GDI text drawing onto buffer
Public Sub DrawText(ByVal hdc As Long, ByVal x, ByVal y, ByVal Text As String, Color As Long)
Dim OldFont As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call SetFont(FONT_NAME, FONT_SIZE)
    OldFont = SelectObject(hdc, GameFont)
    Call SetBkMode(hdc, vbTransparent)
    Call SetTextColor(hdc, 0)
    Call TextOut(hdc, x + 1, y + 1, Text, Len(Text))
    Call SetTextColor(hdc, Color)
    Call TextOut(hdc, x, y, Text, Len(Text))
    Call SelectObject(hdc, OldFont)
    Call DeleteObject(GameFont)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPlayerName(ByVal index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim name As String
Dim Title As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check access level
    Select Case GetPlayerAccess(index)
        Case 0
            Color = RGB(255, 96, 0)
            Title = ""
        Case 1
            Color = QBColor(DarkGrey)
            Title = "(MOD) "
        Case 2
            Color = QBColor(Cyan)
            Title = "(MAP) "
        Case 3
            Color = QBColor(BrightGreen)
            Title = "(DEV) "
        Case 4
            Color = QBColor(Yellow)
            Title = "(GOD) "
    End Select

    name = Title + Trim$(Player(index).name)
    ' calc pos
    TextX = ConvertMapX(GetPlayerX(index) * PIC_X) + Player(index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(name)))
    If GetPlayerSprite(index) < 1 Or GetPlayerSprite(index) > NumCharacters Then
        TextY = ConvertMapY(GetPlayerY(index) * PIC_Y) + Player(index).YOffset - 16
    Else
        ' Determine location for text
        TextY = ConvertMapY(GetPlayerY(index) * PIC_Y) + Player(index).YOffset - (DDSD_Character(GetPlayerSprite(index)).lHeight / 4) + 16
    End If

    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, name, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpcName(ByVal index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim name As String
Dim NpcNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NpcNum = MapNpc(index).num

    Select Case Npc(NpcNum).Behaviour
        Case NPC_BEHAVIOUR_ATTACKONSIGHT
            Color = QBColor(BrightRed)
        Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
            Color = QBColor(Yellow)
        Case NPC_BEHAVIOUR_GUARD
            Color = QBColor(Grey)
        Case Else
            Color = QBColor(BrightGreen)
    End Select

    name = Trim$(Npc(NpcNum).name)
    TextX = ConvertMapX(MapNpc(index).x * PIC_X) + MapNpc(index).XOffset + (PIC_X \ 2) - getWidth(TexthDC, (Trim$(name)))
    If Npc(NpcNum).Sprite < 1 Or Npc(NpcNum).Sprite > NumCharacters Then
        TextY = ConvertMapY(MapNpc(index).y * PIC_Y) + MapNpc(index).YOffset - 16
    Else
        ' Determine location for text
        TextY = ConvertMapY(MapNpc(index).y * PIC_Y) + MapNpc(index).YOffset - (DDSD_Character(Npc(NpcNum).Sprite).lHeight / 4) + 16
    End If

    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, name, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpcName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function BltMapAttributes()
    Dim x As Long
    Dim y As Long
    Dim tX As Long
    Dim tY As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optAttribs.value Then
        For x = TileView.Left To TileView.Right
            For y = TileView.top To TileView.Bottom
                If IsValidMapPoint(x, y) Then
                    With Map.Tile(x, y)
                        tX = ((ConvertMapX(x * PIC_X)) - 4) + (PIC_X * 0.5)
                        tY = ((ConvertMapY(y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                DrawText TexthDC, tX, tY, "B", QBColor(BrightRed)
                            Case TILE_TYPE_WARP
                                DrawText TexthDC, tX, tY, "W", QBColor(BrightBlue)
                            Case TILE_TYPE_ITEM
                                DrawText TexthDC, tX, tY, "I", QBColor(White)
                            Case TILE_TYPE_NPCAVOID
                                DrawText TexthDC, tX, tY, "N", QBColor(White)
                            Case TILE_TYPE_KEY
                                DrawText TexthDC, tX, tY, "K", QBColor(White)
                            Case TILE_TYPE_KEYOPEN
                                DrawText TexthDC, tX, tY, "O", QBColor(White)
                            Case TILE_TYPE_RESOURCE
                                DrawText TexthDC, tX, tY, "O", QBColor(Green)
                            Case TILE_TYPE_NPCSPAWN
                                DrawText TexthDC, tX, tY, "S", QBColor(Yellow)
                            Case TILE_TYPE_SHOP
                                DrawText TexthDC, tX, tY, "S", QBColor(BrightBlue)
                            Case TILE_TYPE_BANK
                                DrawText TexthDC, tX, tY, "B", QBColor(Blue)
                            Case TILE_TYPE_HEAL
                                DrawText TexthDC, tX, tY, "H", QBColor(BrightGreen)
                            Case TILE_TYPE_TRAP
                                DrawText TexthDC, tX, tY, "T", QBColor(BrightRed)
                            Case TILE_TYPE_SLIDE
                                DrawText TexthDC, tX, tY, "S", QBColor(BrightCyan)
                        End Select
                    End With
                End If
            Next
        Next
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "BltMapAttributes", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub BltActionMsg(ByVal index As Long)
    Dim x As Long, y As Long, i As Long, time As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' does it exist
    If ActionMsg(index).Created = 0 Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(index).Type
        Case ACTIONMSG_STATIC
            time = 1500

            If ActionMsg(index).y > 0 Then
                x = ActionMsg(index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(index).message)) \ 2) * 8)
                y = ActionMsg(index).y - Int(PIC_Y \ 2) - 2
            Else
                x = ActionMsg(index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(index).message)) \ 2) * 8)
                y = ActionMsg(index).y - Int(PIC_Y \ 2) + 18
            End If

        Case ACTIONMSG_SCROLL
            time = 1500
        
            If ActionMsg(index).y > 0 Then
                x = ActionMsg(index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(index).message)) \ 2) * 8)
                y = ActionMsg(index).y - Int(PIC_Y \ 2) - 2 - (ActionMsg(index).Scroll * 0.6)
                ActionMsg(index).Scroll = ActionMsg(index).Scroll + 1
            Else
                x = ActionMsg(index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(index).message)) \ 2) * 8)
                y = ActionMsg(index).y - Int(PIC_Y \ 2) + 18 + (ActionMsg(index).Scroll * 0.6)
                ActionMsg(index).Scroll = ActionMsg(index).Scroll + 1
            End If

        Case ACTIONMSG_SCREEN
            time = 3000

            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1
                If ActionMsg(i).Type = ACTIONMSG_SCREEN Then
                    If i <> index Then
                        ClearActionMsg index
                        index = i
                    End If
                End If
            Next
            x = (frmMain.picScreen.width \ 2) - ((Len(Trim$(ActionMsg(index).message)) \ 2) * 8)
            y = 425

    End Select
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)

    If timeGetTime < ActionMsg(index).Created + time Then
        Call DrawText(TexthDC, x, y, ActionMsg(index).message, QBColor(ActionMsg(index).Color))
    Else
        ClearActionMsg index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltActionMsg", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function getWidth(ByVal DC As Long, ByVal Text As String) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    getWidth = frmMain.TextWidth(Text) \ 2
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "getWidth", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub AddText(ByVal Msg As String, ByVal Color As Integer)
Dim S As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    S = vbNewLine & Msg
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.Text)
    frmMain.txtChat.SelColor = QBColor(Color)
    frmMain.txtChat.SelText = S
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.Text) - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
