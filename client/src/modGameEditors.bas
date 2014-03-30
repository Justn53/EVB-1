Attribute VB_Name = "modGameEditors"
Option Explicit

'**************
'* Map Editor *
'**************
Public Sub MapEditorInit()
Dim i As Long
Dim sMusic() As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' set the width
    frmEditor_Map.width = 7425
    
    ' we're in the map editor
    InMapEditor = True
    
    ' show the form
    frmEditor_Map.Visible = True
    
    ' set the scrolly bars
    frmEditor_Map.scrlTileSet.max = NumTileSets
    frmEditor_Map.fraTileSet.Caption = "Tileset: " & 1
    frmEditor_Map.scrlTileSet.value = 1
    
    ' render the tiles
    Call EditorMap_BltTileset
    
    ' set the scrollbars
    frmEditor_Map.scrlPictureY.max = (frmEditor_Map.picBackSelect.height \ PIC_Y) - (frmEditor_Map.picBack.height \ PIC_Y)
    frmEditor_Map.scrlPictureX.max = (frmEditor_Map.picBackSelect.width \ PIC_X) - (frmEditor_Map.picBack.width \ PIC_X)
    MapEditorTileScroll
    
    ' set shops for the shop attribute
    frmEditor_Map.cmbShop.AddItem "None"
    For i = 1 To MAX_SHOPS
        frmEditor_Map.cmbShop.AddItem i & ": " & Shop(i).name
    Next
    
    ' we're not in a shop
    frmEditor_Map.cmbShop.ListIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorProperties()
Dim X As Long
Dim Y As Long
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_MapProperties.lstMusic.Clear
    frmEditor_MapProperties.lstMusic.AddItem "None."
    For i = 1 To UBound(musicCache)
        frmEditor_MapProperties.lstMusic.AddItem musicCache(i)
    Next
    ' finished populating
    
    With frmEditor_MapProperties
        .txtName.Text = Trim$(Map.name)
        
        ' find the music we have set
        If .lstMusic.ListCount >= 0 Then
            .lstMusic.ListIndex = 0
            For i = 0 To .lstMusic.ListCount
                If .lstMusic.List(i) = Trim$(Map.Music) Then
                    .lstMusic.ListIndex = i
                End If
            Next
        End If
        
        ' rest of it
        .txtUp.Text = CStr(Map.Up)
        .txtDown.Text = CStr(Map.Down)
        .txtLeft.Text = CStr(Map.Left)
        .txtRight.Text = CStr(Map.Right)
        .cmbMoral.ListIndex = Map.Moral
        .txtBootMap.Text = CStr(Map.BootMap)
        .txtBootX.Text = CStr(Map.BootX)
        .txtBootY.Text = CStr(Map.BootY)

        ' show the map npcs
        .lstNpcs.Clear
        For X = 1 To MAX_MAP_NPCS
            If Map.Npc(X) > 0 Then
            .lstNpcs.AddItem X & ": " & Trim$(Npc(Map.Npc(X)).name)
            Else
                .lstNpcs.AddItem X & ": No NPC"
            End If
        Next
        .lstNpcs.ListIndex = 0
        
        ' show the npc selection combo
        .cmbNpc.Clear
        .cmbNpc.AddItem "No NPC"
        For X = 1 To MAX_NPCS
            .cmbNpc.AddItem X & ": " & Trim$(Npc(X).name)
        Next
        
        ' set the combo box properly
        Dim tmpString() As String
        Dim npcNum As Long
        tmpString = Split(.lstNpcs.List(.lstNpcs.ListIndex))
        npcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        .cmbNpc.ListIndex = Map.Npc(npcNum)
    
        ' show the current map
        .lblMap.Caption = "Current map: " & GetPlayerMap(MyIndex)
        .txtMaxX.Text = Map.MaxX
        .txtMaxY.Text = Map.MaxY
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorProperties", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSetTile(ByVal X As Long, ByVal Y As Long, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False)
Dim x2 As Long, y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not multitile Then ' single
        With Map.Tile(X, Y)
            ' set layer
            .Layer(CurLayer).X = EditorTileX
            .Layer(CurLayer).Y = EditorTileY
            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.value
        End With
    Else ' multitile
        y2 = 0 ' starting tile for y axis
        For Y = CurY To CurY + EditorTileHeight - 1
            x2 = 0 ' re-set x count every y loop
            For X = CurX To CurX + EditorTileWidth - 1
                If X >= 0 And X <= Map.MaxX Then
                    If Y >= 0 And Y <= Map.MaxY Then
                        With Map.Tile(X, Y)
                            .Layer(CurLayer).X = EditorTileX + x2
                            .Layer(CurLayer).Y = EditorTileY + y2
                            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.value
                        End With
                    End If
                End If
                x2 = x2 + 1
            Next
            y2 = y2 + 1
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorSetTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal movedMouse As Boolean = True)
Dim i As Long
Dim CurLayer As Long
Dim tmpDir As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).value Then
            CurLayer = i
            Exit For
        End If
    Next

    If Not isInBounds Then Exit Sub
    If Button = vbLeftButton Then
        If frmEditor_Map.optLayers.value Then
            If (EditorTileWidth = 1 And EditorTileHeight = 1) Or CurLayer = MapLayer.Autotile Then 'single tile
                MapEditorSetTile CurX, CurY, CurLayer
            Else ' multi tile!
                MapEditorSetTile CurX, CurY, CurLayer, True
            End If
        ElseIf frmEditor_Map.optAttribs.value Then
            With Map.Tile(CurX, CurY)
                ' blocked tile
                If frmEditor_Map.optBlocked.value Then .Type = TILE_TYPE_BLOCKED
                ' warp tile
                If frmEditor_Map.optWarp.value Then
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                End If
                ' item spawn
                If frmEditor_Map.optItem.value Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                End If
                ' npc avoid
                If frmEditor_Map.optNpcAvoid.value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' key
                If frmEditor_Map.optKey.value Then
                    .Type = TILE_TYPE_KEY
                    .Data1 = KeyEditorNum
                    .Data2 = KeyEditorTake
                    .Data3 = 0
                End If
                ' key open
                If frmEditor_Map.optKeyOpen.value Then
                    .Type = TILE_TYPE_KEYOPEN
                    .Data1 = KeyOpenEditorX
                    .Data2 = KeyOpenEditorY
                    .Data3 = 0
                End If
                ' resource
                If frmEditor_Map.optResource.value Then
                    .Type = TILE_TYPE_RESOURCE
                    .Data1 = ResourceEditorNum
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' npc spawn
                If frmEditor_Map.optNpcSpawn.value Then
                    .Type = TILE_TYPE_NPCSPAWN
                    .Data1 = SpawnNpcNum
                    .Data2 = SpawnNpcDir
                    .Data3 = 0
                End If
                ' shop
                If frmEditor_Map.optShop.value Then
                    .Type = TILE_TYPE_SHOP
                    .Data1 = EditorShop
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' bank
                If frmEditor_Map.optBank.value Then
                    .Type = TILE_TYPE_BANK
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' heal
                If frmEditor_Map.optHeal.value Then
                    .Type = TILE_TYPE_HEAL
                    .Data1 = MapEditorHealType
                    .Data2 = MapEditorHealAmount
                    .Data3 = 0
                End If
                ' trap
                If frmEditor_Map.optTrap.value Then
                    .Type = TILE_TYPE_TRAP
                    .Data1 = MapEditorHealAmount
                    .Data2 = 0
                    .Data3 = 0
                End If
                ' slide
                If frmEditor_Map.optSlide.value Then
                    .Type = TILE_TYPE_SLIDE
                    .Data1 = MapEditorSlideDir
                    .Data2 = 0
                    .Data3 = 0
                End If
            End With
        ElseIf frmEditor_Map.optBlock.value Then
            If movedMouse Then Exit Sub
            ' find what tile it is
            X = X - ((X \ 32) * 32)
            Y = Y - ((Y \ 32) * 32)
            ' see if it hits an arrow
            For i = 1 To 4
                If X >= DirArrowX(i) And X <= DirArrowX(i) + 8 Then
                    If Y >= DirArrowY(i) And Y <= DirArrowY(i) + 8 Then
                        ' flip the value.
                        setDirBlock Map.Tile(CurX, CurY).DirBlock, CByte(i), Not isDirBlocked(Map.Tile(CurX, CurY).DirBlock, CByte(i))
                        Exit Sub
                    End If
                End If
            Next
        End If
    End If

    If Button = vbRightButton Then
        If frmEditor_Map.optLayers.value Then
            With Map.Tile(CurX, CurY)
                ' clear layer
                .Layer(CurLayer).X = 0
                .Layer(CurLayer).Y = 0
                .Layer(CurLayer).Tileset = 0
            End With
        ElseIf frmEditor_Map.optAttribs.value Then
            With Map.Tile(CurX, CurY)
                ' clear attribute
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
            End With

        End If
    End If

    CacheResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorMouseDown", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        
        EditorTileX = X \ PIC_X
        EditorTileY = Y \ PIC_Y
        
        frmEditor_Map.shpSelected.top = EditorTileY * PIC_Y
        frmEditor_Map.shpSelected.Left = EditorTileX * PIC_X
        
        frmEditor_Map.shpSelected.width = PIC_X
        frmEditor_Map.shpSelected.height = PIC_Y
        
        ' change it if it's autotiles
        If frmEditor_Map.optLayer(MapLayer.Autotile) Then
            EditorTileX = (X \ 32)
            EditorTileY = (Y \ 32)
            EditorTileWidth = 3
            EditorTileHeight = 4
            frmEditor_Map.shpSelected.top = EditorTileY * PIC_Y
            frmEditor_Map.shpSelected.Left = EditorTileX * PIC_X
            frmEditor_Map.shpSelected.width = 96
            frmEditor_Map.shpSelected.height = 128
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorChooseTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorDrag(Button As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Button = vbLeftButton Then
        ' exit out if we're autotiling
        If frmEditor_Map.optLayer(MapLayer.Autotile) Then Exit Sub
        ' convert the pixel number to tile number
        X = (X \ PIC_X) + 1
        Y = (Y \ PIC_Y) + 1
        ' check it's not out of bounds
        If X < 0 Then X = 0
        If X > frmEditor_Map.picBackSelect.width / PIC_X Then X = frmEditor_Map.picBackSelect.width / PIC_X
        If Y < 0 Then Y = 0
        If Y > frmEditor_Map.picBackSelect.height / PIC_Y Then Y = frmEditor_Map.picBackSelect.height / PIC_Y
        ' find out what to set the width + height of map editor to
        If X > EditorTileX Then ' drag right
            EditorTileWidth = X - EditorTileX
        Else ' drag left
            ' TO DO
        End If
        If Y > EditorTileY Then ' drag down
            EditorTileHeight = Y - EditorTileY
        Else ' drag up
            ' TO DO
        End If
        frmEditor_Map.shpSelected.width = EditorTileWidth * PIC_X
        frmEditor_Map.shpSelected.height = EditorTileHeight * PIC_Y
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorDrag", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorTileScroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' horizontal scrolling
    If frmEditor_Map.picBackSelect.width < frmEditor_Map.picBack.width Then
        frmEditor_Map.scrlPictureX.Enabled = False
    Else
        frmEditor_Map.scrlPictureX.Enabled = True
        frmEditor_Map.picBackSelect.Left = (frmEditor_Map.scrlPictureX.value * PIC_X) * -1
    End If
    
    ' vertical scrolling
    If frmEditor_Map.picBackSelect.height < frmEditor_Map.picBack.height Then
        frmEditor_Map.scrlPictureY.Enabled = False
    Else
        frmEditor_Map.scrlPictureY.Enabled = True
        frmEditor_Map.picBackSelect.top = (frmEditor_Map.scrlPictureY.value * PIC_Y) * -1
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorTileScroll", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSend()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SendMap
    InMapEditor = False
    Unload frmEditor_Map
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorSend", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorCancel()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set buffer = New clsBuffer
    buffer.WriteLong CNeedMap
    buffer.WriteLong 1
    SendData buffer.ToArray()
    InMapEditor = False
    Unload frmEditor_Map
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearLayer()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim CurLayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).value Then
            CurLayer = i
            Exit For
        End If
    Next
    
    If CurLayer = 0 Then Exit Sub

    ' ask to clear layer
    If MsgBox("Are you sure you wish to clear this layer?", vbYesNo, Options.Game_Name) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).Layer(CurLayer).X = 0
                Map.Tile(X, Y).Layer(CurLayer).Y = 0
                Map.Tile(X, Y).Layer(CurLayer).Tileset = 0
            Next
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorClearLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorFillLayer()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim CurLayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).value Then
            CurLayer = i
            Exit For
        End If
    Next

    ' Ground layer
    If MsgBox("Are you sure you wish to fill this layer?", vbYesNo, Options.Game_Name) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).Layer(CurLayer).X = EditorTileX
                Map.Tile(X, Y).Layer(CurLayer).Y = EditorTileY
                Map.Tile(X, Y).Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.value
            Next
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorFillLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearAttribs()
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, Options.Game_Name) = vbYes Then

        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).Type = 0
            Next
        Next

    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorClearAttribs", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorLeaveMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InMapEditor Then
        If MsgBox("Save changes to current map?", vbYesNo) = vbYes Then
            Call MapEditorSend
        Else
            Call MapEditorCancel
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapEditorLeaveMap", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'***************
'* Item Editor *
'***************
Public Sub ItemEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Item.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Item.cmbSound.Clear
    frmEditor_Item.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Item.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Item(EditorIndex)
        frmEditor_Item.txtName.Text = Trim$(.name)
        If .Pic > frmEditor_Item.scrlPic.max Then .Pic = 0
        frmEditor_Item.scrlPic.value = .Pic
        frmEditor_Item.cmbType.ListIndex = .Type
        frmEditor_Item.scrlAnim.value = .Animation
        frmEditor_Item.txtDesc.Text = Trim$(.Desc)
        
        ' find the sound we have set
        If frmEditor_Item.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Item.cmbSound.ListCount
                If frmEditor_Item.cmbSound.List(i) = Trim$(.Sound) Then
                    frmEditor_Item.cmbSound.ListIndex = i
                    SoundSet = True
                    Exit For
                End If
            Next
            If Not SoundSet Or frmEditor_Item.cmbSound.ListIndex = -1 Then frmEditor_Item.cmbSound.ListIndex = 0
        End If

        ' Type specific settings
        If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            frmEditor_Item.fraEquipment.Visible = True
            frmEditor_Item.scrlDamage.value = .Data2
            frmEditor_Item.cmbTool.ListIndex = .Data3

            If .Speed < 100 Then .Speed = 100
            frmEditor_Item.scrlSpeed.value = .Speed
            
            ' loop for stats
            For i = 1 To Stats.Stat_Count - 1
                frmEditor_Item.scrlStatBonus(i).value = .Add_Stat(i)
            Next
            
            frmEditor_Item.scrlPaperdoll = .Paperdoll
        Else
            frmEditor_Item.fraEquipment.Visible = False
            ' clear the stats
            For i = 1 To Stats.Stat_Count - 1
                frmEditor_Item.scrlStatBonus(i).value = 0
            Next
        End If

        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CONSUME Then
            frmEditor_Item.fraVitals.Visible = True
            frmEditor_Item.scrlAddHp.value = .AddHP
            frmEditor_Item.scrlAddMP.value = .AddMP
            frmEditor_Item.scrlAddExp.value = .AddEXP
            frmEditor_Item.scrlCastSpell.value = .CastSpell
        Else
            frmEditor_Item.fraVitals.Visible = False
        End If

        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditor_Item.fraSpell.Visible = True
            frmEditor_Item.scrlSpell.value = .Data1
        Else
            frmEditor_Item.fraSpell.Visible = False
        End If

        ' Basic requirements
        frmEditor_Item.scrlAccessReq.value = .AccessReq
        frmEditor_Item.scrlLevelReq.value = .LevelReq
        
        ' loop for stats
        For i = 1 To Stats.Stat_Count - 1
            frmEditor_Item.scrlStatReq(i).value = .Stat_Req(i)
        Next
        
        ' Build cmbClassReq
        frmEditor_Item.cmbClassReq.Clear
        frmEditor_Item.cmbClassReq.AddItem "None"

        For i = 1 To Max_Classes
            frmEditor_Item.cmbClassReq.AddItem Class(i).name
        Next

        frmEditor_Item.cmbClassReq.ListIndex = .ClassReq
        ' Info
        frmEditor_Item.scrlPrice.value = .Price
        If .Tradable Then
            frmEditor_Item.chkTradable.value = 1
        Else
            frmEditor_Item.chkTradable.value = 0
        End If
        frmEditor_Item.scrlRarity.value = .Rarity
        EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    End With

    Call EditorItem_BltItem
    Call EditorItem_BltPaperdoll
    Item_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        If Item_Changed(i) Then
            Call SendSaveItem(i)
        End If
    Next
    
    Unload frmEditor_Item
    Editor = 0
    ClearChanged_Item
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Item
    ClearChanged_Item
    ClearItems
    SendRequestItems
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ItemEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Item()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Item_Changed(1), MAX_ITEMS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Item", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'********************
'* Animation Editor *
'********************
Public Sub AnimationEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Animation.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Animation.cmbSound.Clear
    frmEditor_Animation.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Animation.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Animation(EditorIndex)
        frmEditor_Animation.txtName.Text = Trim$(.name)
        
        ' find the sound we have set
        If frmEditor_Animation.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Animation.cmbSound.ListCount
                If frmEditor_Animation.cmbSound.List(i) = Trim$(.Sound) Then
                    frmEditor_Animation.cmbSound.ListIndex = i
                    SoundSet = True
                    Exit For
                End If
            Next
            If Not SoundSet Or frmEditor_Animation.cmbSound.ListIndex = -1 Then frmEditor_Animation.cmbSound.ListIndex = 0
        End If
        
        For i = 0 To 1
            frmEditor_Animation.scrlSprite(i).value = .Sprite(i)
            frmEditor_Animation.scrlFrameCount(i).value = .Frames(i)
            frmEditor_Animation.scrlLoopCount(i).value = .LoopCount(i)
            
            If .looptime(i) > 0 Then
                frmEditor_Animation.scrlLoopTime(i).value = .looptime(i)
            Else
                frmEditor_Animation.scrlLoopTime(i).value = 45
            End If
            
        Next
         
        EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    End With

    Call EditorAnim_BltAnim
    Animation_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        If Animation_Changed(i) Then
            Call SendSaveAnimation(i)
        End If
    Next
    
    Unload frmEditor_Animation
    Editor = 0
    ClearChanged_Animation
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Animation
    ClearChanged_Animation
    ClearAnimations
    SendRequestAnimations
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AnimationEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Animation()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Animation_Changed(1), MAX_ANIMATIONS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Animation", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'**************
'* NPC Editor *
'**************
Public Sub NpcEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_NPC.Visible = False Then Exit Sub
    EditorIndex = frmEditor_NPC.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_NPC.cmbSound.Clear
    frmEditor_NPC.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_NPC.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_NPC
        .txtName.Text = Trim$(Npc(EditorIndex).name)
        .txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
        If Npc(EditorIndex).Sprite < 0 Or Npc(EditorIndex).Sprite > .scrlSprite.max Then Npc(EditorIndex).Sprite = 0
        .scrlSprite.value = Npc(EditorIndex).Sprite
        .txtSpawnSecs.Text = CStr(Npc(EditorIndex).SpawnSecs)
        .cmbBehaviour.ListIndex = Npc(EditorIndex).Behaviour
        .scrlRange.value = Npc(EditorIndex).Range
        
        .txtChance.Text = CStr(Npc(EditorIndex).DropChance(1))
        .scrlNum.value = Npc(EditorIndex).DropItem(1)
        .scrlValue.value = Npc(EditorIndex).DropItemValue(1)
        .scrlDropIndex.max = MAX_NPC_DROPS
        
        .txtHP.Text = Npc(EditorIndex).HP
        .txtEXP.Text = Npc(EditorIndex).EXP
        .txtLevel.Text = Npc(EditorIndex).Level
        .txtDamage.Text = Npc(EditorIndex).Damage
        .scrlAnimation.value = Npc(EditorIndex).Animation
        .scrlConv.value = Npc(EditorIndex).Conv
        .scrlFace.max = NumFaces
        .scrlFace.value = Npc(EditorIndex).Face

        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Npc(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                    Exit For
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
        
        For i = 1 To Stats.Stat_Count - 1
            .scrlStat(i).value = Npc(EditorIndex).Stat(i)
        Next
        
        ' quest data
        Select Case Npc(EditorIndex).questGiver
            Case 0: .chkQuestGiver = vbUnchecked
            Case 1: .chkQuestGiver = vbChecked
        End Select
        
        .scrlQuestIndex.max = MAX_QUESTS
        .scrlQuestIndex = Npc(EditorIndex).questIndex
        
        .scrlQuestStartChat.max = MAX_CONV_CHATS
        .scrlQuestStartChat = Npc(EditorIndex).questStartChat
        
        .scrlQuestInProgressChat.max = MAX_CONV_CHATS
        .scrlQuestInProgressChat = Npc(EditorIndex).questInProgressChat
        
        .scrlPostQuestChat.max = MAX_CONV_CHATS
        .scrlPostQuestChat = Npc(EditorIndex).postQuestChat
    End With
    
    Call EditorNpc_BltSprite
    NPC_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        If NPC_Changed(i) Then
            Call SendSaveNpc(i)
        End If
    Next
    
    Unload frmEditor_NPC
    Editor = 0
    ClearChanged_NPC
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_NPC
    ClearChanged_NPC
    ClearNPCs
    SendRequestNPCS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_NPC()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory NPC_Changed(1), MAX_NPCS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_NPC", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' *******************
' * Resource Editor *
' *******************
Public Sub ResourceEditorInit()
Dim i As Long
Dim SoundSet As Boolean

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Resource.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Resource.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Resource.cmbSound.Clear
    frmEditor_Resource.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Resource.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_Resource
        .scrlExhaustedPic.max = NumResources
        .scrlNormalPic.max = NumResources
        .scrlAnimation.max = MAX_ANIMATIONS
        
        .txtName.Text = Trim$(Resource(EditorIndex).name)
        .txtMessage.Text = Trim$(Resource(EditorIndex).SuccessMessage)
        .txtMessage2.Text = Trim$(Resource(EditorIndex).EmptyMessage)
        .cmbType.ListIndex = Resource(EditorIndex).ResourceType
        .scrlNormalPic.value = Resource(EditorIndex).ResourceImage
        .scrlExhaustedPic.value = Resource(EditorIndex).ExhaustedImage
        .scrlReward.value = Resource(EditorIndex).ItemReward
        .scrlTool.value = Resource(EditorIndex).ToolRequired
        .scrlHealth.value = Resource(EditorIndex).health
        .scrlRespawn.value = Resource(EditorIndex).RespawnTime
        .scrlAnimation.value = Resource(EditorIndex).Animation
        
        .scrlQuestIndex.value = Resource(EditorIndex).questIndex
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Resource(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                    Exit For
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
        
    Call EditorResource_BltSprite
    
    Resource_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        If Resource_Changed(i) Then
            Call SendSaveResource(i)
        End If
    Next
    
    Unload frmEditor_Resource
    Editor = 0
    ClearChanged_Resource
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Resource
    ClearChanged_Resource
    ClearResources
    SendRequestResources
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResourceEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Resource()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Resource_Changed(1), MAX_RESOURCES * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Resource", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'***************
'* Shop Editor *
'***************
Public Sub ShopEditorInit()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If frmEditor_Shop.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Shop.lstIndex.ListIndex + 1
    
    With frmEditor_Shop
        .txtName.Text = Trim$(Shop(EditorIndex).name)
        If Shop(EditorIndex).BuyRate > 0 Then
            .scrlBuy.value = Shop(EditorIndex).BuyRate
        Else
            .scrlBuy.value = 100
        End If
    
        .cmbItem.Clear
        .cmbItem.AddItem "None"
        .cmbCostItem.Clear
        .cmbCostItem.AddItem "None"
        
        For i = 1 To MAX_ITEMS
            .cmbItem.AddItem i & ": " & Trim$(Item(i).name)
            .cmbCostItem.AddItem i & ": " & Trim$(Item(i).name)
        Next
        
        .cmbItem.ListIndex = 0
        .cmbCostItem.ListIndex = 0
        
        .cmbCostItem.ListIndex = Shop(EditorIndex).TradeItem(EditorIndex).CostItem
        .txtCostValue.Text = Shop(EditorIndex).TradeItem(EditorIndex).CostValue
        .cmbItem.ListIndex = Shop(EditorIndex).TradeItem(EditorIndex).Item
        .txtItemValue.Text = Shop(EditorIndex).TradeItem(EditorIndex).ItemValue
    End With
    
    UpdateShopTrade
    
    Shop_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateShopTrade(Optional ByVal tmpPos As Long = 0)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Shop.lstTradeItem.Clear

    For i = 1 To MAX_TRADES
        With Shop(EditorIndex).TradeItem(i)
            ' if none, show as none
            If .Item = 0 And .CostItem = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem "Empty Trade Slot"
            Else
                frmEditor_Shop.lstTradeItem.AddItem i & ": " & .ItemValue & "x " & Trim$(Item(.Item).name) & " for " & .CostValue & "x " & Trim$(Item(.CostItem).name)
            End If
        End With
    Next

    frmEditor_Shop.lstTradeItem.ListIndex = tmpPos
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateShopTrade", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        If Shop_Changed(i) Then
            Call SendSaveShop(i)
        End If
    Next
    
    Unload frmEditor_Shop
    Editor = 0
    ClearChanged_Shop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Shop
    ClearChanged_Shop
    ClearShops
    SendRequestShops
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ShopEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Shop()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Shop_Changed(1), MAX_SHOPS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Shop", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'****************
'* Spell Editor *
'****************
Public Sub SpellEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Spell.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Spell.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Spell.cmbSound.Clear
    frmEditor_Spell.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Spell.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_Spell
        ' set max values
        .scrlAnimCast.max = MAX_ANIMATIONS
        .scrlAnim.max = MAX_ANIMATIONS
        .scrlAOE.max = MAX_BYTE
        .scrlRange.max = MAX_BYTE
        .scrlMap.max = MAX_MAPS
        
        ' build class combo
        .cmbClass.Clear
        .cmbClass.AddItem "None"
        For i = 1 To Max_Classes
            .cmbClass.AddItem Trim$(Class(i).name)
        Next
        
        ' set values
        .txtName.Text = Trim$(Spell(EditorIndex).name)
        .txtDesc.Text = Trim$(Spell(EditorIndex).Desc)
        .cmbType.ListIndex = Spell(EditorIndex).Type
        .scrlMP.value = Spell(EditorIndex).MPCost
        .scrlLevel.value = Spell(EditorIndex).LevelReq
        .scrlAccess.value = Spell(EditorIndex).AccessReq
        .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        .scrlCast.value = Spell(EditorIndex).CastTime
        .scrlCool.value = Spell(EditorIndex).CDTime
        .scrlIcon.value = Spell(EditorIndex).Icon
        .scrlMap.value = Spell(EditorIndex).Map
        .scrlX.value = Spell(EditorIndex).X
        .scrlY.value = Spell(EditorIndex).Y
        .scrlDir.value = Spell(EditorIndex).Dir
        .scrlVital.value = Spell(EditorIndex).Vital
        .scrlDuration.value = Spell(EditorIndex).Duration
        .scrlInterval.value = Spell(EditorIndex).Interval
        .scrlRange.value = Spell(EditorIndex).Range
        If Spell(EditorIndex).IsAoE Then
            .chkAOE.value = 1
        Else
            .chkAOE.value = 0
        End If
        .scrlAOE.value = Spell(EditorIndex).AoE
        .scrlAnimCast.value = Spell(EditorIndex).CastAnim
        .scrlAnim.value = Spell(EditorIndex).SpellAnim
        .scrlStun.value = Spell(EditorIndex).StunDuration
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Spell(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                    Exit For
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
    
    EditorSpell_BltIcon
    
    Spell_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorOk()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        If Spell_Changed(i) Then
            Call SendSaveSpell(i)
        End If
    Next
    
    Unload frmEditor_Spell
    Editor = 0
    ClearChanged_Spell
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Editor = 0
    Unload frmEditor_Spell
    ClearChanged_Spell
    ClearSpells
    SendRequestSpells
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Spell()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Spell_Changed(1), MAX_SPELLS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Spell", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearAttributeDialogue()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Map.fraNpcSpawn.Visible = False
    frmEditor_Map.fraResource.Visible = False
    frmEditor_Map.fraMapItem.Visible = False
    frmEditor_Map.fraMapKey.Visible = False
    frmEditor_Map.fraKeyOpen.Visible = False
    frmEditor_Map.fraMapWarp.Visible = False
    frmEditor_Map.fraShop.Visible = False

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAttributeDialogue", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

'***********************
'* Conversation Editor *
'***********************
Public Sub ConvEditorInit()
Dim i As Long, j As Long
Dim CurChat As Byte, SoundSet As Boolean

    If frmEditor_Conv.Visible = False Then Exit Sub
    If ConvEditorLoaded = True Then Exit Sub
    
    With frmEditor_Conv
        ConvEditorLoaded = False
        
        ' Set the index
        EditorIndex = frmEditor_Conv.lstIndex.ListIndex + 1
        
        ' Set the name
        .txtName = Conv(EditorIndex).name
        
        ' Reset the bars
        If Conv(EditorIndex).ChatCount <= 0 Then Conv(EditorIndex).ChatCount = 1
        .scrlChatCount.value = Conv(EditorIndex).ChatCount
        .scrlCurChat.max = Conv(EditorIndex).ChatCount
        If .scrlCurChat.value = 0 Then .scrlCurChat.value = 1
        
        ' Lazyness
        CurChat = .scrlCurChat.value
        
        ' Populate the cache if we need to
        If Not hasPopulated Then
            PopulateLists
        End If
        
        ' Add the array to the combo
        .cmbSound.Clear
        .cmbSound.AddItem "None."
        For i = 1 To UBound(soundCache)
            .cmbSound.AddItem soundCache(i)
        Next
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Conv(EditorIndex).Chat(CurChat).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                    Exit For
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If

        ' Maximum amount of chats
        .scrlChatCount.max = MAX_CONV_CHATS
        .fraConv.Caption = "Conversation: (" & .scrlCurChat.value & "/" & .scrlChatCount.value & ")"
        
        ' Populate the reply comboboxes
        For i = 1 To 4
            .cmbToConv(i).Clear
        Next
        
        ' Add all the possible conv-to's
        For i = 1 To 4
            .cmbToConv(i).AddItem "None", 0
            For j = 1 To Conv(EditorIndex).ChatCount
                .cmbToConv(i).AddItem CStr(j), j
            Next
            .cmbToConv(i).ListIndex = Conv(EditorIndex).Chat(CurChat).ReplyConvTo(i)
        Next
        
        ' Set the replies
        For i = 1 To 4
            .txtReply(i).Text = Trim$(Conv(EditorIndex).Chat(CurChat).ReplyText(i))
        Next
        
        ' Populate the event combobox
        .cmbEvent.Clear
        .cmbEvent.AddItem "None", 0
        .cmbEvent.AddItem "Open Bank", 1
        .cmbEvent.AddItem "Open Shop", 2
        .cmbEvent.AddItem "Give Item", 3
        .cmbEvent.AddItem "Take Item", 4
        .cmbEvent.AddItem "Warp", 5
        .cmbEvent.AddItem "Heal", 6
        .cmbEvent.AddItem "Start Quest", 7
        .cmbEvent.AddItem "Finish Quest", 8
        
        ' Set it to the event
        .cmbEvent.ListIndex = Conv(EditorIndex).Chat(CurChat).Event
        
        ' Set the event data
        .scrlData1 = Conv(EditorIndex).Chat(CurChat).Data1
        .scrlData2 = Conv(EditorIndex).Chat(CurChat).Data2
        .scrlData3 = Conv(EditorIndex).Chat(CurChat).Data3
        InitEventData
        
        ' Set the text
        .txtConvText = Trim$(Conv(EditorIndex).Chat(CurChat).Text)
        
        ' Loaded
        ConvEditorLoaded = True
    End With
    
    ' Allows us to clear it if we make a mistake
    Conv_Changed(EditorIndex) = True
End Sub

Public Sub ConvEditorOK()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_CONVS
        If Conv_Changed(i) Then
            Call SendSaveConv(i)
        End If
    Next
    
    Unload frmEditor_Conv
    Editor = 0
    ClearChanged_Conv
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ConvEditorOK", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ConvEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Unload frmEditor_Conv
    Editor = 0
    ClearChanged_Conv
    
    ClearConvs
    SendRequestConvs
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ConvEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Conv()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Conv_Changed(1), MAX_CONVS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Conv", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub InitEventData()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Change the data labels to match the event
    With frmEditor_Conv
        Select Case .cmbEvent.ListIndex
            Case 0 ' None
                .lblData1.Visible = False
                .lblData2.Visible = False
                .lblData3.Visible = False
                
                .scrlData1.Visible = False
                .scrlData2.Visible = False
                .scrlData3.Visible = False
            Case 1 ' Open Bank
                .lblData1.Visible = False
                .lblData2.Visible = False
                .lblData3.Visible = False
                
                .scrlData1.Visible = False
                .scrlData2.Visible = False
                .scrlData3.Visible = False
            Case 2 ' Open Shop
                .lblData1.Visible = True
                .lblData2.Visible = False
                .lblData3.Visible = False
                
                .scrlData1.Visible = True
                .scrlData2.Visible = False
                .scrlData3.Visible = False
                
                ' Allow them to set the index of the shop
                If .scrlData1.value > MAX_SHOPS Then .scrlData1.value = MAX_SHOPS
                .lblData1.Caption = "Shop Index: " & .scrlData1.value
                .scrlData1.max = MAX_SHOPS
                .scrlData1.min = 1
            Case 3 ' Give Item
                .lblData1.Visible = True
                .lblData2.Visible = True
                .lblData3.Visible = False
                
                .scrlData1.Visible = True
                .scrlData2.Visible = True
                .scrlData3.Visible = False
                
                ' Allow them to set the index of the item
                If .scrlData1.value > MAX_ITEMS Then .scrlData1.value = MAX_ITEMS
                .lblData1.Caption = "Item Index: " & .scrlData1.value
                .scrlData1.max = MAX_ITEMS
                .scrlData1.min = 1
                
                ' Allow them to set the amount of the item
                If .scrlData2.value > 32767 Then .scrlData2.value = 32767
                .lblData2.Caption = "Value: " & .scrlData2.value
                .scrlData2.max = 32767
                .scrlData2.min = 1
            Case 4 ' Take Item
                .lblData1.Visible = True
                .lblData2.Visible = True
                .lblData3.Visible = False
                
                .scrlData1.Visible = True
                .scrlData2.Visible = True
                .scrlData3.Visible = False
                
                ' Allow them to set the index of the item
                If .scrlData1.value > MAX_ITEMS Then .scrlData1.value = MAX_ITEMS
                .lblData1.Caption = "Item Index: " & .scrlData1.value
                .scrlData1.max = MAX_ITEMS
                .scrlData1.min = 1
                
                ' Allow them to set the amount of the item
                If .scrlData2.value > 32767 Then .scrlData2.value = 32767
                .lblData2.Caption = "Value: " & .scrlData2.value
                .scrlData2.max = 32767
                .scrlData2.min = 1
            Case 5 ' Warp
                .lblData1.Visible = True
                .lblData2.Visible = True
                .lblData3.Visible = True
                
                .scrlData1.Visible = True
                .scrlData2.Visible = True
                .scrlData3.Visible = True
                
                ' Allow them to set the index of the map
                If .scrlData1.value > MAX_MAPS Then .scrlData1.value = MAX_MAPS
                .lblData1.Caption = "Map Index: " & .scrlData1.value
                .scrlData1.max = MAX_MAPS
                .scrlData1.min = 1
                
                ' Allow them to set the x
                If .scrlData2.value > 255 Then .scrlData2.value = 255
                .lblData2.Caption = "X: " & .scrlData2.value
                .scrlData2.max = 255
                .scrlData2.min = 1
                
                ' Allow them to set the y
                If .scrlData3.value > 255 Then .scrlData3.value = 255
                .lblData3.Caption = "Y: " & .scrlData3.value
                .scrlData3.max = 255
                .scrlData3.min = 1
            Case 6 ' Heal
                .lblData1.Visible = True
                .lblData2.Visible = False
                .lblData3.Visible = False
                
                .scrlData1.Visible = True
                .scrlData2.Visible = False
                .scrlData3.Visible = False
                
                ' Allow them to set the amount to heal
                If .scrlData1.value > 32767 Then .scrlData1.value = 32767
                .lblData1.Caption = "Heal amount: " & .scrlData1.value
                .scrlData1.max = 32767
                .scrlData1.min = 1
            Case 7 ' Start quest
                .lblData1.Visible = True
                .lblData2.Visible = False
                .lblData3.Visible = False
                
                .scrlData1.Visible = True
                .scrlData2.Visible = False
                .scrlData3.Visible = False
                If .scrlData1.value > MAX_QUESTS Then .scrlData1.value = MAX_QUESTS
                
                .lblData1.Caption = "Quest Index: " & .scrlData1.value
                .scrlData1.max = MAX_QUESTS
            Case 8 ' Finish quest
                .lblData1.Visible = True
                .lblData2.Visible = False
                .lblData3.Visible = False
                
                .scrlData1.Visible = True
                .scrlData2.Visible = False
                .scrlData3.Visible = False
                If .scrlData1.value > MAX_QUESTS Then .scrlData1.value = MAX_QUESTS

                .lblData1.Caption = "Quest Index: " & .scrlData1.value
                .scrlData1.max = MAX_QUESTS
        End Select
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitConvData", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub QuestEditorInit()
Dim i As Long

    With frmEditor_Quest
        ' Not loaded yet
        QuestEditorLoaded = False
        
        ' Set the index
        EditorIndex = frmEditor_Quest.lstIndex.ListIndex + 1
        
        ' set values
        .txtName.Text = Trim$(Quest(EditorIndex).name)
        .txtDescription.Text = Trim$(Quest(EditorIndex).Description)
        
        ' Rewards
        .scrlReward.max = MAX_ITEMS
        .scrlReward.value = Quest(EditorIndex).Reward
        .scrlXP.value = Quest(EditorIndex).XPReward
        
        If Quest(EditorIndex).Reward > 0 Then
            If Item(Quest(EditorIndex).Reward).Type = ITEM_TYPE_CURRENCY Then
                .scrlAmount.value = Quest(EditorIndex).RewardAmount
                .scrlAmount.Visible = True
                .lblAmount.Visible = True
            Else
                .scrlAmount.Visible = False
                .lblAmount.Visible = False
            End If
        Else
            .scrlAmount.Visible = False
        End If
        
        ' xp reward
        If Quest(EditorIndex).XPReward > 0 Then
            .lblXP.Caption = "XP Reward: " & Quest(EditorIndex).XPReward
        Else
            .lblXP.Caption = "XP Reward: None"
        End If
            
        ' Objectives
        For i = 1 To MAX_QUEST_OBJECTIVES
            .cmbObjectiveType(i).ListIndex = Quest(EditorIndex).Objective(i).ObjectiveType
            
            ' Data
            .scrlDataValue(i).value = Quest(EditorIndex).Objective(i).DataAmount
            .scrlDataIndex(i).value = Quest(EditorIndex).Objective(i).DataIndex
        Next
        
        ' Build cmbClassReq
        .cmbClassReq.Clear
        .cmbClassReq.AddItem "None"

        For i = 1 To Max_Classes
            .cmbClassReq.AddItem Class(i).name
        Next
        
        ' Class req
        .cmbClassReq.ListIndex = Quest(EditorIndex).ClassReq
        
        ' Level req
        .scrlLevelReq.value = Quest(EditorIndex).LevelReq
        
        ' Quest req
        .scrlQuestReq.max = MAX_QUESTS
        .scrlQuestReq.value = Quest(EditorIndex).QuestReq
        
        ' Loaded
        QuestEditorLoaded = True
    End With
    
    ' Allows us to clear it if we make a mistake
    Quest_Changed(EditorIndex) = True
End Sub

Public Sub QuestEditorOK()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_QUESTS
        If Quest_Changed(i) Then
            Call SendSaveQuest(i)
        End If
    Next
    
    Unload frmEditor_Quest
    Editor = 0
    ClearChanged_Quest
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "QuestEditorOK", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub QuestEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Unload frmEditor_Quest
    Editor = 0
    ClearChanged_Quest
    
    ClearQuests
    SendRequestQuests
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "QuestEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Quest()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ZeroMemory Quest_Changed(1), MAX_QUESTS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChanged_Quest", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
