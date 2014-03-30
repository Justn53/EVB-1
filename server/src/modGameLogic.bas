Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
Dim i As Long

    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next
End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
Dim i As Long

    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS
        If MapItem(MapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If
    Next
End Function

Function TotalOnlinePlayers() As Long
Dim i As Long

    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If
    Next
End Function

Function FindPlayer(ByVal name As String) As Long
Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(name)))) = UCase$(Trim$(name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString)
Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    Call SpawnItemSlot(i, ItemNum, ItemVal, MapNum, x, y, playerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True)
Dim i As Long
Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
            MapItem(MapNum, i).playerName = playerName
            MapItem(MapNum, i).playerTimer = timeGetTime + ITEM_SPAWN_TIME
            MapItem(MapNum, i).canDespawn = canDespawn
            MapItem(MapNum, i).despawnTimer = timeGetTime + ITEM_DESPAWN_TIME
            MapItem(MapNum, i).Num = ItemNum
            MapItem(MapNum, i).Value = ItemVal
            MapItem(MapNum, i).x = x
            MapItem(MapNum, i).y = y
            ' send to map
            SendSpawnItemToMap MapNum, i
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
Dim x As Long
Dim y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And Map(MapNum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, 1, MapNum, x, y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, Map(MapNum).Tile(x, y).Data2, MapNum, x, y)
                End If
            End If
        Next
    Next
End Sub

Public Sub SpawnNpc(ByVal mapNpcNum As Long, ByVal MapNum As Long)
Dim Buffer As clsBuffer
Dim NPCNum As Long
Dim i As Long
Dim x As Long
Dim y As Long
Dim Spawned As Boolean

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    NPCNum = Map(MapNum).Npc(mapNpcNum)

    If NPCNum > 0 Then
    
        MapNpc(MapNum).Npc(mapNpcNum).Num = NPCNum
        MapNpc(MapNum).Npc(mapNpcNum).target = 0
        MapNpc(MapNum).Npc(mapNpcNum).targetType = 0 ' clear
        
        MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(NPCNum, Vitals.HP)
        MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.MP) = GetNpcMaxVital(NPCNum, Vitals.MP)
        
        MapNpc(MapNum).Npc(mapNpcNum).Dir = Int(Rnd * 4)
        
        ' Check if theres a spawn tile for the specific npc
        For x = 0 To Map(MapNum).MaxX
            For y = 0 To Map(MapNum).MaxY
                If Map(MapNum).Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(MapNum).Tile(x, y).Data1 = mapNpcNum Then
                        MapNpc(MapNum).Npc(mapNpcNum).x = x
                        MapNpc(MapNum).Npc(mapNpcNum).y = y
                        MapNpc(MapNum).Npc(mapNpcNum).Dir = Map(MapNum).Tile(x, y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next y
        Next x
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                x = Rand(0, Map(MapNum).MaxX)
                y = Rand(0, Map(MapNum).MaxY)
    
                If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
                If y > Map(MapNum).MaxY Then y = Map(MapNum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(MapNum, x, y) Then
                    MapNpc(MapNum).Npc(mapNpcNum).x = x
                    MapNpc(MapNum).Npc(mapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To Map(MapNum).MaxX
                For y = 0 To Map(MapNum).MaxY

                    If NpcTileIsOpen(MapNum, x, y) Then
                        MapNpc(MapNum).Npc(mapNpcNum).x = x
                        MapNpc(MapNum).Npc(mapNpcNum).y = y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned And PlayersOnMap(MapNum) > 0 Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).Num
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).Dir
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        End If
        
        SendMapNpcVitals MapNum, mapNpcNum
    End If
End Sub

Public Function NpcTileIsOpen(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(MapNum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = MapNum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(MapNum).Npc(LoopI).Num > 0 Then
            If MapNpc(MapNum).Npc(LoopI).x = x Then
                If MapNpc(MapNum).Npc(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(MapNum).Tile(x, y).Type <> TILE_TYPE_WALKABLE Then
        If Map(MapNum).Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(MapNum).Tile(x, y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal MapNum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal mapNpcNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If

    x = MapNpc(MapNum).Npc(mapNpcNum).x
    y = MapNpc(MapNum).Npc(mapNpcNum).y
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(MapNum).Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(mapNpcNum).x) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(mapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapNpcNum).x) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(mapNpcNum).x) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(mapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(MapNum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(mapNpcNum).x - 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(mapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> mapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(mapNpcNum).x + 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal MapNum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(MapNum).Npc(mapNpcNum).Dir = Dir

    Select Case Dir
        Case DIR_UP
            MapNpc(MapNum).Npc(mapNpcNum).y = MapNpc(MapNum).Npc(mapNpcNum).y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            MapNpc(MapNum).Npc(mapNpcNum).y = MapNpc(MapNum).Npc(mapNpcNum).y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            MapNpc(MapNum).Npc(mapNpcNum).x = MapNpc(MapNum).Npc(mapNpcNum).x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            MapNpc(MapNum).Npc(mapNpcNum).x = MapNpc(MapNum).Npc(mapNpcNum).x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong mapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).Dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select

End Sub

Public Sub NpcDir(ByVal MapNum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(MapNum).Npc(mapNpcNum).Dir = Dir
    
    ' Check if there's at least one player on the map
    If PlayersOnMap(MapNum) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong mapNpcNum
    Buffer.WriteLong Dir
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Sub ClearTempTiles()
    Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next

End Sub

Sub ClearTempTile(ByVal MapNum As Long)
    Dim y As Long
    Dim x As Long
    TempTile(MapNum).DoorTimer = 0
    ReDim TempTile(MapNum).DoorOpen(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            TempTile(MapNum).DoorOpen(x, y) = 0
        Next
    Next

End Sub

Public Sub CacheResources(ByVal MapNum As Long)
    Dim x As Long, y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY

            If Map(MapNum).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(MapNum).ResourceData(0 To Resource_Count)
                ResourceCache(MapNum).ResourceData(Resource_Count).x = x
                ResourceCache(MapNum).ResourceData(Resource_Count).y = y
                ResourceCache(MapNum).ResourceData(Resource_Count).cur_health = Resource(Map(MapNum).Tile(x, y).Data1).health
            End If

        Next
    Next

    ResourceCache(MapNum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(index, oldSlot)
    OldValue = GetPlayerBankItemValue(index, oldSlot)
    NewNum = GetPlayerBankItemNum(index, newSlot)
    NewValue = GetPlayerBankItemValue(index, newSlot)
    
    SetPlayerBankItemNum index, newSlot, OldNum
    SetPlayerBankItemValue index, newSlot, OldValue
    
    SetPlayerBankItemNum index, oldSlot, NewNum
    SetPlayerBankItemValue index, oldSlot, NewValue
        
    SendBank index
End Sub

Sub PlayerSwitchInvSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(index, oldSlot)
    OldValue = GetPlayerInvItemValue(index, oldSlot)
    NewNum = GetPlayerInvItemNum(index, newSlot)
    NewValue = GetPlayerInvItemValue(index, newSlot)
    SetPlayerInvItemNum index, newSlot, OldNum
    SetPlayerInvItemValue index, newSlot, OldValue
    SetPlayerInvItemNum index, oldSlot, NewNum
    SetPlayerInvItemValue index, oldSlot, NewValue
    SendInventory index
End Sub

Sub PlayerSwitchSpellSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(index, oldSlot)
    NewNum = GetPlayerSpell(index, newSlot)
    SetPlayerSpell index, oldSlot, NewNum
    SetPlayerSpell index, newSlot, OldNum
    SendPlayerSpells index
End Sub

Sub PlayerUnequipItem(ByVal index As Long, ByVal EqSlot As Long)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(index, GetPlayerEquipment(index, EqSlot)) > 0 Then
        GiveInvItem index, GetPlayerEquipment(index, EqSlot), 0
        PlayerMsg index, "You unequip " & CheckGrammar(Item(GetPlayerEquipment(index, EqSlot)).name), Yellow
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerEquipment(index, EqSlot)
        ' remove equipment
        SetPlayerEquipment index, 0, EqSlot
        SendWornEquipment index
        SendMapEquipment index
        SendStats index
        ' send vitals
        Call SendVital(index, Vitals.HP)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
    Else
        PlayerMsg index, "Your inventory is full.", BrightRed
    End If

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party functions ##
' #####################

Public Sub Party_PlayerLeave(ByVal index As Long)
Dim partyNum As Long, i As Long

    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' find out how many members we have
        Party_CountMembers partyNum
        ' make sure there's more than 2 people
        If Party(partyNum).MemberCount > 2 Then
        
            ' check if leader
            If Party(partyNum).Leader = index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) > 0 And Party(partyNum).Member(i) <> index Then
                        Party(partyNum).Leader = Party(partyNum).Member(i)
                        PartyMsg partyNum, GetPlayerName(i) & " é o lider do grupo.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partyNum, GetPlayerName(index) & " saiu do grupo.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = index Then
                        Party(partyNum).Member(i) = 0
                        TempPlayer(index).inParty = 0
                        TempPlayer(index).partyInvite = 0
                        Exit For
                        End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            Else
                ' not the leader, just leave
                PartyMsg partyNum, GetPlayerName(index) & " saiu do grupo.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = index Then
                        Party(partyNum).Member(i) = 0
                        TempPlayer(index).inParty = 0
                        TempPlayer(index).partyInvite = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partyNum
            ' only 2 people, disband
            PartyMsg partyNum, "Grupo desfeito.", BrightRed
            ' clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                index = Party(partyNum).Member(i)
                ' player exist?
                If index > 0 Then
                    ' remove them
                    TempPlayer(index).partyInvite = 0
                    TempPlayer(index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo index
                End If
            Next
            ' clear out the party itself
            ClearParty partyNum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if the person is a valid target
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then Exit Sub
    
    ' make sure they're not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Or TempPlayer(targetPlayer).TradeRequest > 0 Then
        ' they've already got a request for trade/party
        PlayerMsg index, "This player is busy.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they're already in a party
        PlayerMsg index, "This player is already in a party.", BrightRed
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(index).inParty > 0 Then
        partyNum = TempPlayer(index).inParty
        ' make sure we're the leader
        If Party(partyNum).Leader = index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partyNum).Member(i) = 0 Then
                    ' send the invitation
                    SendPartyInvite targetPlayer, index
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = index
                    ' let them know
                    PlayerMsg index, "Invitation sent.", Pink
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite targetPlayer, index
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = index
        ' let them know
        PlayerMsg index, "Invitation sent.", Pink
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, i As Long

    ' check if already in a party
    If TempPlayer(index).inParty > 0 Then
        ' get the partynumber
        partyNum = TempPlayer(index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partyNum).Member(i) = 0 Then
                ' add to the party
                Party(partyNum).Member(i) = targetPlayer
                ' recount party
                Party_CountMembers partyNum
                ' send update to all - including new player
                SendPartyUpdate partyNum
                SendPartyVitals partyNum, targetPlayer
                ' let everyone know they've joined
                PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
                ' add them in
                TempPlayer(targetPlayer).inParty = partyNum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg index, "Party is full.", BrightRed
        PlayerMsg targetPlayer, "Party is full.", BrightRed
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                partyNum = i
                Exit For
            End If
        Next
        ' create the party
        Party(partyNum).MemberCount = 2
        Party(partyNum).Leader = index
        Party(partyNum).Member(1) = index
        Party(partyNum).Member(2) = targetPlayer
        SendPartyUpdate partyNum
        SendPartyVitals partyNum, index
        SendPartyVitals partyNum, targetPlayer
        ' let them know it's created
        PartyMsg partyNum, "Party created.", BrightGreen
        PartyMsg partyNum, GetPlayerName(index) & " has joined the party.", Pink
        PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(index).inParty = partyNum
        TempPlayer(targetPlayer).inParty = partyNum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal index As Long, ByVal targetPlayer As Long)
    PlayerMsg index, GetPlayerName(targetPlayer) & " has declined to join the party.", BrightRed
    PlayerMsg targetPlayer, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partyNum As Long)
Dim i As Long, highIndex As Long, x As Long
    ' find the high index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partyNum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partyNum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For x = i To MAX_PARTY_MEMBERS - 1
                    Party(partyNum).Member(x) = Party(partyNum).Member(x + 1)
                    Party(partyNum).Member(x + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partyNum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(partyNum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partyNum
End Sub

Public Sub Party_ShareExp(ByVal partyNum As Long, ByVal exp As Long, ByVal index As Long, ByVal MapNum As Long)
Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long, LoseMemberCount As Byte

    ' check if it's worth sharing
    If Not exp >= Party(partyNum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP index, exp
        Exit Sub
    End If
    
    ' check members in outhers maps
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        If tmpIndex > 0 Then
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                If GetPlayerMap(tmpIndex) <> MapNum Then
                    LoseMemberCount = LoseMemberCount + 1
                End If
            End If
        End If
    Next i
    
    ' find out the equal share
    expShare = exp \ (Party(partyNum).MemberCount - LoseMemberCount)
    leftOver = exp Mod (Party(partyNum).MemberCount - LoseMemberCount)
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                If GetPlayerMap(tmpIndex) = MapNum Then
                    ' give them their share
                    GivePlayerEXP tmpIndex, expShare
                End If
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(partyNum).Member(Rand(1, Party(partyNum).MemberCount))
    ' give the exp
    GivePlayerEXP tmpIndex, leftOver
End Sub

Public Sub GivePlayerEXP(ByVal index As Long, ByVal exp As Long)
    ' give the exp
    Call SetPlayerExp(index, GetPlayerExp(index) + exp)
    SendEXP index
    SendActionMsg GetPlayerMap(index), "+" & exp & " EXP", White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
    ' check if we've leveled
    CheckPlayerLevelUp index
End Sub

Public Sub StartQuest(ByVal index As Long, ByVal questIndex As Long)
Dim i As Long, j As Long

    ' make sure we've got a quest
    If questIndex = 0 Then Exit Sub
    
    ' make sure they meet the requirements
    If Player(index).Level < Quest(questIndex).LevelReq And Quest(questIndex).LevelReq > 0 Then
        Call PlayerMsg(index, "You don't meet the level requirement for this quest!", BrightRed)
        Call SendCloseConv(index)
        Exit Sub
    ElseIf Player(index).Class <> Quest(questIndex).ClassReq And Quest(questIndex).ClassReq > 0 Then
        Call PlayerMsg(index, "You don't meet the class requirement for this quest!", BrightRed)
        Call SendCloseConv(index)
        Exit Sub
    ElseIf Player(index).Quest(questIndex).QuestStatus <> 2 And Quest(questIndex).QuestReq > 0 Then
        Call PlayerMsg(index, "You need to complete '" & Quest(Quest(questIndex).QuestReq).name + "' first!", BrightRed)
        Call SendCloseConv(index)
        Exit Sub
    End If
    
    If Player(index).Quest(questIndex).QuestStatus = 0 Then
        Player(index).Quest(questIndex).QuestStatus = 1
        For i = 1 To 3
            Player(index).Quest(questIndex).DataAmountLeft(i) = Quest(questIndex).Objective(i).DataAmount
        Next
        
        ' check for already existing items
        For i = 1 To MAX_QUEST_OBJECTIVES
            ' get item quest
            If Quest(questIndex).Objective(i).ObjectiveType = 3 Then
                ' see if they've already got it
                For j = 1 To MAX_INV
                    If GetPlayerInvItemNum(index, i) = Quest(questIndex).Objective(i).DataIndex Then
                        ' update the requirement
                        Player(index).Quest(questIndex).DataAmountLeft(j) = Player(index).Quest(questIndex).DataAmountLeft(j) - 1
                    End If
                Next
            End If
        Next
        
        Call SendPlayerData(index)
        Call PlayerMsg(index, "Quest accepted: " & Trim$(Quest(questIndex).name) & ".", Yellow)
        Call SendPopulateList(index)
    Else
        If Player(index).Quest(questIndex).QuestStatus = 1 Then
            Call PlayerMsg(index, "You've already started this quest!", BrightRed)
        ElseIf Player(index).Quest(questIndex).QuestStatus = 2 Then
            Call PlayerMsg(index, "You've already completed this quest!", BrightRed)
        End If
        Call SendCloseConv(index)
    End If
End Sub

Public Sub CheckFinishQuest(ByVal index As Long, ByVal QuestNum As Long)
Dim i As Long
Dim QuestObjectives As Byte
Dim CompletedObjectives As Byte

    ' check to see how many objectives we have
    For i = 1 To MAX_QUEST_OBJECTIVES
        If Quest(QuestNum).Objective(i).DataIndex > 0 Then
            QuestObjectives = QuestObjectives + 1
        End If
    Next
    
    ' check to see how many are finished
    For i = 1 To QuestObjectives
        If Player(index).Quest(QuestNum).DataAmountLeft(i) <= 0 Then
            CompletedObjectives = CompletedObjectives + 1
        End If
    Next
    
    ' check if they're all done
    If CompletedObjectives = QuestObjectives Then
        Call FinishQuest(index, QuestNum)
    Else
        Call PlayerMsg(index, "You have not completed all the quest objectives yet!", BrightRed)
    End If
End Sub

Public Sub FinishQuest(ByVal index As Long, ByVal QuestNum As Long)
Dim i As Long, j As Long

    Player(index).Quest(QuestNum).QuestStatus = 2
    SendQuestFinish index, QuestNum
    SendPopulateList index
    
    ' remove the items
    For i = 1 To MAX_QUEST_OBJECTIVES
        If Quest(QuestNum).Objective(i).ObjectiveType = 3 Then
            If Quest(QuestNum).Objective(i).DataIndex > 0 Then
                For j = 1 To Quest(QuestNum).Objective(i).DataAmount
                    TakeInvItem index, Quest(QuestNum).Objective(i).DataIndex, 1
                Next
            End If
        End If
    Next
    
    ' quest rewards
    If Quest(QuestNum).Reward > 0 Then
        If FindOpenInvSlot(index, Quest(QuestNum).Reward) > 0 Then
            Call GiveInvItem(index, Quest(QuestNum).Reward, Quest(QuestNum).RewardAmount)
        Else
            Call SpawnItem(Quest(QuestNum).Reward, Quest(QuestNum).RewardAmount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            Call PlayerMsg(index, "Your reward fell to the floor.", Cyan)
        End If
    End If
    
    If Quest(QuestNum).XPReward > 0 Then
        Call GivePlayerEXP(index, Quest(QuestNum).XPReward)
        Call PlayerMsg(index, "You received: " & Quest(QuestNum).RewardAmount & " " & Trim$(Item(Quest(QuestNum).Reward).name) & " and " & Quest(QuestNum).XPReward & " XP", BrightGreen)
    Else
        Call PlayerMsg(index, "You received: " & Quest(QuestNum).RewardAmount & " " & Trim$(Item(Quest(QuestNum).Reward).name), BrightGreen)
    End If
    
    Call PlayerMsg(index, "Quest completed: " & Trim$(Quest(QuestNum).name) & "!", Yellow)
End Sub
