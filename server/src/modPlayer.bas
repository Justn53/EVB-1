Attribute VB_Name = "modPlayer"
Option Explicit

Public Sub HandleUseChar(ByVal index As Long)
    If Not IsPlaying(index) Then
        Call JoinGame(index)
        Call TextAdd(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Game_Name & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal index As Long)
    Dim i As Long
    
    ' Set the flag so we know the person is in the game
    TempPlayer(index).InGame = True
    
    ' Update the log
    frmServer.lvwInfo.ListItems(index).SubItems(1) = GetPlayerIP(index)
    frmServer.lvwInfo.ListItems(index).SubItems(2) = GetPlayerLogin(index)
    frmServer.lvwInfo.ListItems(index).SubItems(3) = GetPlayerName(index)
    
    ' send the login ok
    SendLoginOk index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendAnimations(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendResources(index)
    Call SendInventory(index)
    Call SendWornEquipment(index)
    Call SendMapEquipment(index)
    Call SendPlayerSpells(index)
    Call SendHotbar(index)
    Call SendConvs(index)
    Call SendQuests(index)
    
    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
    Call SendEXP(index)
    Call SendStats(index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(index) <= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", White)
    End If
    
    ' Send welcome messages
    Call SendWelcome(index)

    ' Send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, i
    Next
    
    ' Send the flag so they know they can start doing stuff
    SendInGame index
End Sub

Public Sub LeftGame(ByVal index As Long)
Dim n As Long, i As Long
Dim tradeTarget As Long
    
    If TempPlayer(index).InGame Then
        TempPlayer(index).InGame = False
        
        ' Get rid of their targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    If TempPlayer(i).targetType = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = index Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
        
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) < 1 Then
            PlayersOnMap(GetPlayerMap(index)) = 0
        End If
        
        ' cancel any trade they're in
        If TempPlayer(index).InTrade > 0 Then
            tradeTarget = TempPlayer(index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " has declined the trade.", BrightRed
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).Num = 0
                TempPlayer(tradeTarget).TradeOffer(i).Value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' clear target
        For i = 1 To Player_HighIndex
            ' Prevent subscript out range
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(index) Then
                ' clear players target
                If TempPlayer(i).targetType = TARGET_TYPE_PLAYER And TempPlayer(i).target = index Then
                    TempPlayer(i).target = 0
                    TempPlayer(i).targetType = TARGET_TYPE_NONE
                    SendTarget i
                End If
            End If
        Next
        
        ' leave party.
        Party_PlayerLeave index

        ' save and clear data.
        Call SavePlayer(index)
        Call SaveBank(index)
        Call ClearBank(index)

        ' Send a global message that he/she left
        If GetPlayerAccess(index) <= ADMIN_MONITOR Then
            Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", White)
        End If

        Call TextAdd(GetPlayerName(index) & " has disconnected from " & Options.Game_Name & ".")
        Call SendLeftGame(index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(index)
End Sub

Function GetPlayerProtection(ByVal index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long
    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > Player_HighIndex Then
        Exit Function
    End If

    Armor = GetPlayerEquipment(index, Armor)
    Helm = GetPlayerEquipment(index, Helmet)
    GetPlayerProtection = (GetPlayerStat(index, Stats.Endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
    End If

End Function

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim n As Long

    If GetPlayerEquipment(index, Weapon) > 0 Then
        n = (Rnd) * 2

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.Strength) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(index, Shield)

    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            i = (GetPlayerStat(index, Stats.Endurance) \ 2) + (GetPlayerLevel(index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function

Public Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim shopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
    If y > Map(MapNum).MaxY Then y = Map(MapNum).MaxY
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    ' if same map then just send their co-ordinates
    If MapNum = GetPlayerMap(index) Then
        SendPlayerXYToMap index
    End If
    
    ' clear target
    For i = 1 To Player_HighIndex
        ' Prevent subscript out range
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(index) Then
            If TempPlayer(i).targetType = TARGET_TYPE_PLAYER And TempPlayer(i).target = index Then
                TempPlayer(i).target = 0
                TempPlayer(i).targetType = TARGET_TYPE_NONE
                SendTarget i
            End If
        End If
    Next
    
    ' clear target
    TempPlayer(index).target = 0
    TempPlayer(index).targetType = TARGET_TYPE_NONE
    SendTarget index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)

    If OldMap <> MapNum Then
        Call SendLeaveMap(index, OldMap)
    End If

    Call SetPlayerMap(index, MapNum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)
    
    ' send player's equipment to new map
    SendMapEquipment index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(MapNum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = MapNum Then
                    SendMapEquipmentTo i, index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = 0

        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(OldMap).Npc(i).Num > 0 Then
                MapNpc(OldMap).Npc(i).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(OldMap).Npc(i).Num, Vitals.HP)
            End If
        Next
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = 1
    TempPlayer(index).GettingMap = 1
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong MapNum
    Buffer.WriteLong Map(MapNum).Revision
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, MapNum As Long
    Dim x As Long, y As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte
    Dim TileType As Long, VitalType As Long, Colour As Long, amount As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    Moved = 0
    MapNum = GetPlayerMap(index)
    
    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_UP + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_RESOURCE Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) - 1) = 1) Then
                                Call SetPlayerY(index, GetPlayerY(index) - 1)
                                SendPlayerMove index, movement, sendToSelf
                                Moved = 1
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Up > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(index)).Up).MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), NewMapY)
                    Moved = 1
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < Map(MapNum).MaxY Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_RESOURCE Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index), GetPlayerY(index) + 1) = 1) Then
                                Call SetPlayerY(index, GetPlayerY(index) + 1)
                                SendPlayerMove index, movement, sendToSelf
                                Moved = 1
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Down > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                    Moved = 1
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_LEFT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) - 1, GetPlayerY(index)) = 1) Then
                                Call SetPlayerX(index, GetPlayerX(index) - 1)
                                SendPlayerMove index, movement, sendToSelf
                                Moved = 1
                            End If
                        End If
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Left > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(index)).Left).MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Left, NewMapX, GetPlayerY(index))
                    Moved = 1
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < Map(MapNum).MaxX Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_RIGHT + 1) Then
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
    
                            ' Check to see if the tile is a key and if it is check if its opened
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(GetPlayerX(index) + 1, GetPlayerY(index)) = 1) Then
                                Call SetPlayerX(index, GetPlayerX(index) + 1)
                                SendPlayerMove index, movement, sendToSelf
                                Moved = 1
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Right > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                    Moved = 1
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
    End Select
    
    With Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            MapNum = .Data1
            x = .Data2
            y = .Data3
            Call PlayerWarp(index, MapNum, x, y)
            Moved = 1
        End If

        ' Check for key trigger open
        If .Type = TILE_TYPE_KEYOPEN Then
            x = .Data1
            y = .Data2
    
            If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(x, y) = 0 Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = 1
                TempTile(GetPlayerMap(index)).DoorTimer = timeGetTime
                SendMapKey index, x, y, 1
                Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
            End If
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            x = .Data1
            If x > 0 Then ' shop exists?
                If Len(Trim$(Shop(x).name)) > 0 Then ' name exists?
                    SendOpenShop index, x
                    TempPlayer(index).InShop = x ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank index
            TempPlayer(index).InBank = True
            Moved = 1
        End If
        
        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            VitalType = .Data1
            amount = .Data2
            If Not GetPlayerVital(index, VitalType) = GetPlayerMaxVital(index, VitalType) Then
                If VitalType = Vitals.HP Then
                    Colour = BrightGreen
                Else
                    Colour = BrightBlue
                End If
                SendActionMsg GetPlayerMap(index), "+" & amount, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
                SetPlayerVital index, VitalType, GetPlayerVital(index, VitalType) + amount
                PlayerMsg index, "You feel rejuvenating forces flowing through your body.", BrightGreen
                Call SendVital(index, VitalType)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = 1
        End If
        
        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            amount = .Data1
            SendActionMsg GetPlayerMap(index), "-" & amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
            If GetPlayerVital(index, HP) - amount <= 0 Then
                KillPlayer index
                PlayerMsg index, "You were killed by a trap.", BrightRed
            Else
                SetPlayerVital index, HP, GetPlayerVital(index, HP) - amount
                PlayerMsg index, "You were injured by a trap.", BrightRed
                Call SendVital(index, HP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = 1
        End If
        
        ' Slide
         If .Type = TILE_TYPE_SLIDE Then
            Select Case .Data1
                Case DIR_UP
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_RESOURCE Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_BLOCKED Then Exit Sub
                Case DIR_LEFT
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_RESOURCE Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_BLOCKED Then Exit Sub
                Case DIR_DOWN
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_RESOURCE Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_BLOCKED Then Exit Sub
                Case DIR_RIGHT
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_RESOURCE Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_BLOCKED Then Exit Sub
            End Select
            ForcePlayerMove index, MOVING_WALKING, .Data1
            Moved = 1
         End If
    End With

    ' They tried to hack
    If Moved = 0 Then
        PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
    End If

End Sub

Sub ForcePlayerMove(ByVal index As Long, ByVal movement As Long, ByVal Direction As Long)
    If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub
    
    Select Case Direction
        Case DIR_UP
            If GetPlayerY(index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
    End Select
    
    PlayerMove index, Direction, movement, True
End Sub

Sub CheckEquippedItems(ByVal index As Long)
    Dim Slot As Long
    Dim ItemNum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(index, i)

        If ItemNum > 0 Then

            Select Case i
                Case Equipment.Weapon

                    If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment index, 0, i
                Case Equipment.Armor

                    If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment index, 0, i
                Case Equipment.Helmet

                    If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment index, 0, i
                Case Equipment.Shield

                    If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment index, 0, i
            End Select

        Else
            SetPlayerEquipment index, 0, i
        End If

    Next

End Sub

Function FindOpenInvSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If

        Next

    End If

    For i = 1 To MAX_INV

        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    If Not IsPlaying(index) Then Exit Function
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

End Function

Function HasItem(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function TakeInvItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
                Exit Function
            End If
        End If

    Next

End Function

Function TakeInvSlot(ByVal index As Long, ByVal invSlot As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ItemNum
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invSlot <= 0 Or invSlot > MAX_ITEMS Then
        Exit Function
    End If
    
    ItemNum = GetPlayerInvItemNum(index, invSlot)

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(index, invSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(index, invSlot, GetPlayerInvItemValue(index, invSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(index, invSlot, 0)
        Call SetPlayerInvItemValue(index, invSlot, 0)
        Exit Function
    End If

End Function

Function GiveInvItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, Optional ByVal sendUpdate As Boolean = True) As Boolean
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(index, ItemNum)

    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(index, i, ItemNum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        If sendUpdate Then Call SendInventoryUpdate(index, i)
        GiveInvItem = True
    Else
        Call PlayerMsg(index, "Your inventory is full.", BrightRed)
        GiveInvItem = False
    End If

End Function

Function HasSpell(ByVal index As Long, ByVal spellNum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(index, i) = spellNum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal index As Long)
Dim i As Long, j As Long, k As Long
Dim n As Long
Dim MapNum As Long
Dim Msg As String
Dim ItemName As String, Progress As Long, OutOf As String

    If Not IsPlaying(index) Then Exit Sub
    MapNum = GetPlayerMap(index)

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(MapNum, i).x = GetPlayerX(index)) Then
                    If (MapItem(MapNum, i).y = GetPlayerY(index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(index, MapItem(MapNum, i).Num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventory
                            Call SetPlayerInvItemNum(index, n, MapItem(MapNum, i).Num)
    
                            If Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Then
                                Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(MapNum, i).Value)
                                Msg = MapItem(MapNum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(index, n)).name)
                            Else
                                Call SetPlayerInvItemValue(index, n, 0)
                                Msg = Trim$(Item(GetPlayerInvItemNum(index, n)).name)
                            End If
                            
                            ' quests
                            For j = 1 To MAX_QUESTS
                                For k = 1 To MAX_QUEST_OBJECTIVES
                                    ' get item quest
                                    If Player(index).Quest(j).QuestStatus = 1 And Quest(j).Objective(k).ObjectiveType = 3 Then
                                        ' make sure it's the right item
                                        If MapItem(MapNum, i).Num = Quest(j).Objective(k).DataIndex Then
                                            ' update the requirement
                                            Player(index).Quest(j).DataAmountLeft(k) = Player(index).Quest(j).DataAmountLeft(k) - 1
                                            
                                            ' update them
                                            ItemName = Trim$(Item(Quest(j).Objective(k).DataIndex).name)
                                            Progress = Quest(j).Objective(k).DataAmount - Player(index).Quest(j).DataAmountLeft(k)
                                            OutOf = CStr(Quest(j).Objective(k).DataAmount)
                                            
                                            If Progress <= Quest(j).Objective(k).DataAmount Then
                                                PlayerMsg index, Trim$(Quest(j).name) & ": " & ItemName & ": " & Progress & "/" & OutOf, Yellow
                                            End If
                                        End If
                                    End If
                                Next
                            Next
                            
                            ' Erase item from the map
                            ClearMapItem i, MapNum
                            
                            Call SendInventoryUpdate(index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), 0, 0)
                            SendActionMsg GetPlayerMap(index), Msg, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                            Exit For
                        Else
                            Call PlayerMsg(index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Function CanPlayerPickupItem(ByVal index As Long, ByVal mapItemNum As Long)
Dim MapNum As Long

    MapNum = GetPlayerMap(index)
    
    ' no lock or locked to player?
    If MapItem(MapNum, mapItemNum).playerName = vbNullString Or MapItem(MapNum, mapItemNum).playerName = Trim$(GetPlayerName(index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    CanPlayerPickupItem = False
End Function

Public Sub PlayerMapDropItem(ByVal index As Long, ByVal invNum As Long, ByVal amount As Long)
Dim i As Long, j As Long, k As Long
Dim ItemName As String, Progress As Long, OutOf As String

    ' Check for subscript out of range
    If IsPlaying(index) = False Or invNum <= 0 Or invNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Or TempPlayer(index).InTrade > 0 Or TempPlayer(index).InChat > 0 Then Exit Sub
    
    If (GetPlayerInvItemNum(index, invNum) > 0) Then
        If (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
            i = FindOpenMapItemSlot(GetPlayerMap(index))
            
            If Item(GetPlayerInvItemNum(index, invNum)).Tradable = False Then
                Call PlayerMsg(index, "You cannot drop this item!", BrightRed)
                Exit Sub
            End If

            If i <> 0 Then
                MapItem(GetPlayerMap(index), i).Num = GetPlayerInvItemNum(index, invNum)
                MapItem(GetPlayerMap(index), i).x = GetPlayerX(index)
                MapItem(GetPlayerMap(index), i).y = GetPlayerY(index)
                MapItem(GetPlayerMap(index), i).playerName = Trim$(GetPlayerName(index))
                MapItem(GetPlayerMap(index), i).playerTimer = timeGetTime + ITEM_SPAWN_TIME
                MapItem(GetPlayerMap(index), i).canDespawn = True
                MapItem(GetPlayerMap(index), i).despawnTimer = timeGetTime + ITEM_DESPAWN_TIME

                If Item(GetPlayerInvItemNum(index, invNum)).Type = ITEM_TYPE_CURRENCY Then

                    ' Check if its more then they have and if so drop it all
                    If amount >= GetPlayerInvItemValue(index, invNum) Then
                        MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, invNum)
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " dropped " & GetPlayerInvItemValue(index, invNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).name) & ".", Yellow)
                        Call SetPlayerInvItemNum(index, invNum, 0)
                        Call SetPlayerInvItemValue(index, invNum, 0)
                    Else
                        MapItem(GetPlayerMap(index), i).Value = amount
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " dropped " & amount & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).name) & ".", Yellow)
                        Call SetPlayerInvItemValue(index, invNum, GetPlayerInvItemValue(index, invNum) - amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(index), i).Value = 0
                    ' send message
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " dropped " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(index, invNum)).name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, invNum, 0)
                    Call SetPlayerInvItemValue(index, invNum, 0)
                End If

                ' Send inventory update
                Call SendInventoryUpdate(index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).Num, amount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), Trim$(GetPlayerName(index)), MapItem(GetPlayerMap(index), i).canDespawn)
            
                ' quests
                For j = 1 To MAX_QUESTS
                    For k = 1 To MAX_QUEST_OBJECTIVES
                        ' get item quest
                        If Player(index).Quest(j).QuestStatus = 1 And Quest(j).Objective(k).ObjectiveType = 3 Then
                            ' make sure it's the right item
                            If MapItem(GetPlayerMap(index), i).Num = Quest(j).Objective(k).DataIndex Then
                                ' update the requirement
                                If Player(index).Quest(j).DataAmountLeft(k) < Quest(j).Objective(k).DataAmount Then
                                    Player(index).Quest(j).DataAmountLeft(k) = Player(index).Quest(j).DataAmountLeft(k) + 1
                                End If
                                
                                ' update them
                                ItemName = Trim$(Item(Quest(j).Objective(k).DataIndex).name)
                                Progress = Quest(j).Objective(k).DataAmount - Player(index).Quest(j).DataAmountLeft(k)
                                OutOf = CStr(Quest(j).Objective(k).DataAmount)
                                
                                If Progress <= Quest(j).Objective(k).DataAmount Then
                                    PlayerMsg index, Trim$(Quest(j).name) & ": " & ItemName & ": " & Progress & "/" & OutOf, Yellow
                                End If
                            End If
                        End If
                    Next
                Next
            Else
                Call PlayerMsg(index, "Too many items already on the ground.", BrightRed)
            End If
        End If
    End If

End Sub

Public Sub CheckPlayerLevelUp(ByVal index As Long)
Dim i As Long
Dim expRollover As Long
Dim LevelCount As Long
    
    LevelCount = 0
    
    Do While GetPlayerExp(index) >= GetPlayerNextLevel(index)
        expRollover = GetPlayerExp(index) - GetPlayerNextLevel(index)
        
        ' can level up?
        If Not SetPlayerLevel(index, GetPlayerLevel(index) + 1) Then
            Exit Sub
        End If
        
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 3)
        Call SetPlayerExp(index, expRollover)
        LevelCount = LevelCount + 1
    Loop
    
    If LevelCount > 0 Then
        If LevelCount = 1 Then
            ' singular
            GlobalMsg GetPlayerName(index) & " has gained " & LevelCount & " level!", Brown
        Else
            ' plural
            GlobalMsg GetPlayerName(index) & " has gained " & LevelCount & " levels!", Brown
        End If

        For i = 1 To Vitals.Vital_Count - 1
            SendVital index, i
        Next
        
        SendEXP index
        SendPlayerData index
    End If
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim$(Player(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    Player(index).Login = Login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
    GetPlayerPassword = Trim$(Player(index).Password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
    Player(index).Password = Password
End Sub

Function GetPlayerName(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(index).name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal name As String)
    Player(index).name = name
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(index).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(index).Level
End Function

Function SetPlayerLevel(ByVal index As Long, ByVal Level As Long) As Boolean
    SetPlayerLevel = False
    If Level > MAX_LEVELS Then Exit Function
    Player(index).Level = Level
    SetPlayerLevel = True
End Function

Function GetPlayerNextLevel(ByVal index As Long) As Long
    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(index) + 1) ^ 3 - (6 * (GetPlayerLevel(index) + 1) ^ 2) + 17 * (GetPlayerLevel(index) + 1) - 12)
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = Player(index).exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal exp As Long)
    Player(index).exp = exp
End Sub

Function GetPlayerDataAmountLeft(ByVal index As Long, ByVal QuestNum As Long, ByVal ObjectiveNum As Long) As Long
    GetPlayerDataAmountLeft = Player(index).Quest(QuestNum).DataAmountLeft(ObjectiveNum)
End Function

Function GetPlayerQuestStatus(ByVal index As Long, ByVal QuestNum As Long) As Long
    GetPlayerQuestStatus = Player(index).Quest(QuestNum).QuestStatus
End Function

Function GetPlayerAccess(ByVal index As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(index).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Access = Access
End Sub

Function GetPlayerVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(index).Vital(Vital) = Value

    If GetPlayerVital(index, Vital) > GetPlayerMaxVital(index, Vital) Then
        Player(index).Vital(Vital) = GetPlayerMaxVital(index, Vital)
    End If

    If GetPlayerVital(index, Vital) < 0 Then
        Player(index).Vital(Vital) = 0
    End If

End Sub

Public Function GetPlayerStat(ByVal index As Long, ByVal Stat As Stats) As Long
    Dim x As Long, i As Long
    If index > MAX_PLAYERS Then Exit Function
    
    x = Player(index).Stat(Stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If Player(index).Equipment(i) > 0 Then
            If Item(Player(index).Equipment(i)).Add_Stat(Stat) > 0 Then
                x = x + Item(Player(index).Equipment(i)).Add_Stat(Stat)
            End If
        End If
    Next
    
    GetPlayerStat = x
End Function

Public Function GetPlayerRawStat(ByVal index As Long, ByVal Stat As Stats) As Long
    If index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStat = Player(index).Stat(Stat)
End Function

Public Sub SetPlayerStat(ByVal index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(index).Stat(Stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    If POINTS <= 0 Then POINTS = 0
    Player(index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(index).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)

    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(index).Map = MapNum
    End If

End Sub

Function GetPlayerX(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(index).x
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    Player(index).x = x
End Sub

Function GetPlayerY(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(index).y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    Player(index).y = y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(index).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(index).Inv(invSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long, ByVal ItemNum As Long)
    Player(index).Inv(invSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Player(index).Inv(invSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    Player(index).Inv(invSlot).Value = ItemValue
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal spellslot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Player(index).Spell(spellslot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal spellslot As Long, ByVal spellNum As Long)
    Player(index).Spell(spellslot) = spellNum
End Sub

Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Equipment) As Long

    If index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipment = Player(index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    Player(index).Equipment(EquipmentSlot) = invNum
End Sub

Public Sub OnDeath(ByVal index As Long)
Dim i As Long, j As Long
Dim RandSlot(1 To 3) As Long
    
    ' Set HP to nothing
    Call SetPlayerVital(index, Vitals.HP, 0)
    
    ' Get rid of their targets
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And IsConnected(i) Then
            If GetPlayerMap(i) = GetPlayerMap(index) Then
                If TempPlayer(i).targetType = TARGET_TYPE_PLAYER Then
                    If TempPlayer(i).target = index Then
                        TempPlayer(i).target = 0
                        TempPlayer(i).targetType = TARGET_TYPE_NONE
                        SendTarget i
                    End If
                End If
            End If
        End If
    Next
    
    If TotalInventoryItems(index) > 3 Then
        ' Get three random items to drop
        For i = 1 To 3
            RandSlot(i) = Rand(1, TotalInventoryItems(index))
        Next
        
        ' See if they're the same, if they are loop to find ones that aren't
        Do Until RandSlot(1) <> RandSlot(2) <> RandSlot(3)
            For i = 1 To 3
                RandSlot(i) = Rand(1, TotalInventoryItems(index))
            Next
        Loop
        
        ' Drop three random items
        For i = 1 To MAX_INV
            If i = RandSlot(1) Or i = RandSlot(2) Or i = RandSlot(3) Then
                ' Drop it if it isn't gold
                If GetPlayerInvItemNum(index, i) <> 1 Then PlayerMapDropItem index, i, 0
            End If
        Next
    Else
        ' Drop everything
        For i = 1 To MAX_INV
            ' Drop it if it isn't gold
            If GetPlayerInvItemNum(index, i) <> 1 Then PlayerMapDropItem index, i, 0
        Next
    End If
     
    ' Warp player away
    Call SetPlayerDir(index, DIR_DOWN)
    
    With Map(GetPlayerMap(index))
        ' to the bootmap if it is set
        If .BootMap > 0 Then
            PlayerWarp index, .BootMap, .BootX, .BootY
        Else
            Call PlayerWarp(index, Options.StartMap, Options.StartX, Options.StartY)
        End If
    End With
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    TempPlayer(index).spellBuffer.Spell = 0
    TempPlayer(index).spellBuffer.Timer = 0
    TempPlayer(index).spellBuffer.target = 0
    TempPlayer(index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(index)
    
    TempPlayer(index).InBank = False
    TempPlayer(index).InChat = 0
    TempPlayer(index).InShop = 0
    
    ' Clear trade
    If TempPlayer(index).InTrade > 0 Then
        For i = 1 To MAX_INV
            TempPlayer(index).TradeOffer(i).Num = 0
            TempPlayer(index).TradeOffer(i).Value = 0
            TempPlayer(TempPlayer(index).InTrade).TradeOffer(i).Num = 0
            TempPlayer(TempPlayer(index).InTrade).TradeOffer(i).Value = 0
        Next
        
        TempPlayer(index).InTrade = 0
        TempPlayer(TempPlayer(index).InTrade).InTrade = 0
        
        SendCloseTrade index
        SendCloseTrade TempPlayer(index).InTrade
    End If
    
    ' Restore vitals
    Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
    Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))
    Call SendVital(index, Vitals.HP)
    Call SendVital(index, Vitals.MP)
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
End Sub

Public Sub CheckResource(ByVal index As Long, ByVal x As Long, ByVal y As Long)
Dim Resource_num As Long
Dim Resource_index As Long
Dim rX As Long, rY As Long
Dim i As Long, j As Long
Dim Damage As Long
Dim ResourceName As String, Progress As Long, OutOf As String
   
    ' Check the attack timer
    If GetPlayerEquipment(index, Weapon) > 0 Then
        If timeGetTime < TempPlayer(index).AttackTimer + Item(GetPlayerEquipment(index, Weapon)).Speed Then Exit Sub
    Else
        If timeGetTime < TempPlayer(index).AttackTimer + 1000 Then Exit Sub
    End If
    
    If Map(GetPlayerMap(index)).Tile(x, y).Type <> TILE_TYPE_RESOURCE Then Exit Sub
   
    Resource_num = 0
    Resource_index = Map(GetPlayerMap(index)).Tile(x, y).Data1
    ' Get the cache number
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        If ResourceCache(GetPlayerMap(index)).ResourceData(i).x = x Then
            If ResourceCache(GetPlayerMap(index)).ResourceData(i).y = y Then
                Resource_num = i
            End If
        End If
    Next
    
    If Resource_num > 0 Then
        If Resource(Resource_index).ToolRequired > 0 Then
            If GetPlayerEquipment(index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(index, Weapon)).Data3 <> Resource(Resource_index).ToolRequired Then
                    PlayerMsg index, "You have the wrong type of tool equiped.", BrightRed
                    Exit Sub
                Else
                    Damage = Rand(1, Item(GetPlayerEquipment(index, Weapon)).Data2)
                End If
            Else
                PlayerMsg index, "You need a tool to collect this resource!", BrightRed
                Exit Sub
            End If
        Else
            If GetPlayerEquipment(index, Weapon) > 0 Then
                Damage = Rand(1, Item(GetPlayerEquipment(index, Weapon)).Data2)
            Else
                Damage = Rand(1, (GetPlayerStat(index, Stats.Strength) / 5))
            End If
        End If
                   
        ' inv space?
        If Resource(Resource_index).ItemReward > 0 Then
            If FindOpenInvSlot(index, Resource(Resource_index).ItemReward) = 0 Then
                PlayerMsg index, "You have no inventory space.", BrightRed
                Exit Sub
                End If
            End If
        ' check if already cut down
        If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 0 Then
                                   
            rX = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).x
            rY = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).y
            ' check if damage is more than health
            If Damage > 0 Then
                ' cut it down!
                If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                    SendActionMsg GetPlayerMap(index), "-" & ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32)
                    ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                    ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceTimer = timeGetTime
                    SendResourceCacheToMap GetPlayerMap(index), Resource_num
                    ' send message if it exists
                    If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                        SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                    End If
                                                
                    ' quests
                    For i = 1 To MAX_QUESTS
                        For j = 1 To MAX_QUEST_OBJECTIVES
                            ' resource quest
                            If Player(index).Quest(i).QuestStatus = 1 And Quest(i).Objective(j).ObjectiveType = 4 Then
                                ' make sure it's the right resource
                                If Resource_index = Quest(i).Objective(j).DataIndex Then
                                    ' update the requirement
                                    Player(index).Quest(i).DataAmountLeft(j) = Player(index).Quest(i).DataAmountLeft(j) - 1
                                    
                                    ' update them
                                    ResourceName = Trim$(Resource(Quest(i).Objective(j).DataIndex).name)
                                    Progress = Quest(i).Objective(j).DataAmount - Player(index).Quest(i).DataAmountLeft(j)
                                    OutOf = CStr(Quest(i).Objective(j).DataAmount)
                                    
                                    If Progress <= Quest(i).Objective(j).DataAmount Then
                                        PlayerMsg index, Trim$(Quest(i).name) & ": " & ResourceName & ": " & Progress & "/" & OutOf, Yellow
                                    End If
                                End If
                            End If
                        Next
                    Next
                    Call StartQuest(index, Resource(Resource_index).questIndex)
                                                
                    ' carry on
                    GiveInvItem index, Resource(Resource_index).ItemReward, 1
                    SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                Else
                
                    ' just do the damage
                    ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage
                    SendActionMsg GetPlayerMap(index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32)
                    SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
                End If
                    ' reset the attack timer
                    TempPlayer(index).AttackTimer = timeGetTime
                    ' send the sound
                    SendMapSound index, rX, rY, SoundEntity.seResource, Resource_index
            Else
                ' too weak
                SendActionMsg GetPlayerMap(index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
            End If
        Else
            ' send message if it exists
            If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
            End If
        End If
    End If
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Bank(index).Item(BankSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Bank(index).Item(BankSlot).Num = ItemNum
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Bank(index).Item(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Bank(index).Item(BankSlot).Value = ItemValue
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal invSlot As Long, ByVal amount As Long)
Dim BankSlot As Long

    If invSlot < 0 Or invSlot > MAX_INV Then
        Exit Sub
    End If
    
    If amount < 0 Or amount > GetPlayerInvItemValue(index, invSlot) Then
        Exit Sub
    End If
    
    If Item(GetPlayerInvItemNum(index, invSlot)).Type = ITEM_TYPE_CURRENCY Then
    If amount < 1 Then Exit Sub
    End If
    
    BankSlot = FindOpenBankSlot(index, GetPlayerInvItemNum(index, invSlot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(index, invSlot)).Type = ITEM_TYPE_CURRENCY Then
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), amount)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), amount)
            End If
        Else
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
                Call SetPlayerBankItemValue(index, BankSlot, 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index
End Sub

Sub TakeBankItem(ByVal index As Long, ByVal BankSlot As Long, ByVal amount As Long)
Dim invSlot As Long

    If BankSlot < 0 Or BankSlot > MAX_BANK Then
        Exit Sub
    End If
    
    If amount < 0 Or amount > GetPlayerBankItemValue(index, BankSlot) Then
        Exit Sub
    End If
    
    If Item(GetPlayerBankItemNum(index, BankSlot)).Type = ITEM_TYPE_CURRENCY Then
    If amount < 1 Then Exit Sub
    End If
    
    invSlot = FindOpenInvSlot(index, GetPlayerBankItemNum(index, BankSlot))
        
    If invSlot > 0 Then
        If Item(GetPlayerBankItemNum(index, BankSlot)).Type = ITEM_TYPE_CURRENCY Then
            Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), amount)
            Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - amount)
            If GetPlayerBankItemValue(index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(index, BankSlot) > 1 Then
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - 1)
            Else
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
            End If
        End If
    End If
    
    SaveBank index
    SavePlayer index
    SendBank index

End Sub

Public Sub KillPlayer(ByVal index As Long)
Dim exp As Long

    ' Calculate exp to give attacker
    exp = GetPlayerExp(index) \ 3

    ' Make sure we dont get less then 0
    If exp < 0 Then exp = 0
    If exp = 0 Then
        Call PlayerMsg(index, "You lost no exp.", BrightRed)
    Else
        Call SetPlayerExp(index, GetPlayerExp(index) - exp)
        SendEXP index
        Call PlayerMsg(index, "You lost " & exp & " exp.", BrightRed)
    End If
    
    Call OnDeath(index)
End Sub

Public Sub UseItem(ByVal index As Long, ByVal invNum As Long)
Dim n As Long, i As Long, tempItem As Long, x As Long, y As Long, ItemNum As Long

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_ITEMS Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(index, invNum) > 0) And (GetPlayerInvItemNum(index, invNum) <= MAX_ITEMS) Then
        n = Item(GetPlayerInvItemNum(index, invNum)).Data2
        ItemNum = GetPlayerInvItemNum(index, invNum)
        
        ' Find out what kind of item it is
        Select Case Item(ItemNum).Type
            Case ITEM_TYPE_ARMOR
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(ItemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(ItemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Armor) > 0 Then
                    tempItem = GetPlayerEquipment(index, Armor)
                End If

                SetPlayerEquipment index, ItemNum, Armor
                PlayerMsg index, "You equip " & CheckGrammar(Item(ItemNum).name), BrightGreen
                TakeInvItem index, ItemNum, 0

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_WEAPON
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(ItemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(ItemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Weapon) > 0 Then
                    tempItem = GetPlayerEquipment(index, Weapon)
                End If

                SetPlayerEquipment index, ItemNum, Weapon
                PlayerMsg index, "You equip " & CheckGrammar(Item(ItemNum).name), BrightGreen
                TakeInvItem index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_HELMET
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(ItemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(ItemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Helmet) > 0 Then
                    tempItem = GetPlayerEquipment(index, Helmet)
                End If

                SetPlayerEquipment index, ItemNum, Helmet
                PlayerMsg index, "You equip " & CheckGrammar(Item(ItemNum).name), BrightGreen
                TakeInvItem index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_SHIELD
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(ItemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(ItemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
                    Exit Sub
                End If

                If GetPlayerEquipment(index, Shield) > 0 Then
                    tempItem = GetPlayerEquipment(index, Shield)
                End If

                SetPlayerEquipment index, ItemNum, Shield
                PlayerMsg index, "You equip " & CheckGrammar(Item(ItemNum).name), BrightGreen
                TakeInvItem index, ItemNum, 1

                If tempItem > 0 Then
                    GiveInvItem index, tempItem, 0 ' give back the stored item
                    tempItem = 0
                End If
                
                ' send vitals
                Call SendVital(index, Vitals.HP)
                Call SendVital(index, Vitals.MP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

                Call SendWornEquipment(index)
                Call SendMapEquipment(index)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            ' consumable
            Case ITEM_TYPE_CONSUME
                If TempPlayer(index).FreeAction = False Then Exit Sub

                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(ItemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(ItemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' add hp
                If Item(ItemNum).AddHP > 0 Then
                    Player(index).Vital(Vitals.HP) = Player(index).Vital(Vitals.HP) + Item(ItemNum).AddHP
                    SendActionMsg GetPlayerMap(index), "+" & Item(ItemNum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, HP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add mp
                If Item(ItemNum).AddMP > 0 Then
                    Player(index).Vital(Vitals.MP) = Player(index).Vital(Vitals.MP) + Item(ItemNum).AddMP
                    SendActionMsg GetPlayerMap(index), "+" & Item(ItemNum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendVital index, MP
                    ' send vitals to party if in one
                    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                End If
                ' add exp
                If Item(ItemNum).AddEXP > 0 Then
                    SetPlayerExp index, GetPlayerExp(index) + Item(ItemNum).AddEXP
                    CheckPlayerLevelUp index
                    SendActionMsg GetPlayerMap(index), "+" & Item(ItemNum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                    SendEXP index
                End If
                Call SendAnimation(GetPlayerMap(index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                Call TakeInvItem(index, Player(index).Inv(invNum).Num, 0)
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_KEY
                TempPlayer(index).FreeAction = False

                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(ItemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(ItemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If

                Select Case GetPlayerDir(index)
                    Case DIR_UP

                        If GetPlayerY(index) > 0 Then
                            x = GetPlayerX(index)
                            y = GetPlayerY(index) - 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_DOWN

                        If GetPlayerY(index) < Map(GetPlayerMap(index)).MaxY Then
                            x = GetPlayerX(index)
                            y = GetPlayerY(index) + 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_LEFT

                        If GetPlayerX(index) > 0 Then
                            x = GetPlayerX(index) - 1
                            y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                    Case DIR_RIGHT

                        If GetPlayerX(index) < Map(GetPlayerMap(index)).MaxX Then
                            x = GetPlayerX(index) + 1
                            y = GetPlayerY(index)
                        Else
                            Exit Sub
                        End If

                End Select

                ' Check if a key exists
                If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY Then

                    ' Check if the key they are using matches the map key
                    If ItemNum = Map(GetPlayerMap(index)).Tile(x, y).Data1 Then
                        TempTile(GetPlayerMap(index)).DoorOpen(x, y) = 1
                        TempTile(GetPlayerMap(index)).DoorTimer = timeGetTime
                        SendMapKey index, x, y, 1
                        Call MapMsg(GetPlayerMap(index), "A door has been unlocked.", White)
                        
                        Call SendAnimation(GetPlayerMap(index), Item(ItemNum).Animation, x, y)

                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(index)).Tile(x, y).Data2 = 1 Then
                            Call TakeInvItem(index, ItemNum, 0)
                            Call PlayerMsg(index, "The key is destroyed in the lock.", Yellow)
                        End If
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
            Case ITEM_TYPE_SPELL
                ' stat requirements
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerRawStat(index, i) < Item(ItemNum).Stat_Req(i) Then
                        PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' level requirement
                If GetPlayerLevel(index) < Item(ItemNum).LevelReq Then
                    PlayerMsg index, "You do not meet the level requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' class requirement
                If Item(ItemNum).ClassReq > 0 Then
                    If Not GetPlayerClass(index) = Item(ItemNum).ClassReq Then
                        PlayerMsg index, "You do not meet the class requirement to use this item.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' access requirement
                If Not GetPlayerAccess(index) >= Item(ItemNum).AccessReq Then
                    PlayerMsg index, "You do not meet the access requirement to use this item.", BrightRed
                    Exit Sub
                End If
                
                ' Get the spell num
                n = Item(ItemNum).Data1

                If n > 0 Then

                    ' Make sure they are the right class
                    If Spell(n).ClassReq = GetPlayerClass(index) Or Spell(n).ClassReq = 0 Then
                        ' Make sure they are the right level
                        i = Spell(n).LevelReq

                        If i <= GetPlayerLevel(index) Then
                            i = FindOpenSpellSlot(index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(index, n) Then
                                    Call SetPlayerSpell(index, i, n)
                                    Call SendAnimation(GetPlayerMap(index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                                    Call TakeInvItem(index, ItemNum, 0)
                                    Call PlayerMsg(index, "You feel the rush of knowledge fill your mind. You can now use " & Trim$(Spell(n).name) & ".", BrightGreen)
                                    
                                    ' update 'em
                                    Call SendPlayerSpells(index)
                                Else
                                    Call PlayerMsg(index, "You already have knowledge of this skill.", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(index, "You cannot learn any more skills.", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(index, "You must be level " & i & " to learn this skill.", BrightRed)
                        End If

                    Else
                        Call PlayerMsg(index, "This spell can only be learned by " & CheckGrammar(GetClassName(Spell(n).ClassReq)) & ".", BrightRed)
                    End If
                End If
                
                ' send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum
        End Select
    End If
End Sub

Public Function TotalInventoryItems(ByVal index As Long) As Byte
Dim tempTotal As Byte
Dim i As Long

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(index, i) > 0 Then
            tempTotal = tempTotal + 1
        End If
    Next
    
    TotalInventoryItems = tempTotal
End Function
