Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub ServerLoop()
Dim i As Long, x As Long
Dim Tick As Long, TickCPS As Long, CPS As Long, FrameTime As Long
Dim tmr20 As Long, tmr500 As Long, tmr1000 As Long, tmr10000 As Long
Dim LastUpdateSavePlayers, LastUpdateMapSpawnItems(1 To MAX_MAPS) As Long, LastUpdatePlayerVitals As Long
Dim MapNum As Long, LastUpdateMapLogic(1 To MAX_MAPS) As Long

    ServerOnline = True

    Do While ServerOnline
        Tick = timeGetTime
        ElapsedTime = Tick - FrameTime
        FrameTime = Tick
        
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If Tick > tmr20 Then
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).spellBuffer.Spell > 0 Then
                        If timeGetTime > TempPlayer(i).spellBuffer.Timer + (Spell(Player(i).Spell(TempPlayer(i).spellBuffer.Spell)).CastTime * 1000) Then
                            CastSpell i, TempPlayer(i).spellBuffer.Spell, TempPlayer(i).spellBuffer.target, TempPlayer(i).spellBuffer.tType
                            TempPlayer(i).spellBuffer.Spell = 0
                            TempPlayer(i).spellBuffer.Timer = 0
                            TempPlayer(i).spellBuffer.target = 0
                            TempPlayer(i).spellBuffer.tType = 0
                        End If
                    End If
                    ' check if need to turn off stunned
                    If TempPlayer(i).StunDuration > 0 Then
                        If timeGetTime > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                            TempPlayer(i).StunDuration = 0
                            TempPlayer(i).StunTimer = 0
                            SendStunned i
                        End If
                    End If
                    ' check regen timer
                    If TempPlayer(i).stopRegen Then
                        If TempPlayer(i).stopRegenTimer + 5000 < timeGetTime Then
                            TempPlayer(i).stopRegen = False
                            TempPlayer(i).stopRegenTimer = 0
                        End If
                    End If
                    ' HoT and DoT logic
                    For x = 1 To MAX_DOTS
                        HandleDoT_Player i, x
                        HandleHoT_Player i, x
                    Next
                    tmr20 = timeGetTime + 20
                End If
            
                ' Checks to update player vitals every 5 seconds - Can be tweaked
                If Tick > LastUpdatePlayerVitals Then
                    UpdatePlayerVitals i
                    LastUpdatePlayerVitals = timeGetTime + 5000
                End If
                
                ' Checks to save players every 5 minutes - Can be tweaked
                If Tick > LastUpdateSavePlayers Then
                    UpdateSavePlayers i
                    LastUpdateSavePlayers = timeGetTime + 300000
                End If
            End If
        Next

        ' Check for disconnections every half second
        If Tick > tmr500 Then
            For i = 1 To MAX_PLAYERS
                If Player(i).Map > 0 Then
                    If Map(Player(i).Map).Moral = MAP_MORAL_NONE Then
                        If frmServer.Socket(i).State > sckConnected And TempPlayer(i).logout = False Then
                            TempPlayer(i).logout = True
                            TempPlayer(i).logoutTmr = timeGetTime + 10000
                        End If
                    Else
                        If frmServer.Socket(i).State > sckConnected Then
                        CloseSocket (i)
                        
                    End If
                End If
                End If
            Next
            
            tmr500 = timeGetTime + 500
        End If
        
        ' 10 second logout
            For i = 1 To MAX_PLAYERS
            If TempPlayer(i).logout Then
            If Tick > TempPlayer(i).logoutTmr Then
                TempPlayer(i).logout = False
                TempPlayer(i).logoutTmr = 0
            CloseSocket (i)
            End If
        End If
    Next

        If Tick > tmr1000 Then
            For i = 1 To Player_HighIndex
                If TempPlayer(i).FreeAction = False Then TempPlayer(i).FreeAction = True
            Next
            
            If isShuttingDown Then
                Call HandleShutdown
            End If
            tmr1000 = timeGetTime + 1000
        End If
        
        For MapNum = 1 To MAX_MAPS
            ' Checks to spawn map items every 5 minutes - Can be tweaked
            If Tick > LastUpdateMapSpawnItems(MapNum) Then
                UpdateMapSpawnItems MapNum
                LastUpdateMapSpawnItems(MapNum) = timeGetTime + 300000
            End If
            
            ' update map logic
            If Tick > LastUpdateMapLogic(MapNum) Then
                UpdateMapLogic MapNum
                LastUpdateMapLogic(MapNum) = timeGetTime + 500
            End If
        Next

        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
             
        ' Set the server CPS
        frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
    Loop
End Sub

Private Sub UpdateMapSpawnItems(ByVal y As Long)
Dim x As Long

    ' Clear out unnecessary junk
    For x = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(x, y)
    Next

    ' Spawn the items
    Call SpawnMapItems(y)
    Call SendMapItemsToAll(y)
End Sub

Private Sub UpdateMapLogic(MapNum As Long)
Dim i As Long, x As Long, n As Long, x1 As Long, y1 As Long
Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, NPCNum As Long
Dim target As Long, targetType As Byte, DidWalk As Boolean, Buffer As clsBuffer, Resource_index As Long
Dim TargetX As Long, TargetY As Long, Target_Verify As Boolean
Dim ReplacedAttackSay As String

    ' items appearing to everyone
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(MapNum, i).Num > 0 Then
            If MapItem(MapNum, i).playerName <> vbNullString Then
                ' make item public?
                If MapItem(MapNum, i).playerTimer < timeGetTime Then
                    ' make it public
                    MapItem(MapNum, i).playerName = vbNullString
                    MapItem(MapNum, i).playerTimer = 0
                    ' send updates to everyone
                    SendMapItemsToAll MapNum
                End If
                ' despawn item?
                If MapItem(MapNum, i).canDespawn Then
                    If MapItem(MapNum, i).despawnTimer < timeGetTime Then
                        ' despawn it
                        ClearMapItem i, MapNum
                        ' send updates to everyone
                        SendMapItemsToAll MapNum
                    End If
                End If
            End If
        End If
    Next
    
    '  Close the doors
    If TickCount > TempTile(MapNum).DoorTimer + 5000 Then
        For x1 = 0 To Map(MapNum).MaxX
            For y1 = 0 To Map(MapNum).MaxY
                If Map(MapNum).Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(MapNum).DoorOpen(x1, y1) = 1 Then
                    TempTile(MapNum).DoorOpen(x1, y1) = 0
                    SendMapKeyToMap MapNum, x1, y1, 0
                End If
            Next
        Next
    End If
    
    ' check for DoTs + hots
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(MapNum).Npc(i).Num > 0 Then
            For x = 1 To MAX_DOTS
                HandleDoT_Npc MapNum, i, x
                HandleHoT_Npc MapNum, i, x
            Next
        End If
    Next

    ' Respawning Resources
    If ResourceCache(MapNum).Resource_Count > 0 Then
        For i = 0 To ResourceCache(MapNum).Resource_Count
            Resource_index = Map(MapNum).Tile(ResourceCache(MapNum).ResourceData(i).x, ResourceCache(MapNum).ResourceData(i).y).Data1

            If Resource_index > 0 Then
                If ResourceCache(MapNum).ResourceData(i).ResourceState = 1 Or ResourceCache(MapNum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                    If ResourceCache(MapNum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < timeGetTime Then
                        ResourceCache(MapNum).ResourceData(i).ResourceTimer = timeGetTime
                        ResourceCache(MapNum).ResourceData(i).ResourceState = 0 ' normal
                        ' re-set health to resource root
                        ResourceCache(MapNum).ResourceData(i).cur_health = Resource(Resource_index).health
                        SendResourceCacheToMap MapNum, i
                    End If
                End If
            End If
        Next
    End If

    TickCount = timeGetTime
    
    For x = 1 To MAX_MAP_NPCS
        NPCNum = MapNpc(MapNum).Npc(x).Num

        ' /////////////////////////////////////////
        ' // This is used for ATTACKING ON SIGHT //
        ' /////////////////////////////////////////
        ' Make sure theres a npc with the map
        If Map(MapNum).Npc(x) > 0 And MapNpc(MapNum).Npc(x).Num > 0 Then

            ' If the npc is a attack on sight, search for a player on the map
            If Npc(NPCNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(NPCNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
            
                ' make sure it's not stunned
                If Not MapNpc(MapNum).Npc(x).StunDuration > 0 Then

                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = MapNum And MapNpc(MapNum).Npc(x).target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                n = Npc(NPCNum).Range
                                DistanceX = MapNpc(MapNum).Npc(x).x - GetPlayerX(i)
                                DistanceY = MapNpc(MapNum).Npc(x).y - GetPlayerY(i)

                                ' Make sure we get a positive value
                                If DistanceX < 0 Then DistanceX = DistanceX * -1
                                If DistanceY < 0 Then DistanceY = DistanceY * -1

                                ' Are they in range? if so GET'M!
                                If DistanceX <= n And DistanceY <= n Then
                                    If Npc(NPCNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Then
                                        If Len(Trim$(Npc(NPCNum).AttackSay)) > 0 Then
                                        
                                            ' See if we have any replacement strings
                                            ReplacedAttackSay = Trim$(Replace$(Npc(NPCNum).AttackSay, "<playername>", GetPlayerName(i)))
                                            ReplacedAttackSay = Replace$(ReplacedAttackSay, "<class>", Trim$(Class(Player(i).Class).name))
                                            
                                            ' Output it
                                            Call PlayerMsg(i, Trim$(Npc(NPCNum).name) & " says: " & ReplacedAttackSay, SayColor)
                                        End If
                                        MapNpc(MapNum).Npc(x).targetType = 1 ' player
                                        MapNpc(MapNum).Npc(x).target = i
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End If
        
        Target_Verify = False

        ' /////////////////////////////////////////////
        ' // This is used for NPC walking/targetting //
        ' /////////////////////////////////////////////
        ' Make sure theres a npc with the map
        If Map(MapNum).Npc(x) > 0 And MapNpc(MapNum).Npc(x).Num > 0 Then
            If MapNpc(MapNum).Npc(x).StunDuration > 0 Then
                ' check if we can unstun them
                If timeGetTime > MapNpc(MapNum).Npc(x).StunTimer + (MapNpc(MapNum).Npc(x).StunDuration * 1000) Then
                    MapNpc(MapNum).Npc(x).StunDuration = 0
                    MapNpc(MapNum).Npc(x).StunTimer = 0
                End If
            Else
                ' check if they're chatting
                If MapNpc(MapNum).Npc(x).InChat > 0 Then
                    If Not TempPlayer(MapNpc(MapNum).Npc(x).InChat).InChat = NPCNum Then
                        MapNpc(MapNum).Npc(x).InChat = 0
                    End If
                Else
                    target = MapNpc(MapNum).Npc(x).target
                    targetType = MapNpc(MapNum).Npc(x).targetType
    
                    ' Check to see if its time for the npc to walk
                    If Npc(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    
                        If targetType = 1 Then ' player
    
                            ' Check to see if we are following a player or not
                            If target > 0 Then
    
                                ' Check if the player is even playing, if so follow'm
                                If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                                    DidWalk = False
                                    Target_Verify = True
                                    TargetY = GetPlayerY(target)
                                    TargetX = GetPlayerX(target)
                                Else
                                    MapNpc(MapNum).Npc(x).targetType = 0 ' clear
                                    MapNpc(MapNum).Npc(x).target = 0
                                End If
                            End If
                        
                        ElseIf targetType = 2 Then 'npc
                            
                            If target > 0 Then
                                
                                If MapNpc(MapNum).Npc(target).Num > 0 Then
                                    DidWalk = False
                                    Target_Verify = True
                                    TargetY = MapNpc(MapNum).Npc(target).y
                                    TargetX = MapNpc(MapNum).Npc(target).x
                                Else
                                    MapNpc(MapNum).Npc(x).targetType = 0 ' clear
                                    MapNpc(MapNum).Npc(x).target = 0
                                End If
                            End If
                        End If
                        
                        If Target_Verify Then
                            
                            i = Int(Rnd * 5)
    
                            ' Lets move the npc
                            Select Case i
                                Case 0
    
                                    ' Up
                                    If MapNpc(MapNum).Npc(x).y > TargetY And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_UP) Then
                                            Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                    ' Down
                                    If MapNpc(MapNum).Npc(x).y < TargetY And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_DOWN) Then
                                            Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                    ' Left
                                    If MapNpc(MapNum).Npc(x).x > TargetX And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_LEFT) Then
                                            Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                    ' Right
                                    If MapNpc(MapNum).Npc(x).x < TargetX And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_RIGHT) Then
                                            Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                Case 1
    
                                    ' Right
                                    If MapNpc(MapNum).Npc(x).x < TargetX And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_RIGHT) Then
                                            Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                    ' Left
                                    If MapNpc(MapNum).Npc(x).x > TargetX And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_LEFT) Then
                                            Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                    ' Down
                                    If MapNpc(MapNum).Npc(x).y < TargetY And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_DOWN) Then
                                            Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                    ' Up
                                    If MapNpc(MapNum).Npc(x).y > TargetY And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_UP) Then
                                            Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                Case 2
    
                                    ' Down
                                    If MapNpc(MapNum).Npc(x).y < TargetY And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_DOWN) Then
                                            Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                    ' Up
                                    If MapNpc(MapNum).Npc(x).y > TargetY And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_UP) Then
                                            Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                    ' Right
                                    If MapNpc(MapNum).Npc(x).x < TargetX And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_RIGHT) Then
                                            Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                    ' Left
                                    If MapNpc(MapNum).Npc(x).x > TargetX And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_LEFT) Then
                                            Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                Case 3
    
                                    ' Left
                                    If MapNpc(MapNum).Npc(x).x > TargetX And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_LEFT) Then
                                            Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                    ' Right
                                    If MapNpc(MapNum).Npc(x).x < TargetX And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_RIGHT) Then
                                            Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                    ' Up
                                    If MapNpc(MapNum).Npc(x).y > TargetY And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_UP) Then
                                            Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                                    ' Down
                                    If MapNpc(MapNum).Npc(x).y < TargetY And Not DidWalk Then
                                        If CanNpcMove(MapNum, x, DIR_DOWN) Then
                                            Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING)
                                            DidWalk = True
                                        End If
                                    End If
    
                            End Select
    
                            ' Check if we can't move and if Target is behind something and if we can just switch dirs
                            If Not DidWalk Then
                                If MapNpc(MapNum).Npc(x).x - 1 = TargetX And MapNpc(MapNum).Npc(x).y = TargetY Then
                                    If MapNpc(MapNum).Npc(x).Dir <> DIR_LEFT Then
                                        Call NpcDir(MapNum, x, DIR_LEFT)
                                    End If
    
                                    DidWalk = True
                                End If
    
                                If MapNpc(MapNum).Npc(x).x + 1 = TargetX And MapNpc(MapNum).Npc(x).y = TargetY Then
                                    If MapNpc(MapNum).Npc(x).Dir <> DIR_RIGHT Then
                                        Call NpcDir(MapNum, x, DIR_RIGHT)
                                    End If
    
                                    DidWalk = True
                                End If
    
                                If MapNpc(MapNum).Npc(x).x = TargetX And MapNpc(MapNum).Npc(x).y - 1 = TargetY Then
                                    If MapNpc(MapNum).Npc(x).Dir <> DIR_UP Then
                                        Call NpcDir(MapNum, x, DIR_UP)
                                    End If
    
                                    DidWalk = True
                                End If
    
                                If MapNpc(MapNum).Npc(x).x = TargetX And MapNpc(MapNum).Npc(x).y + 1 = TargetY Then
                                    If MapNpc(MapNum).Npc(x).Dir <> DIR_DOWN Then
                                        Call NpcDir(MapNum, x, DIR_DOWN)
                                    End If
    
                                    DidWalk = True
                                End If
    
                                ' We could not move so Target must be behind something, walk randomly.
                                If Not DidWalk Then
                                    i = Int(Rnd * 2)
    
                                    If i = 1 Then
                                        i = Int(Rnd * 4)
    
                                        If CanNpcMove(MapNum, x, i) Then
                                            Call NpcMove(MapNum, x, i, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            End If
    
                        Else
                            i = Int(Rnd * 4)
    
                            If i = 1 Then
                                i = Int(Rnd * 4)
    
                                If CanNpcMove(MapNum, x, i) Then
                                    Call NpcMove(MapNum, x, i, MOVING_WALKING)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        ' /////////////////////////////////////////////
        ' // This is used for npcs to attack targets //
        ' /////////////////////////////////////////////
        ' Make sure theres a npc with the map
        If Map(MapNum).Npc(x) > 0 And MapNpc(MapNum).Npc(x).Num > 0 Then
            target = MapNpc(MapNum).Npc(x).target
            targetType = MapNpc(MapNum).Npc(x).targetType

            ' Check if the npc can attack the targeted player player
            If target > 0 Then
            
                If targetType = 1 Then ' player

                    ' Is the target playing and on the same map?
                    If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                        TryNpcAttackPlayer x, target
                    Else
                        ' Player left map or game, set target to 0
                        MapNpc(MapNum).Npc(x).target = 0
                        MapNpc(MapNum).Npc(x).targetType = 0 ' clear
                    End If
                Else
                    ' lol no npc combat :(
                End If
            End If
        End If

        ' ////////////////////////////////////////////
        ' // This is used for regenerating NPC's HP //
        ' ////////////////////////////////////////////
        ' Check to see if we want to regen some of the npc's hp
        If Not MapNpc(MapNum).Npc(x).stopRegen Then
            If MapNpc(MapNum).Npc(x).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                If MapNpc(MapNum).Npc(x).Vital(Vitals.HP) > 0 Then
                    MapNpc(MapNum).Npc(x).Vital(Vitals.HP) = MapNpc(MapNum).Npc(x).Vital(Vitals.HP) + GetNpcVitalRegen(NPCNum, Vitals.HP)

                    ' Check if they have more then they should and if so just set it to max
                    If MapNpc(MapNum).Npc(x).Vital(Vitals.HP) > GetNpcMaxVital(NPCNum, Vitals.HP) Then
                        MapNpc(MapNum).Npc(x).Vital(Vitals.HP) = GetNpcMaxVital(NPCNum, Vitals.HP)
                    End If
                End If
            End If
        End If

        ' ////////////////////////////////////////////////////////
        ' // This is used for checking if an NPC is dead or not //
        ' ////////////////////////////////////////////////////////
        ' Check if the npc is dead or not
        'If MapNpc(y, x).Num > 0 Then
        '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).STR > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
        '        MapNpc(y, x).Num = 0
        '        MapNpc(y, x).SpawnWait = TickCount
        '   End If
        'End If
        
        ' //////////////////////////////////////
        ' // This is used for spawning an NPC //
        ' //////////////////////////////////////
        ' Check if we are supposed to spawn an npc or not
        If MapNpc(MapNum).Npc(x).Num = 0 And Map(MapNum).Npc(x) > 0 Then
            If TickCount > MapNpc(MapNum).Npc(x).SpawnWait + (Npc(Map(MapNum).Npc(x)).SpawnSecs * 1000) Then
                Call SpawnNpc(x, MapNum)
            End If
        End If

    Next
    DoEvents

    ' Make sure we reset the timer for npc hp regeneration
    If timeGetTime > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = timeGetTime
    End If
End Sub

Private Sub UpdatePlayerVitals(ByVal index As Long)

    If Not TempPlayer(index).stopRegen Then
        If GetPlayerVital(index, Vitals.HP) <> GetPlayerMaxVital(index, Vitals.HP) Then
            Call SetPlayerVital(index, Vitals.HP, GetPlayerVital(index, Vitals.HP) + GetPlayerVitalRegen(index, Vitals.HP))
            Call SendVital(index, Vitals.HP)
            
            ' send vitals to party if in one
            If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
        End If
    
        If GetPlayerVital(index, Vitals.MP) <> GetPlayerMaxVital(index, Vitals.MP) Then
            Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) + GetPlayerVitalRegen(index, Vitals.MP))
            Call SendVital(index, Vitals.MP)
            
            ' send vitals to party if in one
            If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
        End If
    End If
End Sub

Private Sub UpdateSavePlayers(ByVal index As Long)
Dim i As Long

    ' Prevent subscript out range
    If Not IsPlaying(index) Then Exit Sub
    
    ' Save player
    Call TextAdd("Saving all online players...")
    Call SavePlayer(index)
    Call SaveBank(index)
End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", BrightRed)
        Call DestroyServer
    End If

End Sub
