Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    Select Case Vital
        Case HP
            Select Case GetPlayerClass(index)
                Case 1 ' Warrior
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 15 + 150
                Case 2 ' Magician
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 5 + 65
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 15 + 150
            End Select
        Case MP
            Select Case GetPlayerClass(index)
                Case 1 ' Warrior
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 5 + 25
                Case 2 ' Magician
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 30 + 85
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 5 + 25
            End Select
    End Select
End Function

Function GetPlayerVitalRegen(ByVal index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (GetPlayerStat(index, Stats.Willpower) * 0.8) + 6
        Case MP
            i = (GetPlayerStat(index, Stats.Willpower) / 4) + 12.5
    End Select

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i
End Function

Function GetPlayerDamage(ByVal index As Long) As Long
    Dim weaponNum As Long
    
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(index, Weapon)
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(index, Strength) * Item(weaponNum).Data2 + (GetPlayerLevel(index) / 5)
    Else
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(index, Strength) + (GetPlayerLevel(index) / 5)
    End If

End Function

Public Function GetPlayerDef(ByVal index As Long) As Long
Dim ItemNum As Long
Dim Def As Long
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    
    ' Set it to 1 so we don't mess up our calculations
    Def = 1
    
    If GetPlayerEquipment(index, Helmet) > 0 Then
        ItemNum = GetPlayerEquipment(index, Helmet)
        Def = Def + Item(ItemNum).Data2
    End If
    
    If GetPlayerEquipment(index, Armor) > 0 Then
        ItemNum = GetPlayerEquipment(index, Armor)
        Def = Def + Item(ItemNum).Data2
    End If
    
    If GetPlayerEquipment(index, Shield) > 0 Then
        ItemNum = GetPlayerEquipment(index, Shield)
        Def = Def + Item(ItemNum).Data2
    End If
    
    GetPlayerDef = 0.085 * GetPlayerStat(index, Endurance) * Def + (GetPlayerLevel(index) / 5)
End Function

Function GetNpcMaxVital(ByVal NPCNum As Long, ByVal Vital As Vitals) As Long
    Dim x As Long

    ' Prevent subscript out of range
    If NPCNum <= 0 Or NPCNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            GetNpcMaxVital = Npc(NPCNum).HP
        Case MP
            GetNpcMaxVital = 30 + (Npc(NPCNum).Stat(Intelligence) * 10) + 2
    End Select

End Function

Function GetNpcVitalRegen(ByVal NPCNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range
    If NPCNum <= 0 Or NPCNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (Npc(NPCNum).Stat(Stats.Willpower) * 0.8) + 6
        Case MP
            i = (Npc(NPCNum).Stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetNpcVitalRegen = i

End Function

Function GetNpcDamage(ByVal NPCNum As Long) As Long
    GetNpcDamage = 0.085 * 5 * Npc(NPCNum).Stat(Stats.Strength) * Npc(NPCNum).Damage + (Npc(NPCNum).Level / 5)
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerBlock(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long
Dim shieldNum As Long

    CanPlayerBlock = False

    rate = 0

    If GetPlayerEquipment(index, Shield) > 0 Then
        shieldNum = GetPlayerEquipment(index, Shield)
        rate = Item(shieldNum).Data2 / 43.3
        rndNum = Rand(1, 100)
        If rndNum <= rate Then
            CanPlayerBlock = True
        Else
            CanPlayerBlock = False
        End If
    Else
        rate = 0.5
        rndNum = Rand(1, 100)
        If rndNum <= rate Then
            CanPlayerBlock = True
        Else
            CanPlayerBlock = False
        End If
    End If
End Function

Public Function CanPlayerCrit(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerCrit = False

    rate = GetPlayerStat(index, Agility) / 52.08
    rndNum = Rand(1, 100)
    If rndNum <= rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerDodge = False

    rate = GetPlayerStat(index, Agility) / 83.3
    rndNum = Rand(1, 100)
    If rndNum <= rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal index As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanPlayerParry = False

    rate = GetPlayerStat(index, Strength) * 0.25
    rndNum = Rand(1, 100)
    If rndNum <= rate Then
        CanPlayerParry = True
    End If
End Function

Public Function CanNpcBlock(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcBlock = False

    rate = 0
    ' TODO : make it based on shield lol
End Function

Public Function CanNpcCrit(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcCrit = False

    rate = Npc(NPCNum).Stat(Stats.Agility) / 52.08
    rndNum = Rand(1, 100)
    If rndNum <= rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodge(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcDodge = False

    rate = Npc(NPCNum).Stat(Stats.Agility) / 83.3
    rndNum = Rand(1, 100)
    If rndNum <= rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal NPCNum As Long) As Boolean
Dim rate As Long
Dim rndNum As Long

    CanNpcParry = False

    rate = Npc(NPCNum).Stat(Stats.Strength) * 0.25
    rndNum = Rand(1, 100)
    If rndNum <= rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal index As Long, ByVal mapNpcNum As Long)
Dim blockAmount As Long
Dim NPCNum As Long
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(index, mapNpcNum) Then
    
        MapNum = GetPlayerMap(index)
        NPCNum = MapNpc(MapNum).Npc(mapNpcNum).Num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(NPCNum) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (MapNpc(MapNum).Npc(mapNpcNum).x * 32), (MapNpc(MapNum).Npc(mapNpcNum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(NPCNum) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (MapNpc(MapNum).Npc(mapNpcNum).x * 32), (MapNpc(MapNum).Npc(mapNpcNum).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(mapNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - Rand(1, (Npc(NPCNum).Stat(Stats.Agility) * 2))
        ' randomise from 1 to max hit
        Damage = Rand(1, Damage)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(index) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
        
        ' make sure it isn't less than 0
        If Damage < 0 Then Damage = 0
        
        If Damage > 0 Then
            Call PlayerAttackNpc(index, mapNpcNum, Damage)
        Else
            Call PlayerAttackNpc(index, mapNpcNum, Damage)
            Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal attacker As Long, ByVal mapNpcNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
Dim MapNum As Long, NPCNum As Long
Dim NpcX As Long, NpcY As Long
Dim attackspeed As Long
Dim ReplacedAttackSay As String

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(attacker)).Npc(mapNpcNum).Num <= 0 Then
        Exit Function
    End If
    
    NPCNum = MapNpc(GetPlayerMap(attacker)).Npc(mapNpcNum).Num
    MapNum = GetPlayerMap(attacker)
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        If Npc(MapNpc(MapNum).Npc(mapNpcNum).Num).Behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
            If Npc(MapNpc(MapNum).Npc(mapNpcNum).Num).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                Exit Function
            End If
        End If
    End If

    ' Make sure they are on the same map
    If IsPlaying(attacker) Then
    
        ' exit out early
        If IsSpell Then
             If NPCNum > 0 Then
                If Npc(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If NPCNum > 0 And timeGetTime > TempPlayer(attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(attacker)
                Case DIR_UP
                    NpcX = MapNpc(MapNum).Npc(mapNpcNum).x
                    NpcY = MapNpc(MapNum).Npc(mapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(MapNum).Npc(mapNpcNum).x
                    NpcY = MapNpc(MapNum).Npc(mapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(MapNum).Npc(mapNpcNum).x + 1
                    NpcY = MapNpc(MapNum).Npc(mapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(MapNum).Npc(mapNpcNum).x - 1
                    NpcY = MapNpc(MapNum).Npc(mapNpcNum).y
            End Select

            If NpcX = GetPlayerX(attacker) Then
                If NpcY = GetPlayerY(attacker) Then
                    If Npc(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And Npc(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        CanPlayerAttackNpc = True
                        
                        ' set them as our target
                        TempPlayer(attacker).targetType = TARGET_TYPE_NPC
                        TempPlayer(attacker).target = mapNpcNum
                        SendTarget attacker
                    Else
                        ' Output it
                        If Npc(NPCNum).Conv > 0 Then
                            If TempPlayer(attacker).InChat <> NPCNum Then
                                TempPlayer(attacker).InChat = NPCNum
                                MapNpc(MapNum).Npc(mapNpcNum).InChat = attacker
                                Call SendStartConv(attacker, Npc(NPCNum).Conv, NPCNum)
                            End If
                        ElseIf Len(Trim$(Npc(NPCNum).AttackSay)) > 0 Then
                            ' See if we have any replacement strings
                            ReplacedAttackSay = Trim$(Replace$(Npc(NPCNum).AttackSay, "<playername>", GetPlayerName(attacker)))
                            ReplacedAttackSay = Replace$(ReplacedAttackSay, "<class>", Trim$(Class(Player(attacker).Class).name))
                            Call PlayerMsg(attacker, Trim$(Npc(NPCNum).name) & " says: " & ReplacedAttackSay, SayColor)
                        End If
                        
                         ' Reset attack timer
                        TempPlayer(attacker).AttackTimer = timeGetTime
                    End If
                End If
            End If
        End If
    End If
End Function

Public Sub PlayerAttackNpc(ByVal attacker As Long, ByVal mapNpcNum As Long, ByVal Damage As Long, Optional ByVal spellNum As Long, Optional ByVal overTime As Boolean = False)
Dim name As String
Dim exp As Long
Dim n As Long, i As Long, j As Long
Dim STR As Long, Def As Long
Dim MapNum As Long, NPCNum As Long
Dim DecimalChance As Double
Dim Buffer As clsBuffer
Dim NPCName As String, Progress As Long, OutOf As String

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(attacker)
    NPCNum = MapNpc(MapNum).Npc(mapNpcNum).Num
    name = Trim$(Npc(NPCNum).name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = timeGetTime

    If Damage >= MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.HP) Then
    
        SendActionMsg GetPlayerMap(attacker), "-" & MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(MapNum).Npc(mapNpcNum).x * 32), (MapNpc(MapNum).Npc(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y
        
        ' send the sound
        If spellNum > 0 Then SendMapSound attacker, MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y, SoundEntity.seSpell, spellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If spellNum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y)
            End If
        End If

        ' Calculate exp to give attacker
        exp = Npc(NPCNum).exp

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 1
        End If

        ' in party?
        If TempPlayer(attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(attacker).inParty, exp, attacker, GetPlayerMap(attacker)
        Else
            ' no party - keep exp for self
            GivePlayerEXP attacker, exp
        End If
        
        ' Drop the goods if they get it
        For n = 1 To MAX_NPC_DROPS
            If Npc(NPCNum).DropItem(n) = 0 Then Exit For
            DecimalChance = Npc(NPCNum).DropChance(n) / 100
            
            If Rnd <= DecimalChance Then
                Call SpawnItem(Npc(NPCNum).DropItem(n), Npc(NPCNum).DropItemValue(n), MapNum, _
                MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y)
            End If
        Next
        
        ' quests
        For i = 1 To MAX_QUESTS
            For j = 1 To MAX_QUEST_OBJECTIVES
                ' kill quest
                If Player(attacker).Quest(i).QuestStatus = 1 And Quest(i).Objective(j).ObjectiveType = 2 Then
                    ' make sure it's the right npc
                    If NPCNum = Quest(i).Objective(j).DataIndex Then
                        ' update the requirement
                        Player(attacker).Quest(i).DataAmountLeft(j) = Player(attacker).Quest(i).DataAmountLeft(j) - 1
                        
                        ' update them
                        NPCName = Trim$(Npc(Quest(i).Objective(j).DataIndex).name)
                        Progress = Quest(i).Objective(j).DataAmount - Player(attacker).Quest(i).DataAmountLeft(j)
                        OutOf = CStr(Quest(i).Objective(j).DataAmount)
                        
                        If Progress <= Quest(i).Objective(j).DataAmount Then
                            PlayerMsg attacker, Trim$(Quest(i).name) & ": " & NPCName & " slain: " & Progress & "/" & OutOf, Yellow
                        End If
                    End If
                End If
            Next
        Next
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).Npc(mapNpcNum).Num = 0
        MapNpc(MapNum).Npc(mapNpcNum).SpawnWait = timeGetTime
        MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(MapNum).Npc(mapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(MapNum).Npc(mapNpcNum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong mapNpcNum
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
        
        ' Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = MapNum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = mapNpcNum Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.HP) = MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).Npc(mapNpcNum).x * 32), (MapNpc(MapNum).Npc(mapNpcNum).y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y
        
        ' send the sound
        If spellNum > 0 Then SendMapSound attacker, MapNpc(MapNum).Npc(mapNpcNum).x, MapNpc(MapNum).Npc(mapNpcNum).y, SoundEntity.seSpell, spellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If spellNum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, mapNpcNum)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(MapNum).Npc(mapNpcNum).targetType = 1 ' player
        MapNpc(MapNum).Npc(mapNpcNum).target = attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum).Npc(mapNpcNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(i).Num = MapNpc(MapNum).Npc(mapNpcNum).Num Then
                    MapNpc(MapNum).Npc(i).target = attacker
                    MapNpc(MapNum).Npc(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(MapNum).Npc(mapNpcNum).stopRegen = True
        MapNpc(MapNum).Npc(mapNpcNum).stopRegenTimer = timeGetTime
        
        ' if stunning spell, stun the npc
        If spellNum > 0 Then
            If Spell(spellNum).StunDuration > 0 Then StunNPC mapNpcNum, MapNum, spellNum
            ' DoT
            If Spell(spellNum).Duration > 0 Then
                AddDoT_Npc MapNum, mapNpcNum, spellNum, attacker
            End If
        End If
        
        SendMapNpcVitals MapNum, mapNpcNum
    End If

    If spellNum = 0 Then
        ' Reset attack timer
        TempPlayer(attacker).AttackTimer = timeGetTime
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal mapNpcNum As Long, ByVal index As Long)
Dim MapNum As Long, NPCNum As Long, blockAmount As Long, Damage As Long

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(mapNpcNum, index) Then
        MapNum = GetPlayerMap(index)
        NPCNum = MapNpc(MapNum).Npc(mapNpcNum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            Exit Sub
        End If
        If CanPlayerParry(index) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(NPCNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(index)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - Rand(1, (GetPlayerStat(index, Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = Rand(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(NPCNum) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (MapNpc(MapNum).Npc(mapNpcNum).x * 32), (MapNpc(MapNum).Npc(mapNpcNum).y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(mapNpcNum, index, Damage)
        Else
            SendActionMsg MapNum, "Block!", Cyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
    End If
End Sub

Public Function CanNpcAttackPlayer(ByVal mapNpcNum As Long, ByVal index As Long) As Boolean
Dim MapNum As Long
Dim NPCNum As Long
Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index)).Npc(mapNpcNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(index)
    NPCNum = MapNpc(MapNum).Npc(mapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < MapNpc(MapNum).Npc(mapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = 1 Then
        Exit Function
    End If
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SNpcAttack
    Buffer.WriteLong mapNpcNum
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing

    MapNpc(MapNum).Npc(mapNpcNum).AttackTimer = timeGetTime

    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If NPCNum > 0 Then

            ' Check if at same coordinates
            If (GetPlayerY(index) + 1 = MapNpc(MapNum).Npc(mapNpcNum).y) And (GetPlayerX(index) = MapNpc(MapNum).Npc(mapNpcNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(index) - 1 = MapNpc(MapNum).Npc(mapNpcNum).y) And (GetPlayerX(index) = MapNpc(MapNum).Npc(mapNpcNum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) = MapNpc(MapNum).Npc(mapNpcNum).y) And (GetPlayerX(index) + 1 = MapNpc(MapNum).Npc(mapNpcNum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = MapNpc(MapNum).Npc(mapNpcNum).y) And (GetPlayerX(index) - 1 = MapNpc(MapNum).Npc(mapNpcNum).x) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Public Sub NpcAttackPlayer(ByVal mapNpcNum As Long, ByVal victim As Long, ByVal Damage As Long)
    Dim name As String
    Dim exp As Long
    Dim MapNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or IsPlaying(victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(victim)).Npc(mapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(victim)
    name = Trim$(Npc(MapNpc(MapNum).Npc(mapNpcNum).Num).name)

    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(MapNum).Npc(mapNpcNum).stopRegen = True
    MapNpc(MapNum).Npc(mapNpcNum).stopRegenTimer = timeGetTime

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(mapNpcNum).Num
        
        ' kill player
        KillPlayer victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & name, BrightRed)

        ' Set NPC target to 0
        MapNpc(MapNum).Npc(mapNpcNum).target = 0
        MapNpc(MapNum).Npc(mapNpcNum).targetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        Call SendAnimation(MapNum, Npc(MapNpc(GetPlayerMap(victim)).Npc(mapNpcNum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, victim)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(mapNpcNum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = timeGetTime
    End If

End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long)
Dim blockAmount As Long
Dim NPCNum As Long
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(attacker, victim) Then
    
        MapNum = GetPlayerMap(attacker)
    
        ' check if they can avoid the attack
        If CanPlayerDodge(victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If
        
        If CanPlayerParry(victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - Rand(1, (GetPlayerStat(victim, Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = Rand(1, Damage)
        
        ' * 1.5 if can crit
        If CanPlayerCrit(attacker) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32)
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(attacker, victim, Damage)
        Else
            Call PlayerMsg(attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean

    If Not IsSpell Then
        ' Check attack timer
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            If timeGetTime < TempPlayer(attacker).AttackTimer + Item(GetPlayerEquipment(attacker, Weapon)).Speed Then Exit Function
        Else
            If timeGetTime < TempPlayer(attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = 1 Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP
    
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Then
        Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
        Exit Function
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function

    ' Make sure attacker is high enough level
    If GetPlayerLevel(attacker) < MIN_PVP_LEVEL Then
        Call PlayerMsg(attacker, "You are below level " & MIN_PVP_LEVEL & ", you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < MIN_PVP_LEVEL Then
        Call PlayerMsg(attacker, GetPlayerName(victim) & " is below level " & MIN_PVP_LEVEL & ", you cannot attack this player yet!", BrightRed)
        Exit Function
    End If
    
    ' Set them as our target
    TempPlayer(attacker).targetType = TARGET_TYPE_PLAYER
    TempPlayer(attacker).target = victim
    SendTarget attacker
    
    CanPlayerAttackPlayer = True
End Function

Public Sub PlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal spellNum As Long = 0)
Dim exp As Long
Dim n As Long
Dim i As Long
Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(attacker) = False Or IsPlaying(victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = timeGetTime

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        If spellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, spellNum
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(attacker), BrightRed)
        ' Calculate exp to give attacker
        exp = (GetPlayerExp(victim) \ 10)

        ' Make sure we dont get less then 0
        If exp < 0 Then
            exp = 0
        End If

        If exp = 0 Then
            Call PlayerMsg(victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(victim, GetPlayerExp(victim) - exp)
            SendEXP victim
            Call PlayerMsg(victim, "You lost " & exp & " exp.", BrightRed)
            
            ' check if we're in a party
            If TempPlayer(attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(attacker).inParty, exp, attacker, GetPlayerMap(attacker)
            Else
                ' not in party, get exp for self
                GivePlayerEXP attacker, exp
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(attacker) Then
                    If TempPlayer(i).target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        Call OnDeath(victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        If spellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, spellNum
        
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = timeGetTime
        
        ' if a stunning spell, stun the player
        If spellNum > 0 Then
            If Spell(spellNum).StunDuration > 0 Then StunPlayer victim, spellNum
            ' DoT
            If Spell(spellNum).Duration > 0 Then
                AddDoT_Player victim, spellNum, attacker
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(attacker).AttackTimer = timeGetTime
End Sub

' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal index As Long, ByVal spellslot As Long)
    Dim spellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    Dim targetType As Byte
    Dim target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    spellNum = GetPlayerSpell(index, spellslot)
    MapNum = GetPlayerMap(index)
    
    If spellNum <= 0 Or spellNum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(index, spellNum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > timeGetTime Then
        PlayerMsg index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = Spell(spellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(spellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(spellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(spellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(ClassReq).name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    targetType = TempPlayer(index).targetType
    target = TempPlayer(index).target
    Range = Spell(spellNum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not target > 0 Then
                PlayerMsg index, "You do not have a target.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(target), GetPlayerY(target)) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                Else
                    ' go through spell types
                    If Spell(spellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), MapNpc(MapNum).Npc(target).x, MapNpc(MapNum).Npc(target).y) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(spellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation MapNum, Spell(spellNum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, index
        TempPlayer(index).spellBuffer.Spell = spellslot
        TempPlayer(index).spellBuffer.Timer = timeGetTime
        TempPlayer(index).spellBuffer.target = TempPlayer(index).target
        TempPlayer(index).spellBuffer.tType = TempPlayer(index).targetType
        Exit Sub
    Else
        SendClearSpellBuffer index
    End If
End Sub

Public Sub CastSpell(ByVal index As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Byte)
Dim spellNum As Long
Dim MPCost As Long
Dim LevelReq As Long
Dim MapNum As Long
Dim Vital As Long
Dim DidCast As Boolean
Dim ClassReq As Long
Dim AccessReq As Long
Dim i As Long
Dim AoE As Long
Dim Range As Long
Dim VitalType As Byte
Dim increment As Boolean
Dim x As Long, y As Long
Dim Buffer As clsBuffer
Dim SpellCastType As Long
   
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    spellNum = GetPlayerSpell(index, spellslot)
    MapNum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, spellNum) Then Exit Sub

    MPCost = Spell(spellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
   
    LevelReq = Spell(spellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
   
    AccessReq = Spell(spellNum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
   
    ClassReq = Spell(spellNum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(ClassReq).name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
   
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
   
    ' set the vital
    Vital = Spell(spellNum).Vital
    AoE = Spell(spellNum).AoE
    Range = Spell(spellNum).Range
   
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_HEALHP
                    SpellPlayer_Effect Vitals.HP, True, index, Vital, spellNum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, index, Vital, spellNum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation MapNum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    PlayerWarp index, Spell(spellNum).Map, Spell(spellNum).x, Spell(spellNum).y
                    SendAnimation GetPlayerMap(index), Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                x = GetPlayerX(index)
                y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
               
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(target)
                    y = GetPlayerY(target)
                Else
                    x = MapNpc(MapNum).Npc(target).x
                    y = MapNpc(MapNum).Npc(target).y
                End If
               
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                    SendClearSpellBuffer index
                End If
            End If
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(index, i, True) Then
                                            SendAnimation MapNum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            PlayerAttackPlayer index, i, Vital, spellNum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(i).Num > 0 Then
                            If MapNpc(MapNum).Npc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(MapNum).Npc(i).x, MapNpc(MapNum).Npc(i).y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        SendAnimation MapNum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc index, i, Vital, spellNum
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(spellNum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                   
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect VitalType, increment, i, Vital, spellNum
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(i).Num > 0 Then
                            If MapNpc(MapNum).Npc(i).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(MapNum).Npc(i).x, MapNpc(MapNum).Npc(i).y) Then
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellNum, MapNum
                                End If
                            End If
                        End If
                    Next
            End Select
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub
           
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(target)
                y = GetPlayerY(target)
            Else
                x = MapNpc(MapNum).Npc(target).x
                y = MapNpc(MapNum).Npc(target).y
            End If
               
            If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                PlayerMsg index, "Target not in range.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
           
            Select Case Spell(spellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                                PlayerAttackPlayer index, target, Vital, spellNum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
                                PlayerAttackNpc index, target, Vital, spellNum
                                DidCast = True
                            End If
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellNum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    End If
                   
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, target, True) Then
                                SpellPlayer_Effect VitalType, increment, target, Vital, spellNum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, target, Vital, spellNum
                        End If
                    Else
                        If Spell(spellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, target, True) Then
                                SpellNpc_Effect VitalType, increment, target, Vital, spellNum, MapNum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, target, Vital, spellNum, MapNum
                        End If
                    End If
            End Select
    End Select
   
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - MPCost)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
       
        TempPlayer(index).SpellCD(spellslot) = timeGetTime + (Spell(spellNum).CDTime * 1000)
        Call SendCooldown(index, spellslot)
        SendActionMsg MapNum, Trim$(Spell(spellNum).name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal spellNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation GetPlayerMap(index), Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        SendActionMsg GetPlayerMap(index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, spellNum
        
        If increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) + Damage
            If Spell(spellNum).Duration > 0 Then
                AddHoT_Player index, spellNum
            End If
        ElseIf Not increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) - Damage
        End If
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal spellNum As Long, ByVal MapNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation MapNum, Spell(spellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, index
        SendActionMsg MapNum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(index).x * 32, MapNpc(MapNum).Npc(index).y * 32
        
        ' send the sound
        SendMapSound index, MapNpc(MapNum).Npc(index).x, MapNpc(MapNum).Npc(index).y, SoundEntity.seSpell, spellNum
        
        If increment Then
            If MapNpc(MapNum).Npc(index).Vital(Vital) + Damage <= GetNpcMaxVital(index, Vitals.HP) Then
                MapNpc(MapNum).Npc(index).Vital(Vital) = MapNpc(MapNum).Npc(index).Vital(Vital) + Damage
            Else
                MapNpc(MapNum).Npc(index).Vital(Vital) = GetNpcMaxVital(index, Vitals.HP)
            End If
            
            If Spell(spellNum).Duration > 0 Then
                AddHoT_Npc MapNum, index, spellNum
            End If
        ElseIf Not increment Then
            MapNpc(MapNum).Npc(index).Vital(Vital) = MapNpc(MapNum).Npc(index).Vital(Vital) - Damage
        End If
    End If
End Sub

Public Sub AddDoT_Player(ByVal index As Long, ByVal spellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            If .Spell = spellNum Then
                .Timer = timeGetTime
                .Caster = Caster
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellNum
                .Timer = timeGetTime
                .Caster = Caster
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal index As Long, ByVal spellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).HoT(i)
            If .Spell = spellNum Then
                .Timer = timeGetTime
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellNum
                .Timer = timeGetTime
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal spellNum As Long, ByVal Caster As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).Npc(index).DoT(i)
            If .Spell = spellNum Then
                .Timer = timeGetTime
                .Caster = Caster
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellNum
                .Timer = timeGetTime
                .Caster = Caster
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal spellNum As Long)
Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).Npc(index).HoT(i)
            If .Spell = spellNum Then
                .Timer = timeGetTime
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellNum
                .Timer = timeGetTime
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal index As Long, ByVal dotNum As Long)
    With TempPlayer(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, index, True) Then
                    PlayerAttackPlayer .Caster, index, Spell(.Spell).Vital
                End If
                .Timer = timeGetTime
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If timeGetTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal index As Long, ByVal hotNum As Long)
    With TempPlayer(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                   SendActionMsg Player(index).Map, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, Player(index).x * 32, Player(index).y * 32
                   SetPlayerVital index, Vitals.HP, GetPlayerVital(index, Vitals.HP) + Spell(.Spell).Vital
                Else
                   SendActionMsg Player(index).Map, "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, Player(index).x * 32, Player(index).y * 32
                   SetPlayerVital index, Vitals.MP, GetPlayerVital(index, Vitals.MP) + Spell(.Spell).Vital
                End If
                .Timer = timeGetTime
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If timeGetTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal dotNum As Long)
    With MapNpc(MapNum).Npc(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, index, True) Then
                    PlayerAttackNpc .Caster, index, Spell(.Spell).Vital, , True
                End If
                .Timer = timeGetTime
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If timeGetTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal hotNum As Long)
    With MapNpc(MapNum).Npc(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                    SendActionMsg MapNum, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(index).x * 32, MapNpc(MapNum).Npc(index).y * 32
                    MapNpc(MapNum).Npc(index).Vital(Vitals.HP) = MapNpc(MapNum).Npc(index).Vital(Vitals.HP) + Spell(.Spell).Vital
                    
                    If MapNpc(MapNum).Npc(index).Vital(Vitals.HP) > GetNpcMaxVital(index, Vitals.HP) Then
                        MapNpc(MapNum).Npc(index).Vital(Vitals.HP) = GetNpcMaxVital(index, Vitals.HP)
                    End If
                Else
                    SendActionMsg MapNum, "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(index).x * 32, MapNpc(MapNum).Npc(index).y * 32
                    MapNpc(MapNum).Npc(index).Vital(Vitals.MP) = MapNpc(MapNum).Npc(index).Vital(Vitals.MP) + Spell(.Spell).Vital
                    
                    If MapNpc(MapNum).Npc(index).Vital(Vitals.MP) > GetNpcMaxVital(index, Vitals.MP) Then
                        MapNpc(MapNum).Npc(index).Vital(Vitals.MP) = GetNpcMaxVital(index, Vitals.MP)
                    End If
                End If
                
                .Timer = timeGetTime
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If timeGetTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunPlayer(ByVal index As Long, ByVal spellNum As Long)
    ' check if it's a stunning spell
    If Spell(spellNum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(index).StunDuration = Spell(spellNum).StunDuration
        TempPlayer(index).StunTimer = timeGetTime + (Spell(spellNum).StunDuration * 1000)
        ' send it to the index
        SendStunned index
        ' tell him he's stunned
        PlayerMsg index, "You have been stunned.", BrightRed
    End If
End Sub

Public Sub StunNPC(ByVal index As Long, ByVal MapNum As Long, ByVal spellNum As Long)
    ' check if it's a stunning spell
    If Spell(spellNum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(MapNum).Npc(index).StunDuration = Spell(spellNum).StunDuration
        MapNpc(MapNum).Npc(index).StunTimer = timeGetTime
    End If
End Sub

