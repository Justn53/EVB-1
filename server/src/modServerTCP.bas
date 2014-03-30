Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = Options.Game_Name & ", Port: " & CStr(frmServer.Socket(0).LocalPort) & ", Online Players: " & TotalOnlinePlayers
End Sub

Sub CreateFullMapCache()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next

End Sub

Function IsConnected(ByVal index As Long) As Boolean

    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal index As Long) As Boolean

    If IsConnected(index) Then
        If TempPlayer(index).InGame Then
            IsPlaying = True
        End If
    End If

End Function

Function IsLoggedIn(ByVal index As Long) As Boolean

    If IsConnected(index) Then
        If LenB(Trim$(Player(index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If

End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If LCase$(Trim$(Player(i).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If

    Next

End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim i As Long
    Dim n As Long

    For i = 1 To Player_HighIndex

        If IsConnected(i) Then
            If Trim$(GetPlayerIP(i)) = IP Then
                n = n + 1

                If (n > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If

    Next

End Function

Function IsBanned(ByVal IP As String) As Boolean
    Dim FileName As String
    Dim fIP As String
    Dim fName As String
    Dim F As Long
    FileName = App.Path & "\data\banlist.txt"

    ' Check if file exists
    If Not FileExist(App.Path & "\data\banlist.txt") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If

    F = FreeFile
    Open FileName For Input As #F

    Do While Not EOF(F)
        Input #F, fIP
        Input #F, fName

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function

Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim TempData() As Byte

    If IsConnected(index) Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
              
        frmServer.Socket(index).SendData Buffer.ToArray()
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If

    Next

End Sub

Sub SendDataToAllBut(ByVal index As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> index Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If i <> index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If

    Next

End Sub

Sub SendDataToParty(ByVal partyNum As Long, ByRef Data() As Byte)
Dim i As Long

    For i = 1 To Party(partyNum).MemberCount
        If Party(partyNum).Member(i) > 0 Then
            Call SendDataTo(Party(partyNum).Member(i), Data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SGlobalMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataToAll Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAdminMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            SendDataTo i, Buffer.ToArray
        End If
    Next
    
    Set Buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataTo index, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataToMap MapNum, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SAlertMsg
    Buffer.WriteString Msg
    SendDataTo index, Buffer.ToArray
    DoEvents
    Call CloseSocket(index)
    
    Set Buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal partyNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim i As Long
    ' send message to all people
    For i = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partyNum).Member(i) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(partyNum).Member(i)) And IsPlaying(Party(partyNum).Member(i)) Then
                PlayerMsg Party(partyNum).Member(i), Msg, Color
            End If
        End If
    Next
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)

    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If

        Call AlertMsg(index, "You have lost your connection with " & Options.Game_Name & ".")
    End If

End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If

End Sub

Sub SocketConnected(ByVal index As Long)
Dim i As Long

    If index <> 0 Then
        ' make sure they're not banned
        If Not IsBanned(GetPlayerIP(index)) Then
            Call TextAdd("Received connection from " & GetPlayerIP(index) & ".")
        Else
            Call AlertMsg(index, "You have been banned from " & Options.Game_Name & ", and can no longer play.")
        End If
        ' re-set the high index
        Player_HighIndex = 0
        For i = MAX_PLAYERS To 1 Step -1
            If IsConnected(i) Then
                Player_HighIndex = i
                Exit For
            End If
        Next
        ' send the new highindex to all logged in players
        SendHighIndex
    End If
End Sub

Public Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim PacketLength As Long

    If GetPlayerAccess(index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(index).DataBytes > 1000 Then
            If timeGetTime < TempPlayer(index).DataTimer Then
                Exit Sub
            End If
        End If
    
        ' Check for packet flooding
        If TempPlayer(index).DataPackets > 25 Then
            If timeGetTime < TempPlayer(index).DataTimer Then
                Exit Sub
            End If
        End If
    End If
    
    ' Check if elapsed time has passed
    TempPlayer(index).DataBytes = TempPlayer(index).DataBytes + DataLength
    If timeGetTime >= TempPlayer(index).DataTimer Then
        TempPlayer(index).DataTimer = timeGetTime + 1000
        TempPlayer(index).DataBytes = 0
        TempPlayer(index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(index).Buffer.WriteBytes Buffer()
    
    If TempPlayer(index).Buffer.Length >= 4 Then
        PacketLength = TempPlayer(index).Buffer.ReadLong(False)
    
        If PacketLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While PacketLength > 0 And PacketLength <= TempPlayer(index).Buffer.Length - 4
        If PacketLength <= TempPlayer(index).Buffer.Length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).Buffer.ReadLong
            HandleData index, TempPlayer(index).Buffer.ReadBytes(PacketLength)
        End If
        
        PacketLength = 0
        If TempPlayer(index).Buffer.Length >= 4 Then
            PacketLength = TempPlayer(index).Buffer.ReadLong(False)
        
            If PacketLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(index).Buffer.Trim
End Sub

Sub CloseSocket(ByVal index As Long)

    If index > 0 Then
        Call LeftGame(index)
        Call TextAdd("Connection from " & GetPlayerIP(index) & " has been terminated.")
        frmServer.Socket(index).Close
        Call UpdateCaption
        Call ClearPlayer(index)
    End If

End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
    Dim MapData As String
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong MapNum
    Buffer.WriteString Trim$(Map(MapNum).name)
    Buffer.WriteString Trim$(Map(MapNum).Music)
    Buffer.WriteLong Map(MapNum).Revision
    Buffer.WriteByte Map(MapNum).Moral
    Buffer.WriteLong Map(MapNum).Up
    Buffer.WriteLong Map(MapNum).Down
    Buffer.WriteLong Map(MapNum).Left
    Buffer.WriteLong Map(MapNum).Right
    Buffer.WriteLong Map(MapNum).BootMap
    Buffer.WriteByte Map(MapNum).BootX
    Buffer.WriteByte Map(MapNum).BootY
    Buffer.WriteByte Map(MapNum).MaxX
    Buffer.WriteByte Map(MapNum).MaxY

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY

            With Map(MapNum).Tile(x, y)
                For i = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(i).x
                    Buffer.WriteLong .Layer(i).y
                    Buffer.WriteLong .Layer(i).Tileset
                Next
                Buffer.WriteByte .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
                Buffer.WriteByte .DirBlock
            End With

        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map(MapNum).Npc(x)
    Next

    MapCache(MapNum).Data = Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> index Then
                s = s & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If

    Next

    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If

    Call PlayerMsg(index, s, WhoColor)
End Sub

Function PlayerData(ByVal index As Long) As Byte()
Dim Buffer As clsBuffer, i As Long, j As Long

    If index > MAX_PLAYERS Then Exit Function
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong index
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerLevel(index)
    Buffer.WriteLong GetPlayerPOINTS(index)
    Buffer.WriteLong GetPlayerSprite(index)
    Buffer.WriteLong GetPlayerMap(index)
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerClass(index)
    
    ' Quests
    For i = 1 To MAX_QUESTS
        For j = 1 To MAX_QUEST_OBJECTIVES
            Buffer.WriteLong GetPlayerDataAmountLeft(index, i, j)
        Next
        Buffer.WriteLong GetPlayerQuestStatus(index, i)
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(index, i)
    Next
    
    PlayerData = Buffer.ToArray()
    Set Buffer = Nothing
End Function

Sub SendJoinMap(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> index Then
                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    SendDataTo index, PlayerData(i)
                End If
            End If
        End If
    Next

    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    
    Set Buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLeft
    Buffer.WriteLong index
    SendDataToMapBut index, MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerData(ByVal index As Long)
    Dim packet As String
    SendDataToMap GetPlayerMap(index), PlayerData(index)
End Sub

Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(MapCache(MapNum).Data) - LBound(MapCache(MapNum).Data)) + 5
    Buffer.WriteLong SMapData
    Buffer.WriteBytes MapCache(MapNum).Data()
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(MapNum, i).playerName
        Buffer.WriteLong MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteLong MapItem(MapNum, i).x
        Buffer.WriteLong MapItem(MapNum, i).y
    Next

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Public Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim packet As String
Dim i As Long
Dim Buffer As clsBuffer
    
    ' Check if there's at least one player on the map
    If PlayersOnMap(MapNum) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(MapNum, i).playerName
        Buffer.WriteLong MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteLong MapItem(MapNum, i).x
        Buffer.WriteLong MapItem(MapNum, i).y
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMapNpcVitals(ByVal MapNum As Long, ByVal mapNpcNum As Long)
Dim i As Long
Dim Buffer As clsBuffer
    
    ' Check if there's at least one player on the map
    If PlayersOnMap(MapNum) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapNpcVitals
    
    Buffer.WriteLong mapNpcNum
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong MapNpc(MapNum).Npc(mapNpcNum).Vital(i)
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMapNpcsTo(ByVal index As Long, ByVal MapNum As Long)
Dim i As Long
Dim Buffer As clsBuffer
    
    ' Check if there's at least one player on the map
    If PlayersOnMap(MapNum) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        Buffer.WriteLong MapNpc(MapNum).Npc(i).x
        Buffer.WriteLong MapNpc(MapNum).Npc(i).y
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Vital(HP)
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMapNpcsToMap(ByVal MapNum As Long)
Dim i As Long
Dim Buffer As clsBuffer

    ' Check if there's at least one player on the map
    If PlayersOnMap(MapNum) = 0 Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        Buffer.WriteLong MapNpc(MapNum).Npc(i).x
        Buffer.WriteLong MapNpc(MapNum).Npc(i).y
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Vital(HP)
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendItems(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_ITEMS
        If LenB(Trim$(Item(i).name)) > 0 Then
            Call SendUpdateItemTo(index, i)
        End If
    Next
End Sub

Public Sub SendAnimations(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        If LenB(Trim$(Animation(i).name)) > 0 Then
            Call SendUpdateAnimationTo(index, i)
        End If
    Next
End Sub

Sub SendNpcs(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS

        If LenB(Trim$(Npc(i).name)) > 0 Then
            Call SendUpdateNpcTo(index, i)
        End If

    Next

End Sub

Sub SendResources(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(i).name)) > 0 Then
            Call SendUpdateResourceTo(index, i)
        End If

    Next

End Sub

Public Sub SendConvs(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_CONVS
        If LenB(Trim$(Conv(i).name)) > 0 Then
            Call SendUpdateConvTo(index, i)
        End If
    Next
End Sub

Public Sub SendQuests(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_QUESTS
        If LenB(Trim$(Quest(i).name)) > 0 Then
            Call SendUpdateQuestTo(index, i)
        End If
    Next
End Sub

Sub SendInventory(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(index, i)
        Buffer.WriteLong GetPlayerInvItemValue(index, i)
    Next

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal invSlot As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInvUpdate
    Buffer.WriteLong invSlot
    Buffer.WriteLong GetPlayerInvItemNum(index, invSlot)
    Buffer.WriteLong GetPlayerInvItemValue(index, invSlot)
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerWornEq
    Buffer.WriteLong GetPlayerEquipment(index, Armor)
    Buffer.WriteLong GetPlayerEquipment(index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(index, Helmet)
    Buffer.WriteLong GetPlayerEquipment(index, Shield)
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal index As Long)
Dim Buffer As clsBuffer

    ' Check if there's at least one player on the map
    If PlayersOnMap(GetPlayerMap(index)) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapWornEq
    
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerEquipment(index, Armor)
    Buffer.WriteLong GetPlayerEquipment(index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(index, Helmet)
    Buffer.WriteLong GetPlayerEquipment(index, Shield)
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong PlayerNum
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Armor)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Weapon)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Helmet)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Shield)
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendVital(ByVal index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Select Case Vital
        Case HP
            Buffer.WriteLong SPlayerHp
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.HP)
        Case MP
            Buffer.WriteLong SPlayerMp
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.MP)
    End Select

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendEXP(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerEXP
    Buffer.WriteLong GetPlayerExp(index)
    Buffer.WriteLong GetPlayerNextLevel(index)
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendStats(ByVal index As Long)
Dim i As Long
Dim packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(index, i)
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendWelcome(ByVal index As Long)

    ' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(index, Options.MOTD, BrightCyan)
    End If

    ' Send whos online
    Call SendWhosOnline(index)
End Sub

Sub SendClasses(ByVal index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClassesData
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        
        ' send array size
        Buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        
        ' send array size
        Buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).Stat(q)
        Next
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNewCharClasses
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        ' send array size
        Buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        ' send array size
        Buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).Stat(q)
        Next
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLeftGame(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong index
    Buffer.WriteString vbNullString
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    SendDataToAllBut index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerXY(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXY
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerXYToMap(ByVal index As Long)
Dim Buffer As clsBuffer
    
    ' Check if there's at least one player on the map
    If PlayersOnMap(GetPlayerMap(index)) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXYMap
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    Set Buffer = New clsBuffer
    
    ' Pack it into a binary packet
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    
    Set Buffer = New clsBuffer
    
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateAnimationTo(ByVal index As Long, ByVal AnimationNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateNpcToAll(ByVal NPCNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set Buffer = New clsBuffer
    NPCSize = LenB(Npc(NPCNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(Npc(NPCNum)), NPCSize
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong NPCNum
    Buffer.WriteBytes NPCData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal NPCNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set Buffer = New clsBuffer
    NPCSize = LenB(Npc(NPCNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(Npc(NPCNum)), NPCSize
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong NPCNum
    Buffer.WriteBytes NPCData
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateResourceTo(ByVal index As Long, ByVal ResourceNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendShops(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(i).name)) > 0 Then
            Call SendUpdateShopTo(index, i)
        End If

    Next

End Sub

Sub SendUpdateShopToAll(ByVal shopNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal shopNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSpells(ByVal index As Long)
    Dim i As Long
 
    For i = 1 To MAX_SPELLS
 
        If LenB(Trim$(Spell(i).name)) > 0 Then
            Call SendUpdateSpellTo(index, i)
        End If
 
    Next
    Call SendPlayerSpells(index)
End Sub

Sub SendUpdateSpellToAll(ByVal spellNum As Long)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(spellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellNum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong spellNum
    Buffer.WriteBytes SpellData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal spellNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(spellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellNum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong spellNum
    Buffer.WriteBytes SpellData
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong GetPlayerSpell(index, i)
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal index As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(GetPlayerMap(index)).Resource_Count

    If ResourceCache(GetPlayerMap(index)).Resource_Count > 0 Then

        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            Buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(i).ResourceState
            Buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(i).x
            Buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(i).y
        Next

    End If

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendResourceCacheToMap(ByVal MapNum As Long, ByVal Resource_num As Long)
Dim Buffer As clsBuffer
Dim i As Long
    
    ' Check if there's at least one player on the map
    If PlayersOnMap(MapNum) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(MapNum).Resource_Count

    If ResourceCache(MapNum).Resource_Count > 0 Then
        For i = 0 To ResourceCache(MapNum).Resource_Count
            Buffer.WriteByte ResourceCache(MapNum).ResourceData(i).ResourceState
            Buffer.WriteLong ResourceCache(MapNum).ResourceData(i).x
            Buffer.WriteLong ResourceCache(MapNum).ResourceData(i).y
        Next
    End If

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendActionMsg(ByVal MapNum As Long, ByVal message As String, ByVal Color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal y As Long, Optional PlayerOnlyNum As Long = 0)
Dim Buffer As clsBuffer
    
    ' Check if there's at least one player on the map
    If PlayersOnMap(MapNum) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SActionMsg
    Buffer.WriteString message
    Buffer.WriteLong Color
    Buffer.WriteLong MsgType
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap MapNum, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Public Sub SendBlood(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim Buffer As clsBuffer
    
    ' Check if there's at least one player on the map
    If PlayersOnMap(MapNum) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBlood
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendAnimation(ByVal MapNum As Long, ByVal Anim As Long, ByVal x As Long, ByVal y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0)
Dim Buffer As clsBuffer
    
    ' Check if there's at least one player on the map
    If PlayersOnMap(MapNum) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimation
    Buffer.WriteLong Anim
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte LockType
    Buffer.WriteLong LockIndex
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendCooldown(ByVal index As Long, ByVal Slot As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCooldown
    Buffer.WriteLong Slot
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendClearSpellBuffer(ByVal index As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearSpellBuffer
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SayMsg_Map(ByVal MapNum As Long, ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
Dim Buffer As clsBuffer
    
    ' Check if there's at least one player on the map
    If PlayersOnMap(MapNum) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteString message
    Buffer.WriteString "[Map] "
    Buffer.WriteLong saycolour
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteString message
    Buffer.WriteString "[Global] "
    Buffer.WriteLong saycolour
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendStunned(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStunned
    Buffer.WriteLong TempPlayer(index).StunDuration
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendBank(ByVal index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        Buffer.WriteLong Bank(index).Item(i).Num
        Buffer.WriteLong Bank(index).Item(i).Value
    Next
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapKey(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte Value
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapKeyToMap(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte Value
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendOpenShop(ByVal index As Long, ByVal shopNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenShop
    Buffer.WriteLong shopNum
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
Dim Buffer As clsBuffer
    
    ' Check if there's at least one player on the map
    If PlayersOnMap(GetPlayerMap(index)) = 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMove
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    Buffer.WriteLong movement
    
    If Not sendToSelf Then
        SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendTrade(ByVal index As Long, ByVal tradeTarget As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STrade
    Buffer.WriteLong tradeTarget
    Buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseTrade
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal index As Long, ByVal dataType As Byte)
Dim Buffer As clsBuffer
Dim i As Long
Dim tradeTarget As Long
Dim totalWorth As Long
    
    tradeTarget = TempPlayer(index).InTrade
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeUpdate
    Buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).Num
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(index).TradeOffer(i).Num > 0 Then
                ' currency?
                If Item(TempPlayer(index).TradeOffer(i).Num).Type = ITEM_TYPE_CURRENCY Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).price * TempPlayer(index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).price
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(i).Value
            ' add total worth
            If GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Type = ITEM_TYPE_CURRENCY Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).price * TempPlayer(tradeTarget).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).price
                End If
            End If
        Next
    End If
    
    ' send total worth of trade
    Buffer.WriteLong totalWorth
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal index As Long, ByVal Status As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeStatus
    Buffer.WriteByte Status
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTarget(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STarget
    Buffer.WriteLong TempPlayer(index).target
    Buffer.WriteLong TempPlayer(index).targetType
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHotbar(ByVal index As Long)
Dim i As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        Buffer.WriteLong Player(index).Hotbar(i).Slot
        Buffer.WriteByte Player(index).Hotbar(i).sType
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLoginOk(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLoginOk
    Buffer.WriteLong index
    Buffer.WriteLong Player_HighIndex
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendInGame(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SInGame
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHighIndex()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHighIndex
    Buffer.WriteLong Player_HighIndex
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMapSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer

    ' Check if there's at least one player on the map
    If PlayersOnMap(GetPlayerMap(index)) = 0 Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal index As Long, ByVal TradeRequest As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeRequest
    Buffer.WriteString Trim$(Player(TradeRequest).name)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal index As Long, ByVal targetPlayer As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyInvite
    Buffer.WriteString Trim$(Player(targetPlayer).name)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal partyNum As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    Buffer.WriteByte 1
    Buffer.WriteLong Party(partyNum).Leader
    For i = 1 To MAX_PARTY_MEMBERS
        Buffer.WriteLong Party(partyNum).Member(i)
    Next
    Buffer.WriteLong Party(partyNum).MemberCount
    SendDataToParty partyNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal index As Long)
Dim Buffer As clsBuffer, i As Long, partyNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' send party data
        Buffer.WriteByte 1
        Buffer.WriteLong Party(partyNum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            Buffer.WriteLong Party(partyNum).Member(i)
        Next
        Buffer.WriteLong Party(partyNum).MemberCount
    Else
        ' send clear command
        Buffer.WriteByte 0
    End If
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal partyNum As Long, ByVal index As Long)
Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyVitals
    Buffer.WriteLong index
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(index, i)
        Buffer.WriteLong Player(index).Vital(i)
    Next
    SendDataToParty partyNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSpawnItemToMap(ByVal MapNum As Long, ByVal index As Long)
Dim Buffer As clsBuffer

    ' Check if there's at least one player on the map
    If PlayersOnMap(MapNum) = 0 Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpawnItem
    Buffer.WriteLong index
    Buffer.WriteString MapItem(MapNum, index).playerName
    Buffer.WriteLong MapItem(MapNum, index).Num
    Buffer.WriteLong MapItem(MapNum, index).Value
    Buffer.WriteLong MapItem(MapNum, index).x
    Buffer.WriteLong MapItem(MapNum, index).y
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub PartyChatMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim i As Long
Dim Member As Long
Dim partyNum As Long

    partyNum = TempPlayer(index).inParty

    ' not in a party?
    If TempPlayer(index).inParty = 0 Then
        Call PlayerMsg(index, "You are not in a party.", BrightRed)
        Exit Sub
    End If

    For i = 1 To MAX_PARTY_MEMBERS
        Member = Party(partyNum).Member(i)
        
        ' is online, does exist?
        If IsConnected(Party(partyNum).Member(i)) And IsPlaying(Party(partyNum).Member(i)) Then
            ' yep, send the message!
            Call PlayerMsg(Member, "[Party] " & GetPlayerName(index) & ": " & Msg, Color)
        End If
    Next
End Sub

Public Sub SendUpdateConvToAll(ByVal ConvNum As Long)
Dim Buffer As clsBuffer
Dim ConvSize As Long
Dim ConvData() As Byte

    Set Buffer = New clsBuffer
    
    ' Pack it into a binary packet
    ConvSize = LenB(Conv(ConvNum))
    ReDim ConvData(ConvSize - 1)
    CopyMemory ConvData(0), ByVal VarPtr(Conv(ConvNum)), ConvSize
    
    Buffer.WriteLong SUpdateConv
    Buffer.WriteLong ConvNum
    Buffer.WriteBytes ConvData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateConvTo(ByVal index As Long, ByVal ConvNum As Long)
Dim Buffer As clsBuffer
Dim ConvSize As Long
Dim ConvData() As Byte
    
    Set Buffer = New clsBuffer
    
    ConvSize = LenB(Conv(ConvNum))
    ReDim ConvData(ConvSize - 1)
    
    CopyMemory ConvData(0), ByVal VarPtr(Conv(ConvNum)), ConvSize
    Buffer.WriteLong SUpdateConv
    Buffer.WriteLong ConvNum
    Buffer.WriteBytes ConvData
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendStartConv(ByVal index As Long, ByVal ConvNum As Long, ByVal NPCNum As Long, Optional ByVal chatNum As Byte = 1)
Dim Buffer As clsBuffer

    ' quest stuff
    If Npc(NPCNum).questGiver = 1 Then
        If Player(index).Quest(Npc(NPCNum).questIndex).QuestStatus = 0 Then
            If Npc(NPCNum).questStartChat > 0 Then chatNum = Npc(NPCNum).questStartChat
        ElseIf Player(index).Quest(Npc(NPCNum).questIndex).QuestStatus = 1 Then
            If Npc(NPCNum).questInProgressChat > 0 Then chatNum = Npc(NPCNum).questInProgressChat
        ElseIf Player(index).Quest(Npc(NPCNum).questIndex).QuestStatus = 2 Then
            If Npc(NPCNum).postQuestChat > 0 Then chatNum = Npc(NPCNum).postQuestChat
        End If
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSendStartConv
    Buffer.WriteLong ConvNum
    Buffer.WriteLong NPCNum
    Buffer.WriteByte chatNum
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendCloseConv(ByVal index As Long)
Dim Buffer As clsBuffer

    TempPlayer(index).InChat = 0
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSendCloseConv
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendUpdateQuestToAll(ByVal QuestNum As Long)
Dim Buffer As clsBuffer
Dim QuestSize As Long
Dim QuestData() As Byte

    Set Buffer = New clsBuffer
    
    ' Pack it into a binary packet
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateQuestTo(ByVal index As Long, ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    
    Set Buffer = New clsBuffer
    
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    Buffer.WriteLong SUpdateQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPopulateList(ByVal index As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPopulateList
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendQuestFinish(ByVal index As Long, QuestNum As Long)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SQuestFinish
    Buffer.WriteLong QuestNum
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

