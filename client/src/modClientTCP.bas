Attribute VB_Name = "modClientTCP"
Option Explicit

Private PlayerBuffer As clsBuffer

Public Sub TcpInit()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set PlayerBuffer = New clsBuffer

    ' connect
    frmMain.Socket.RemoteHost = Options.IP
    frmMain.Socket.RemotePort = Options.Port

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TcpInit", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyTCP()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmMain.Socket.close
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DestroyTCP", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim PacketLength As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Receive the data
    frmMain.Socket.GetData Buffer, vbUnicode, DataLength
    PlayerBuffer.WriteBytes Buffer()
    
    ' Read the packet's length first of all
    If PlayerBuffer.Length >= 4 Then PacketLength = PlayerBuffer.ReadLong(False)
    
    ' Loop through and get all our data
    Do While PacketLength > 0 And PacketLength <= PlayerBuffer.Length - 4
        If PacketLength <= PlayerBuffer.Length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(PacketLength)
        End If

        PacketLength = 0
        If PlayerBuffer.Length >= 4 Then PacketLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "IncomingData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConnectToServer(ByVal i As Long) As Boolean
Dim Wait As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = timeGetTime
    frmMain.Socket.close
    frmMain.Socket.Connect
    
    SetStatus "Connecting to server..."
    
    ' Wait until connected or 5 seconds have passed and report the server being down
    Do While (Not IsConnected) And (timeGetTime <= Wait + TIMEOUT)
        DoEvents
    Loop
    
    ConnectToServer = IsConnected

    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConnectToServer", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsConnected() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If frmMain.Socket.state = sckConnected Then
        IsConnected = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsConnected", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsPlaying(ByVal Index As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if the player doesn't exist, the name will equal 0
    If LenB(GetPlayerName(Index)) > 0 Then
        IsPlaying = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlaying", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SendData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If IsConnected Then
        Set Buffer = New clsBuffer
                
        Buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        Buffer.WriteBytes Data()
        frmMain.Socket.SendData Buffer.ToArray()
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ***************************
' * Outgoing client packets *
' ***************************
Public Sub SendNewAccount(ByVal name As String, ByVal Password As String)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNewAccount
    Buffer.WriteString name
    Buffer.WriteString Password
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendNewAccount", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDelAccount(ByVal name As String, ByVal Password As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDelAccount
    Buffer.WriteString name
    Buffer.WriteString Password
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDelAccount", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendLogin(ByVal name As String, ByVal Password As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CLogin
    Buffer.WriteString name
    Buffer.WriteString Password
    
    Buffer.WriteLong VERSION_MAJOR
    Buffer.WriteLong VERSION_MINOR
    Buffer.WriteLong VERSION_REVISION
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendLogin", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendAddChar(ByVal name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Sprite As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAddChar
    
    Buffer.WriteString name
    Buffer.WriteLong Sex
    Buffer.WriteLong ClassNum
    Buffer.WriteLong Sprite
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAddChar", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUseChar()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseChar
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUseChar", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SayMsg(ByVal Text As String)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSayMsg
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SayMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BroadcastMsg(ByVal Text As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBroadcastMsg
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BroadcastMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EmoteMsg(ByVal Text As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CEmoteMsg
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EmoteMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMsg
    Buffer.WriteString MsgTo
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerMove()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMove
    Buffer.WriteLong GetPlayerDir(MyIndex)
    Buffer.WriteLong Player(MyIndex).Moving
    Buffer.WriteLong Player(MyIndex).x
    Buffer.WriteLong Player(MyIndex).y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerMove", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerDir()
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerDir
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerDir", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerRequestNewMap()
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNewMap
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerRequestNewMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMap()
Dim x As Long, y As Long
Dim i As Long
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    CanMoveNow = False
    
    ' General properties
    With Map
        Buffer.WriteLong CMapData
        Buffer.WriteString Trim$(.name)
        Buffer.WriteString Trim$(.Music)
        Buffer.WriteByte .Moral
        Buffer.WriteLong .Up
        Buffer.WriteLong .Down
        Buffer.WriteLong .Left
        Buffer.WriteLong .Right
        Buffer.WriteLong .BootMap
        Buffer.WriteByte .BootX
        Buffer.WriteByte .BootY
        Buffer.WriteByte .MaxX
        Buffer.WriteByte .MaxY
    End With
    
    ' Tiles & attributes
    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            With Map.Tile(x, y)
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
    
    ' NPCs
    With Map
        For x = 1 To MAX_MAP_NPCS
            Buffer.WriteLong .Npc(x)
        Next
    End With

    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpMeTo(ByVal name As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpMeTo
    Buffer.WriteString name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarpMeTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpToMe(ByVal name As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpToMe
    Buffer.WriteString name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarptoMe", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpTo(ByVal mapNum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpTo
    Buffer.WriteLong mapNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WarpTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetAccess(ByVal name As String, ByVal Access As Byte)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetAccess
    Buffer.WriteString name
    Buffer.WriteLong Access
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSetAccess", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetSprite(ByVal spriteNum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetSprite
    Buffer.WriteLong spriteNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSetSprite", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendKick(ByVal name As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CKickPlayer
    Buffer.WriteString name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendKick", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBan(ByVal name As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanPlayer
    Buffer.WriteString name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBan", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBanList()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanList
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBanList", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditItem()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditItem
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveItem(ByVal itemNum As Long)
Dim Buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    ' Copy the memory and put it into a binary packet
    ItemSize = LenB(Item(itemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemNum)), ItemSize
    
    Buffer.WriteLong CSaveItem
    Buffer.WriteLong itemNum
    Buffer.WriteBytes ItemData
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditAnimation()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditAnimation
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditAnimation", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveAnimation(ByVal Animationnum As Long)
Dim Buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    ' Copy the memory and put it into a binary packet
    AnimationSize = LenB(Animation(Animationnum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(Animationnum)), AnimationSize
    
    Buffer.WriteLong CSaveAnimation
    Buffer.WriteLong Animationnum
    Buffer.WriteBytes AnimationData
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveAnimation", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditNpc()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditNpc
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditNpc", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveNpc(ByVal NpcNum As Long)
Dim Buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    ' Copy the memory and put it into a binary packet
    NpcSize = LenB(Npc(NpcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(Npc(NpcNum)), NpcSize
    
    Buffer.WriteLong CSaveNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteBytes NpcData
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveNpc", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditResource()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditResource
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditResource", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveResource(ByVal ResourceNum As Long)
Dim Buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    ' Copy the memory and put it into a binary packet
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong CSaveResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveResource", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMapRespawn()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapRespawn
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapRespawn", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUseItem(ByVal InvNum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseItem
    Buffer.WriteLong InvNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUseItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDropItem(ByVal InvNum As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If InBank Or InShop Then Exit Sub
    If InvNum < 1 Or InvNum > MAX_INV Then Exit Sub
    If PlayerInv(InvNum).num < 1 Or PlayerInv(InvNum).num > MAX_ITEMS Then Exit Sub
    
    ' Make sure the currency is a suitable amount
    If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
        If Amount < 1 Or Amount > PlayerInv(InvNum).value Then Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapDropItem
    Buffer.WriteLong InvNum
    Buffer.WriteLong Amount
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDropItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendWhosOnline()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWhosOnline
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendWhosOnline", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetMotd
    Buffer.WriteString MOTD
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMOTDChange", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditShop()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditShop
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditShop", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveShop(ByVal shopnum As Long)
Dim Buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    ' Copy the memory and put it into a binary packet
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopnum)), ShopSize
    
    Buffer.WriteLong CSaveShop
    Buffer.WriteLong shopnum
    Buffer.WriteBytes ShopData
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveShop", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditSpell()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditSpell
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditSpell", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveSpell(ByVal spellnum As Long)
Dim Buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    ' Copy the memory and put it into a binary packet
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    
    Buffer.WriteLong CSaveSpell
    Buffer.WriteLong spellnum
    Buffer.WriteBytes SpellData
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveSpell", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditMap()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBanDestroy()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanDestroy
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBanDestroy", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendChangeInvSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapInvSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendChangeSpellSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapSpellSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub GetPing()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    PingStart = timeGetTime
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCheckPing
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GetPing", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUnequip(ByVal EqNum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUnequip
    Buffer.WriteLong EqNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUnequip", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestPlayerData()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestPlayerData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestPlayerData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestItems()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestItems
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestItems", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestAnimations()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestAnimations
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestAnimations", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestNPCS()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNPCS
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestNPCS", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestResources()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestResources
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestResources", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestSpells()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestSpells
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestSpells", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestShops()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestShops
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestShops", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSpawnItem
    Buffer.WriteLong tmpItem
    Buffer.WriteLong tmpAmount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSpawnItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendTrainStat(ByVal StatNum As Byte)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseStatPoint
    Buffer.WriteByte StatNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTrainStat", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestLevelUp()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestLevelUp
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestLevelUp", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BuyItem(ByVal shopSlot As Long)
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBuyItem
    Buffer.WriteLong shopSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BuyItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SellItem(ByVal invSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSellItem
    Buffer.WriteLong invSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SellItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DepositItem(ByVal invSlot As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDepositItem
    Buffer.WriteLong invSlot
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DepositItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WithdrawItem(ByVal bankslot As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWithdrawItem
    Buffer.WriteLong bankslot
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WithdrawItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseBank()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCloseBank
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Hide the bank GUI
    InBank = False
    frmMain.picCover.Visible = False
    frmMain.picBank.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CloseBank", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChangeBankSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CChangeBankSlots
    Buffer.WriteLong OldSlot
    Buffer.WriteLong NewSlot
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChangeBankSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AdminWarp(ByVal x As Long, ByVal y As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAdminWarp
    Buffer.WriteLong x
    Buffer.WriteLong y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AdminWarp", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AcceptTrade()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AcceptTrade", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DeclineTrade()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DeclineTrade", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub TradeItem(ByVal invSlot As Long, ByVal Amount As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeItem
    Buffer.WriteLong invSlot
    Buffer.WriteLong Amount
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TradeItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UntradeItem(ByVal invSlot As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUntradeItem
    Buffer.WriteLong invSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UntradeItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHotbarChange(ByVal sType As Long, ByVal Slot As Long, ByVal hotbarNum As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CHotbarChange
    Buffer.WriteLong sType
    Buffer.WriteLong Slot
    Buffer.WriteLong hotbarNum
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendHotbarChange", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHotbarUse(ByVal Slot As Long)
Dim Buffer As clsBuffer, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check if spell
    If Hotbar(Slot).sType = 2 Then ' spell
        For i = 1 To MAX_PLAYER_SPELLS
            ' is the spell matching the hotbar?
            If PlayerSpells(i) = Hotbar(Slot).Slot Then
                If SpellBuffer = i Then Exit Sub
                ' found it, cast it
                CastSpell i
                Exit Sub
            End If
        Next
        
        ' can't find the spell, exit out
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong CHotbarUse
    Buffer.WriteLong Slot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendHotbarUse", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMapReport()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapReport
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapReport", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PlayerSearch(ByVal CurX As Long, ByVal CurY As Long)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If isInBounds Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CSearch
        Buffer.WriteLong CurX
        Buffer.WriteLong CurY
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerSearch", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendTradeRequest()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendAcceptTradeRequest()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAcceptTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDeclineTradeRequest()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineTradeRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPartyLeave()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyLeave
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPartyLeave", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPartyRequest()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyRequest
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPartyRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendAcceptParty()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAcceptParty
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAcceptParty", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDeclineParty()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDeclineParty
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDeclineParty", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPartyChatMsg(ByVal Text As String)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CPartyChatMsg
    Buffer.WriteString Text
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPartyChatMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditConv()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditConv
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditConv", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveConv(ByVal ConvNum As Long)
Dim Buffer As clsBuffer
Dim ConvSize As Long
Dim ConvData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    ' Copy the memory and put it into a binary packet
    ConvSize = LenB(Conv(ConvNum))
    ReDim ConvData(ConvSize - 1)
    CopyMemory ConvData(0), ByVal VarPtr(Conv(ConvNum)), ConvSize
    
    Buffer.WriteLong CSaveConv
    Buffer.WriteLong ConvNum
    Buffer.WriteBytes ConvData
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveConv", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestConvs()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestConvs
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestItems", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendConvEvent(ByVal CurChat As Long, ByVal ConvIndex As Long, ByVal Choice As Byte)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSendConvEvent
    Buffer.WriteLong CurChat
    Buffer.WriteLong ConvIndex
    Buffer.WriteLong Choice
    SendData Buffer.ToArray
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendConvEvent", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendCloseConv()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCloseConv
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendCloseConv", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendTargetUpdate(ByVal TargetIndex As Long, TargetType As Byte)
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CTargetUpdate
    Buffer.WriteLong TargetIndex
    Buffer.WriteByte TargetType
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTargetUpdate", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditQuest()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditQuest
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestEditQuest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveQuest(ByVal QuestNum As Long)
Dim Buffer As clsBuffer
Dim QuestSize As Long
Dim QuestData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    ' Copy the memory and put it into a binary packet
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    
    Buffer.WriteLong CSaveQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteBytes QuestData
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSaveQuest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestQuests()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestQuests
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestQuests", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestQuestInfo(ByVal QuestName As String)
Dim Buffer As clsBuffer
Dim i As Long, QuestNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Find the quest's index
    For i = 1 To MAX_QUESTS
        If Trim$(Quest(i).name) = Trim$(QuestName) Then
            QuestNum = i
            Exit For
        End If
    Next
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestQuestInfo
    Buffer.WriteLong QuestNum
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestQuestInfo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendQuestDrop(ByVal QuestName As String)
Dim Buffer As clsBuffer
Dim i As Long, QuestNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Find the quest's index
    For i = 1 To MAX_QUESTS
        If Trim$(Quest(i).name) = Trim$(QuestName) Then
            QuestNum = i
            Exit For
        End If
    Next
    
    ' Make sure we got a quest
    If QuestNum = 0 Then Exit Sub
    
    ' Drop it
    For i = 1 To 3
        Player(MyIndex).Quest(QuestNum).DataAmountLeft(i) = 0
    Next
    Player(MyIndex).Quest(QuestNum).QuestStatus = 0
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CQuestDrop
    
    Buffer.WriteLong QuestNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Populate the quest list
    If frmMain.picQuests.Visible Then
        frmMain.lstQuests.Clear
        
        For i = 1 To MAX_QUESTS
            If Player(MyIndex).Quest(i).QuestStatus = 1 Then
                frmMain.lstQuests.AddItem Quest(i).name
            End If
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendQuestDrop", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestAllEditors()
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestAllEditors
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendRequestAllEditors", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
