Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim FileName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\data files\logs\errors.txt"
    Open FileName For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleError", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function FileExist(ByVal FileName As String) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If LenB(Dir(FileName)) > 0 Then
        FileExist = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' Gets a string from a INI file
Public Function GetVar(file As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), file)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' Writes a variable to a INI file
Public Sub PutVar(file As String, Header As String, Var As String, value As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call WritePrivateProfileString$(Header, Var, value, file)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveOptions()
Dim FileName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\Data Files\config.ini"
    
    Call PutVar(FileName, "Options", "Game_Name", Trim$(Options.Game_Name))
    Call PutVar(FileName, "Options", "Username", Trim$(Options.Username))
    Call PutVar(FileName, "Options", "Password", Trim$(Options.Password))
    Call PutVar(FileName, "Options", "SavePass", Str(Options.SavePass))
    Call PutVar(FileName, "Options", "IP", Options.IP)
    Call PutVar(FileName, "Options", "Port", Str(Options.Port))
    Call PutVar(FileName, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Call PutVar(FileName, "Options", "Music", Str(Options.Music))
    Call PutVar(FileName, "Options", "Sound", Str(Options.Sound))
    Call PutVar(FileName, "Options", "Debug", Str(Options.Debug))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOptions()
Dim FileName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\data files\config.ini"
    
    If Not FileExist(FileName) Then
        Options.Game_Name = "EVB"
        Options.Password = vbNullString
        Options.SavePass = 0
        Options.Username = vbNullString
        Options.IP = "127.0.0.1"
        Options.Port = 7001
        Options.MenuMusic = vbNullString
        Options.Music = 1
        Options.Sound = 1
        Options.Debug = 0
        SaveOptions
    Else
        Options.Game_Name = GetVar(FileName, "Options", "Game_Name")
        Options.Username = GetVar(FileName, "Options", "Username")
        Options.Password = GetVar(FileName, "Options", "Password")
        Options.SavePass = Val(GetVar(FileName, "Options", "SavePass"))
        Options.IP = GetVar(FileName, "Options", "IP")
        Options.Port = Val(GetVar(FileName, "Options", "Port"))
        Options.MenuMusic = GetVar(FileName, "Options", "MenuMusic")
        Options.Music = GetVar(FileName, "Options", "Music")
        Options.Sound = GetVar(FileName, "Options", "Sound")
        Options.Debug = GetVar(FileName, "Options", "Debug")
    End If
    
    ' show in GUI
    If Options.Music = 0 Then
        frmMain.optMOff.value = True
    Else
        frmMain.optMOn.value = True
    End If
    
    If Options.Sound = 0 Then
        frmMain.optSOff.value = True
    Else
        frmMain.optSOn.value = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMap(ByVal mapNum As Long)
Dim FileName As String
Dim f As Long
Dim x As Long, y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & MAP_PATH & "map" & mapNum & MAP_EXT

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Map.name
    Put #f, , Map.Music
    Put #f, , Map.Revision
    Put #f, , Map.Moral
    Put #f, , Map.Up
    Put #f, , Map.Down
    Put #f, , Map.Left
    Put #f, , Map.Right
    Put #f, , Map.BootMap
    Put #f, , Map.BootX
    Put #f, , Map.BootY
    Put #f, , Map.MaxX
    Put #f, , Map.MaxY

    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            Put #f, , Map.Tile(x, y)
        Next

        DoEvents
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #f, , Map.Npc(x)
    Next

    Close #f
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMap(ByVal mapNum As Long)
Dim FileName As String
Dim f As Long
Dim x As Long, y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & MAP_PATH & "map" & mapNum & MAP_EXT
    
    ' Get rid of any left-over data
    ClearMap
    
    f = FreeFile
    Open FileName For Binary As #f
    Get #f, , Map.name
    Get #f, , Map.Music
    Get #f, , Map.Revision
    Get #f, , Map.Moral
    Get #f, , Map.Up
    Get #f, , Map.Down
    Get #f, , Map.Left
    Get #f, , Map.Right
    Get #f, , Map.BootMap
    Get #f, , Map.BootX
    Get #f, , Map.BootY
    Get #f, , Map.MaxX
    Get #f, , Map.MaxY
    
    ' have to set the tile()
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            Get #f, , Map.Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Get #f, , Map.Npc(x)
    Next

    Close #f
    ClearTempTile
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckTilesets()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    Do While FileExist(App.Path & GFX_PATH & "\tilesets\" & i & GFX_EXT)
        NumTileSets = NumTileSets + 1
        i = i + 1
    Loop
    
    If NumTileSets = 0 Then Exit Sub
    
    ReDim DDS_Tileset(1 To NumTileSets)
    ReDim DDSD_Tileset(1 To NumTileSets)
    ReDim TilesetTimer(1 To NumTileSets)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckTilesets", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckCharacters()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    Do While FileExist(App.Path & GFX_PATH & "characters\" & i & GFX_EXT)
        NumCharacters = NumCharacters + 1
        i = i + 1
    Loop
    
    If NumCharacters = 0 Then Exit Sub

    ReDim DDS_Character(1 To NumCharacters)
    ReDim DDSD_Character(1 To NumCharacters)
    ReDim CharacterTimer(1 To NumCharacters)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckCharacters", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckPaperdolls()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    Do While FileExist(App.Path & GFX_PATH & "paperdolls\" & i & GFX_EXT)
        NumPaperdolls = NumPaperdolls + 1
        i = i + 1
    Loop
    
    If NumPaperdolls = 0 Then Exit Sub

    ReDim DDS_Paperdoll(1 To NumPaperdolls)
    ReDim DDSD_Paperdoll(1 To NumPaperdolls)
    ReDim PaperdollTimer(1 To NumPaperdolls)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckPaperdolls", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimations()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    Do While FileExist(App.Path & GFX_PATH & "animations\" & i & GFX_EXT)
        NumAnimations = NumAnimations + 1
        i = i + 1
    Loop
    
    If NumAnimations = 0 Then Exit Sub

    ReDim DDS_Animation(1 To NumAnimations)
    ReDim DDSD_Animation(1 To NumAnimations)
    ReDim AnimationTimer(1 To NumAnimations)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    Do While FileExist(App.Path & GFX_PATH & "Items\" & i & GFX_EXT)
        NumItems = NumItems + 1
        i = i + 1
    Loop
    
    If NumItems = 0 Then Exit Sub

    ReDim DDS_Item(1 To NumItems)
    ReDim DDSD_Item(1 To NumItems)
    ReDim ItemTimer(1 To NumItems)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckResources()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    Do While FileExist(App.Path & GFX_PATH & "Resources\" & i & GFX_EXT)
        NumResources = NumResources + 1
        i = i + 1
    Loop
    
    If NumResources = 0 Then Exit Sub

    ReDim DDS_Resource(1 To NumResources)
    ReDim DDSD_Resource(1 To NumResources)
    ReDim ResourceTimer(1 To NumResources)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckSpellIcons()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    Do While FileExist(App.Path & GFX_PATH & "SpellIcons\" & i & GFX_EXT)
        NumSpellIcons = NumSpellIcons + 1
        i = i + 1
    Loop
    
    If NumSpellIcons = 0 Then Exit Sub

    ReDim DDS_SpellIcon(1 To NumSpellIcons)
    ReDim DDSD_SpellIcon(1 To NumSpellIcons)
    ReDim SpellIconTimer(1 To NumSpellIcons)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckSpellIcons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckFaces()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    i = 1
    Do While FileExist(App.Path & GFX_PATH & "Faces\" & i & GFX_EXT)
        NumFaces = NumFaces + 1
        i = i + 1
    Loop
    
    If NumFaces = 0 Then Exit Sub

    ReDim DDS_Face(1 To NumFaces)
    ReDim DDSD_Face(1 To NumFaces)
    ReDim FaceTimer(1 To NumFaces)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckFaces", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearPlayer(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Player(index)), LenB(Player(index)))
    Player(index).name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearItem(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(index)), LenB(Item(index)))
    Item(index).name = vbNullString
    Item(index).Desc = vbNullString
    Item(index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearAnimInstance(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(AnimInstance(index)), LenB(AnimInstance(index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimInstance", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearAnimation(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(index)), LenB(Animation(index)))
    Animation(index).name = vbNullString
    Animation(index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearAnimations()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearNPC(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Npc(index)), LenB(Npc(index)))
    Npc(index).name = vbNullString
    Npc(index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNPC", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearNPCs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearSpell(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Spell(index)), LenB(Spell(index)))
    Spell(index).name = vbNullString
    Spell(index).Desc = vbNullString
    Spell(index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearSpells()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearShop(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Shop(index)), LenB(Shop(index)))
    Shop(index).name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearShops()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearResource(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(index)), LenB(Resource(index)))
    Resource(index).name = vbNullString
    Resource(index).SuccessMessage = vbNullString
    Resource(index).EmptyMessage = vbNullString
    Resource(index).Sound = "None."
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearResources()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMapItem(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapItem(index)), LenB(MapItem(index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.name = vbNullString
    Map.MaxX = MAX_MAPX
    Map.MaxY = MAX_MAPY
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMapItems()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMapNpc(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapNpc(index)), LenB(MapNpc(index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMapNpcs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearConvs()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_CONVS
        Call ClearConv(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearConvs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearConv(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Conv(index)), LenB(Conv(index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearConv", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearQuests()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearQuest", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearQuest(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Quest(index)), LenB(Quest(index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearQuest", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ********************
' * Player functions *
' ********************
Public Function GetPlayerName(ByVal index As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(index).name)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerName(ByVal index As Long, ByVal name As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).name = name
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerClass(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(index).Class
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Class = ClassNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerSprite(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(index).Sprite
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Sprite = Sprite
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerLevel(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(index).Level
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Level = Level
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerExp(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(index).EXP
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerExp(ByVal index As Long, ByVal EXP As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).EXP = EXP
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerAccess(ByVal index As Long) As Byte
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(index).Access
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Access = Access
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(index).Vital(Vital)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerVital(ByVal index As Long, ByVal Vital As Vitals, ByVal value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Vital(Vital) = value

    If GetPlayerVital(index, Vital) > GetPlayerMaxVital(index, Vital) Then
        Player(index).Vital(Vital) = GetPlayerMaxVital(index, Vital)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerMaxVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    
    GetPlayerMaxVital = Player(index).MaxVital(Vital)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMaxVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function GetPlayerStat(ByVal index As Long, Stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(index).Stat(Stat)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerStat(ByVal index As Long, Stat As Stats, ByVal value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    If value <= 0 Then value = 1
    If value > MAX_BYTE Then value = MAX_BYTE
    Player(index).Stat(Stat) = value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerPOINTS(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(index).POINTS
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).POINTS = POINTS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SetPlayerDataAmountLeft(ByVal index As Long, ByVal QuestNum As Long, ByVal DataAmountLeft As Long, ByVal ObjectiveNum As Long)
If index > MAX_PLAYERS Then Exit Sub
    Player(index).Quest(QuestNum).DataAmountLeft(ObjectiveNum) = DataAmountLeft
End Sub

Sub SetPlayerQuestStatus(ByVal index As Long, ByVal QuestNum As Long, ByVal QuestStatus As Long)
If index > MAX_PLAYERS Then Exit Sub
    Player(index).Quest(QuestNum).QuestStatus = QuestStatus
End Sub

Public Function GetPlayerMap(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or index <= 0 Then Exit Function
    GetPlayerMap = Player(index).Map
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerMap(ByVal index As Long, ByVal mapNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Map = mapNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerX(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(index).x
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).x = x
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerY(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(index).y
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).y = y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerDir(ByVal index As Long) As Byte
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(index).Dir
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    Player(index).Dir = Dir
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(invSlot).num
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long, ByVal itemNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invSlot).num = itemNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invSlot).value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invSlot).value = ItemValue
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Equipment) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerEquipment = Player(index).Equipment(EquipmentSlot)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerEquipment(ByVal index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    Player(index).Equipment(EquipmentSlot) = InvNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
