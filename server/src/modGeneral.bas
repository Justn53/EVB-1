Attribute VB_Name = "modGeneral"
Option Explicit

' Timer functions
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Sub Main()
    Call InitServer
End Sub

Public Sub InitServer()
Dim i As Long
Dim F As Long
Dim StartTime As Long
Dim EndTime As Long

    Call InitMessages
    StartTime = timeGetTime
    frmServer.Show
    
    ' Initialize the random-number generator
    Randomize

    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\data\", "accounts"
    ChkDir App.Path & "\data\", "animations"
    ChkDir App.Path & "\data\", "banks"
    ChkDir App.Path & "\data\", "items"
    ChkDir App.Path & "\data\", "maps"
    ChkDir App.Path & "\data\", "npcs"
    ChkDir App.Path & "\data\", "resources"
    ChkDir App.Path & "\data\", "shops"
    ChkDir App.Path & "\data\", "spells"
    ChkDir App.Path & "\data\", "convs"
    ChkDir App.Path & "\data\", "quests"

    ' load options, set if they dont exist
    If Not FileExist(App.Path & "\data\options.ini") Then
        Options.Game_Name = "EVB"
        Options.Port = 7001
        Options.MOTD = "Welcome to " & Options.Game_Name & "."
        Options.Website = "http://www.indierising.net/"
        Options.StartMap = 1
        Options.StartX = 5
        Options.StartY = 5
        SaveOptions
    Else
        LoadOptions
    End If
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = Options.Port
    
    ' Init all the player sockets
    Call SetStatus("Initializing player array...")

    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next

    ' Serves as a constructor
    Call ClearGameData
    Call LoadGameData
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call SetStatus("Creating map cache...")
    Call CreateFullMapCache

    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist(App.Path & "data\accounts\charlist.txt") Then
        F = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Output As #F
        Close #F
    End If

    ' Start listening
    frmServer.Socket(0).Listen
    Call UpdateCaption
    
    ' Check how long it took to start up
    EndTime = timeGetTime
    Call SetStatus("Initialization complete. Server loaded in " & EndTime - StartTime & "ms.")
    
    ' reset shutdown value
    isShuttingDown = False
    
    ' Starts the server loop
    ServerLoop
End Sub

Public Sub DestroyServer()
Dim i As Long

    ServerOnline = False
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call ClearGameData
    Call SetStatus("Unloading sockets...")

    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next

    End
End Sub

Public Sub SetStatus(ByVal Status As String)
    Call TextAdd(Status)
    DoEvents
End Sub

Public Sub ClearGameData()
    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTiles
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing resources...")
    Call ClearResources
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing animations...")
    Call ClearAnimations
    Call SetStatus("Clearing conversations...")
    Call ClearConvs
    Call SetStatus("Clearing quests...")
    Call ClearQuests
End Sub

Private Sub LoadGameData()
    Call SetStatus("Loading classes...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading resources...")
    Call LoadResources
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading animations...")
    Call LoadAnimations
    Call SetStatus("Loading conversations...")
    Call LoadConvs
    Call SetStatus("Loading quests...")
    Call LoadQuests
End Sub

Public Sub TextAdd(Msg As String)
    NumLines = NumLines + 1

    If NumLines >= MAX_LINES Then
        frmServer.txtText.Text = vbNullString
        NumLines = 0
    End If

    frmServer.txtText.Text = frmServer.txtText.Text & vbNewLine & Msg
    frmServer.txtText.SelStart = Len(frmServer.txtText.Text)
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean

    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If

End Function

' Checks if there are duplicates of an item
Public Function DuplicatesInInv(ByVal index As Long, ItemNum As Long) As Byte
Dim TempDuplicates As Byte
Dim i As Long

    For i = 1 To MAX_INV
        If Player(index).Inv(i).Num = ItemNum Then
            TempDuplicates = TempDuplicates + 1
        End If
    Next
    
    DuplicatesInInv = TempDuplicates - 1
End Function
