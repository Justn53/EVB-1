Attribute VB_Name = "modDatabase"
Option Explicit

' .INI API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

' For clearing things out of the memory
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' We need this to make sure players with names = name_length can login
Private Const PASS_LEN As Byte = NAME_LENGTH + 1

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal FileName As String) As Boolean
    If LenB(Dir(FileName)) > 0 Then
        FileExist = True
    End If
End Function

Public Sub SaveOptions()
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Game_Name", Options.Game_Name
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Port", STR(Options.Port)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Website", Options.Website
    PutVar App.Path & "\data\options.ini", "OPTIONS", "StartMap", STR(Options.StartMap)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "StartX", STR(Options.StartX)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "StartY", STR(Options.StartY)
End Sub

Public Sub LoadOptions()
    Options.Game_Name = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Game_Name")
    Options.Port = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Port")
    Options.MOTD = GetVar(App.Path & "\data\options.ini", "OPTIONS", "MOTD")
    Options.Website = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Website")
    Options.StartMap = GetVar(App.Path & "\data\options.ini", "OPTIONS", "StartMap")
    Options.StartX = GetVar(App.Path & "\data\options.ini", "OPTIONS", "StartX")
    Options.StartY = GetVar(App.Path & "\data\options.ini", "OPTIONS", "StartY")
End Sub

Public Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
    Dim FileName As String
    Dim IP As String
    Dim F As Long
    Dim i As Long
    FileName = App.Path & "\data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist(App.Path & "\data\banlist.txt") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1
        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If
    Next

    IP = Mid$(IP, 1, i)
    F = FreeFile
    Open FileName For Append As #F
        Print #F, IP & "," & GetPlayerName(BannedByIndex)
    Close #F
    
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
End Sub

Public Sub ServerBanIndex(ByVal BanPlayerIndex As Long)
    Dim FileName As String
    Dim IP As String
    Dim F As Long
    Dim i As Long
    FileName = App.Path & "data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist(App.Path & "\data\banlist.txt") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1
        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If
    Next

    IP = Mid$(IP, 1, i)
    F = FreeFile
    Open FileName For Append As #F
        Print #F, IP & "," & "Server"
    Close #F
    
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & "the Server" & "!", White)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & "The Server" & "!")
End Sub

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal name As String) As Boolean
    Dim FileName As String
    FileName = App.Path & "\data\accounts\" & Trim(name) & ".bin"

    If FileExist(FileName) Then
        AccountExist = True
    End If
End Function

Function PasswordOK(ByVal name As String, ByVal Password As String) As Boolean
Dim FileName As String
Dim RightPassword As String * PASS_LEN
Dim nFileNum As Long

    If AccountExist(name) Then
        FileName = App.Path & "\data\accounts\" & Trim$(name) & ".bin"
        nFileNum = FreeFile
        Open FileName For Binary As #nFileNum
        Get #nFileNum, ACCOUNT_LENGTH, RightPassword
        Close #nFileNum

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Sub AddAccount(ByVal index As Long, ByVal name As String, ByVal Password As String)
    Dim i As Long
    
    ClearPlayer index
    
    Player(index).Login = name
    Player(index).Password = Password

    Call SavePlayer(index)
End Sub

Sub DeleteName(ByVal name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    Call FileCopy(App.Path & "\data\accounts\charlist.txt", App.Path & "\data\accounts\chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.Path & "\data\accounts\chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal index As Long) As Boolean

    If LenB(Trim$(Player(index).name)) > 0 Then
        CharExist = True
    End If

End Function

Sub AddChar(ByVal index As Long, ByVal name As String, ByVal Sex As Byte, ByVal ClassNum As Long, ByVal Sprite As Long)
Dim F As Long
Dim n As Long
Dim spritecheck As Boolean
Dim i As Long

    If LenB(Trim$(Player(index).name)) = 0 Then
        
        spritecheck = False
        
        Player(index).name = name
        Player(index).Sex = Sex
        Player(index).Class = ClassNum
        
        If Player(index).Sex = SEX_MALE Then
            Player(index).Sprite = Class(ClassNum).MaleSprite(Sprite)
        Else
            Player(index).Sprite = Class(ClassNum).FemaleSprite(Sprite)
        End If

        Player(index).Level = 1

        For n = 1 To Stats.Stat_Count - 1
            Player(index).Stat(n) = Class(ClassNum).Stat(n)
        Next n

        Player(index).Dir = DIR_DOWN
        Player(index).Map = Options.StartMap
        Player(index).x = Options.StartX
        Player(index).y = Options.StartY
        Player(index).Dir = DIR_DOWN
        Player(index).Vital(Vitals.HP) = GetPlayerMaxVital(index, Vitals.HP)
        Player(index).Vital(Vitals.MP) = GetPlayerMaxVital(index, Vitals.MP)
        
        For n = 1 To MAX_QUESTS
            For i = 1 To 3
                Player(index).Quest(n).DataAmountLeft(i) = 0
            Next
            
            Player(index).Quest(n).QuestStatus = 0
        Next
        
        ' set starter equipment
        If Class(ClassNum).startItemCount > 0 Then
            For n = 1 To Class(ClassNum).startItemCount
                If Class(ClassNum).StartItem(n) > 0 Then
                    ' item exist?
                    If Len(Trim$(Item(Class(ClassNum).StartItem(n)).name)) > 0 Then
                        Player(index).Inv(n).Num = Class(ClassNum).StartItem(n)
                        Player(index).Inv(n).Value = Class(ClassNum).StartValue(n)
                    End If
                End If
            Next
        End If
        
        ' set start spells
        If Class(ClassNum).startSpellCount > 0 Then
            For n = 1 To Class(ClassNum).startSpellCount
                If Class(ClassNum).StartSpell(n) > 0 Then
                    ' spell exist?
                    If Len(Trim$(Spell(Class(ClassNum).StartSpell(n)).name)) > 0 Then
                        Player(index).Spell(n) = Class(ClassNum).StartSpell(n)
                    End If
                End If
            Next
        End If
        
        ' Append name to file
        F = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Append As #F
        Print #F, name
        Close #F
        Call SavePlayer(index)
        Exit Sub
    End If

End Sub

Function FindChar(ByVal name As String) As Boolean
    Dim F As Long
    Dim s As String
    F = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s

        If Trim$(LCase$(s)) = Trim$(LCase$(name)) Then
            FindChar = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SavePlayer(i)
            Call SaveBank(i)
        End If

    Next

End Sub

Sub SavePlayer(ByVal index As Long)
    Dim FileName As String
    Dim F As Long

    FileName = App.Path & "\data\accounts\" & Trim$(Player(index).Login) & ".bin"
    
    F = FreeFile
    
    Open FileName For Binary As #F
    Put #F, , Player(index)
    Close #F
End Sub

Sub LoadPlayer(ByVal index As Long, ByVal name As String)
    Dim FileName As String
    Dim F As Long
    Call ClearPlayer(index)
    FileName = App.Path & "\data\accounts\" & Trim(name) & ".bin"
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , Player(index)
    Close #F
End Sub

Sub ClearPlayer(ByVal index As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(index)), LenB(TempPlayer(index)))
    Set TempPlayer(index).Buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Player(index)), LenB(Player(index)))
    Player(index).Login = vbNullString
    Player(index).Password = vbNullString
    Player(index).name = vbNullString
    Player(index).Class = 1

    frmServer.lvwInfo.ListItems(index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(3) = vbNullString
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
    Dim FileName As String
    Dim File As String
    FileName = App.Path & "\data\classes.ini"
    Max_Classes = 2

    If Not FileExist(FileName) Then
        File = FreeFile
        Open FileName For Output As File
            Print #File, "[INIT]"
            Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If
End Sub

Sub LoadClasses()
    Dim FileName As String
    Dim i As Long, n As Long
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim startItemCount As Long, startSpellCount As Long
    Dim x As Long

    If CheckClasses Then
        ReDim Class(1 To Max_Classes)
        Call SaveClasses
    Else
        FileName = App.Path & "\data\classes.ini"
        Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes)
    End If

    Call ClearClasses

    For i = 1 To Max_Classes
        Class(i).name = GetVar(FileName, "CLASS" & i, "Name")
        
        ' read string of sprites
        tmpSprite = GetVar(FileName, "CLASS" & i, "MaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).MaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).MaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' read string of sprites
        tmpSprite = GetVar(FileName, "CLASS" & i, "FemaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).FemaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).FemaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' continue
        Class(i).Stat(Stats.Strength) = Val(GetVar(FileName, "CLASS" & i, "Strength"))
        Class(i).Stat(Stats.Endurance) = Val(GetVar(FileName, "CLASS" & i, "Endurance"))
        Class(i).Stat(Stats.Intelligence) = Val(GetVar(FileName, "CLASS" & i, "Intelligence"))
        Class(i).Stat(Stats.Agility) = Val(GetVar(FileName, "CLASS" & i, "Agility"))
        Class(i).Stat(Stats.Willpower) = Val(GetVar(FileName, "CLASS" & i, "Willpower"))
        
        ' how many starting items?
        startItemCount = Val(GetVar(FileName, "CLASS" & i, "StartItemCount"))
        If startItemCount > 0 Then ReDim Class(i).StartItem(1 To startItemCount)
        If startItemCount > 0 Then ReDim Class(i).StartValue(1 To startItemCount)
        
        ' loop for items & values
        Class(i).startItemCount = startItemCount
        If startItemCount >= 1 And startItemCount <= MAX_INV Then
            For x = 1 To startItemCount
                Class(i).StartItem(x) = Val(GetVar(FileName, "CLASS" & i, "StartItem" & x))
                Class(i).StartValue(x) = Val(GetVar(FileName, "CLASS" & i, "StartValue" & x))
            Next
        End If
        
        ' how many starting spells?
        startSpellCount = Val(GetVar(FileName, "CLASS" & i, "StartSpellCount"))
        If startSpellCount > 0 Then ReDim Class(i).StartSpell(1 To startSpellCount)
        
        ' loop for spells
        Class(i).startSpellCount = startSpellCount
        If startSpellCount >= 1 And startSpellCount <= MAX_INV Then
            For x = 1 To startSpellCount
                Class(i).StartSpell(x) = Val(GetVar(FileName, "CLASS" & i, "StartSpell" & x))
            Next
        End If
    Next

End Sub

Sub SaveClasses()
    Dim FileName As String
    Dim i As Long
    Dim x As Long
    
    FileName = App.Path & "\data\classes.ini"

    For i = 1 To Max_Classes
        Call PutVar(FileName, "CLASS" & i, "Name", Trim$(Class(i).name))
        Call PutVar(FileName, "CLASS" & i, "Maleprite", "1")
        Call PutVar(FileName, "CLASS" & i, "Femaleprite", "1")
        Call PutVar(FileName, "CLASS" & i, "Strength", STR(Class(i).Stat(Stats.Strength)))
        Call PutVar(FileName, "CLASS" & i, "Endurance", STR(Class(i).Stat(Stats.Endurance)))
        Call PutVar(FileName, "CLASS" & i, "Intelligence", STR(Class(i).Stat(Stats.Intelligence)))
        Call PutVar(FileName, "CLASS" & i, "Agility", STR(Class(i).Stat(Stats.Agility)))
        Call PutVar(FileName, "CLASS" & i, "Willpower", STR(Class(i).Stat(Stats.Willpower)))
        ' loop for items & values
        For x = 1 To UBound(Class(i).StartItem)
            Call PutVar(FileName, "CLASS" & i, "StartItem" & x, STR(Class(i).StartItem(x)))
            Call PutVar(FileName, "CLASS" & i, "StartValue" & x, STR(Class(i).StartValue(x)))
        Next
        ' loop for spells
        For x = 1 To UBound(Class(i).StartSpell)
            Call PutVar(FileName, "CLASS" & i, "StartSpell" & x, STR(Class(i).StartSpell(x)))
        Next
    Next

End Sub

Function CheckClasses() As Boolean
    Dim FileName As String
    FileName = App.Path & "\data\classes.ini"

    If Not FileExist(FileName) Then
        Call CreateClassesINI
        CheckClasses = True
    End If
End Function

Sub ClearClasses()
    Dim i As Long

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).name = vbNullString
    Next

End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next
End Sub

Sub SaveItem(ByVal ItemNum As Long)
Dim FileName As String
Dim F  As Long

    FileName = App.Path & "\data\items\item" & ItemNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Item(ItemNum)
    Close #F
End Sub

Sub LoadItems()
Dim FileName As String
Dim i As Long
Dim F As Long
    
    Call CheckItems

    For i = 1 To MAX_ITEMS
        FileName = App.Path & "\data\Items\Item" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Item(i)
        Close #F
    Next
End Sub

Sub CheckItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Not FileExist(App.Path & "\data\items\item" & i & ".dat") Then
            Call SaveItem(i)
        End If
    Next
End Sub

Sub ClearItem(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(index)), LenB(Item(index)))
    Item(index).name = vbNullString
    Item(index).Desc = vbNullString
    Item(index).Sound = "None."
    
    ' default first item: gold
    If index = 1 Then
        Item(index).name = "Gold"
        Item(index).Pic = 1
        Item(index).Tradable = True
        Item(index).Type = ITEM_TYPE_CURRENCY
    End If
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next
End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next
End Sub

Sub SaveShop(ByVal shopNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\shops\shop" & shopNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Shop(shopNum)
    Close #F
End Sub

Sub LoadShops()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Call CheckShops

    For i = 1 To MAX_SHOPS
        FileName = App.Path & "\data\shops\shop" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Shop(i)
        Close #F
    Next

End Sub

Sub CheckShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        If Not FileExist(App.Path & "\data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If
    Next
End Sub

Sub ClearShop(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(index)), LenB(Shop(index)))
    Shop(index).name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal spellNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\spells\spells" & spellNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Spell(spellNum)
    Close #F
End Sub

Sub SaveSpells()
    Dim i As Long
    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next

End Sub

Sub LoadSpells()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        FileName = App.Path & "\data\spells\spells" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Spell(i)
        Close #F
    Next

End Sub

Sub CheckSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Not FileExist(App.Path & "\data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If
    Next
End Sub

Sub ClearSpell(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(index)), LenB(Spell(index)))
    Spell(index).name = vbNullString
    Spell(index).LevelReq = 1 'Needs to be 1 for the spell editor
    Spell(index).Desc = vbNullString
    Spell(index).Sound = "None."
End Sub

Sub ClearSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next
End Sub

' **********
' ** NPCs **
' **********
Sub SaveNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next

End Sub

Sub SaveNpc(ByVal NPCNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\npcs\npc" & NPCNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Npc(NPCNum)
    Close #F
End Sub

Sub LoadNpcs()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Call CheckNpcs

    For i = 1 To MAX_NPCS
        FileName = App.Path & "\data\npcs\npc" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Npc(i)
        Close #F
    Next

End Sub

Sub CheckNpcs()
Dim i As Long

    For i = 1 To MAX_NPCS
        If Not FileExist(App.Path & "\data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If
    Next
End Sub

Sub ClearNpc(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Npc(index)), LenB(Npc(index)))
    Npc(index).name = vbNullString
    Npc(index).AttackSay = vbNullString
    Npc(index).Sound = "None."
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\resources\resource" & ResourceNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Resource(ResourceNum)
    Close #F
End Sub

Sub LoadResources()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
    Call CheckResources

    For i = 1 To MAX_RESOURCES
        FileName = App.Path & "\data\resources\resource" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Resource(i)
        Close #F
    Next

End Sub

Sub CheckResources()
Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Not FileExist(App.Path & "\data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If
    Next
End Sub

Sub ClearResource(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(index)), LenB(Resource(index)))
    Resource(index).name = vbNullString
    Resource(index).SuccessMessage = vbNullString
    Resource(index).EmptyMessage = vbNullString
    Resource(index).Sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

' **********
' ** animations **
' **********
Sub SaveAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next

End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\animations\animation" & AnimationNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Animation(AnimationNum)
    Close #F
End Sub

Sub LoadAnimations()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        FileName = App.Path & "\data\animations\animation" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Animation(i)
        Close #F
    Next

End Sub

Sub CheckAnimations()
Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        If Not FileExist(App.Path & "\data\animations\animation" & i & ".dat") Then
            Call SaveAnimation(i)
        End If
    Next
End Sub

Sub ClearAnimation(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(index)), LenB(Animation(index)))
    Animation(index).name = vbNullString
    Animation(index).Sound = "None."
End Sub

Sub ClearAnimations()
Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal MapNum As Long)
    Dim FileName As String
    Dim F As Long
    Dim x As Long
    Dim y As Long
    FileName = App.Path & "\data\maps\map" & MapNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
    Put #F, , Map(MapNum).name
    Put #F, , Map(MapNum).Music
    Put #F, , Map(MapNum).Revision
    Put #F, , Map(MapNum).Moral
    Put #F, , Map(MapNum).Up
    Put #F, , Map(MapNum).Down
    Put #F, , Map(MapNum).Left
    Put #F, , Map(MapNum).Right
    Put #F, , Map(MapNum).BootMap
    Put #F, , Map(MapNum).BootX
    Put #F, , Map(MapNum).BootY
    Put #F, , Map(MapNum).MaxX
    Put #F, , Map(MapNum).MaxY

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            Put #F, , Map(MapNum).Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #F, , Map(MapNum).Npc(x)
    Next
    Close #F
    
    DoEvents
End Sub

Sub SaveMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next

End Sub

Sub LoadMaps()
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Dim x As Long
    Dim y As Long
    Call CheckMaps

    For i = 1 To MAX_MAPS
        FileName = App.Path & "\data\maps\map" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Map(i).name
        Get #F, , Map(i).Music
        Get #F, , Map(i).Revision
        Get #F, , Map(i).Moral
        Get #F, , Map(i).Up
        Get #F, , Map(i).Down
        Get #F, , Map(i).Left
        Get #F, , Map(i).Right
        Get #F, , Map(i).BootMap
        Get #F, , Map(i).BootX
        Get #F, , Map(i).BootY
        Get #F, , Map(i).MaxX
        Get #F, , Map(i).MaxY
        ' have to set the tile()
        ReDim Map(i).Tile(0 To Map(i).MaxX, 0 To Map(i).MaxY)

        For x = 0 To Map(i).MaxX
            For y = 0 To Map(i).MaxY
                Get #F, , Map(i).Tile(x, y)
            Next
        Next

        For x = 1 To MAX_MAP_NPCS
            Get #F, , Map(i).Npc(x)
            MapNpc(i).Npc(x).Num = Map(i).Npc(x)
        Next

        Close #F
        
        ClearTempTile i
        CacheResources i
        DoEvents
    Next
End Sub

Sub CheckMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        If Not FileExist(App.Path & "\data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If
    Next
End Sub

Sub ClearMapItem(ByVal index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, index)), LenB(MapItem(MapNum, index)))
    MapItem(MapNum, index).playerName = vbNullString
End Sub

Sub ClearMapItems()
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next
End Sub

Sub ClearMapNpc(ByVal index As Long, ByVal MapNum As Long)
    ReDim MapNpc(MapNum).Npc(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum).Npc(index)), LenB(MapNpc(MapNum).Npc(index)))
End Sub

Sub ClearMapNpcs()
Dim x As Long
Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next
End Sub

Sub ClearMap(ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    Map(MapNum).name = vbNullString
    Map(MapNum).MaxX = MAX_MAPX
    Map(MapNum).MaxY = MAX_MAPY
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = 0
    ' Reset the map cache array for this map.
    MapCache(MapNum).Data = vbNullString
End Sub

Sub ClearMaps()
Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next
End Sub

' *********
' * Convs *
' *********
Public Sub SaveConvs()
Dim i As Long

    For i = 1 To MAX_CONVS
        Call SaveConv(i)
    Next
End Sub

Public Sub SaveConv(ByVal ConvNum As Long)
Dim FileName As String
Dim F As Long

    FileName = App.Path & "\data\convs\conv" & ConvNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Conv(ConvNum)
    Close #F
End Sub

Public Sub LoadConvs()
Dim FileName As String
Dim i As Long
Dim F As Long

    Call CheckConvs

    For i = 1 To MAX_CONVS
        FileName = App.Path & "\data\convs\conv" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Conv(i)
        Close #F
    Next
End Sub

Public Sub CheckConvs()
Dim i As Long

    For i = 1 To MAX_CONVS
        If Not FileExist(App.Path & "\data\convs\conv" & i & ".dat") Then
            Call SaveConv(i)
        End If
    Next
End Sub

Public Sub ClearConv(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Conv(index)), LenB(Conv(index)))
    Conv(index).ChatCount = 1
End Sub

Public Sub ClearConvs()
Dim i As Long

    For i = 1 To MAX_CONVS
        Call ClearConv(i)
    Next
End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).name)
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            With Class(ClassNum)
                GetClassMaxVital = 100 + (.Stat(Endurance) * 5) + 2
            End With
        Case MP
            With Class(ClassNum)
                GetClassMaxVital = 30 + (.Stat(Intelligence) * 10) + 2
            End With
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function

Sub SaveBank(ByVal index As Long)
    Dim FileName As String
    Dim F As Long
    
    FileName = App.Path & "\data\banks\" & Trim$(Player(index).Login) & ".bin"
    
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Bank(index)
    Close #F
End Sub

Public Sub LoadBank(ByVal index As Long, ByVal name As String)
    Dim FileName As String
    Dim F As Long

    Call ClearBank(index)

    FileName = App.Path & "\data\banks\" & Trim$(name) & ".bin"
    
    If Not FileExist(FileName) Then
        Call SaveBank(index)
        Exit Sub
    End If

    F = FreeFile
    Open FileName For Binary As #F
        Get #F, , Bank(index)
    Close #F
End Sub

Sub ClearBank(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Bank(index)), LenB(Bank(index)))
End Sub

Sub ClearParty(ByVal partyNum As Long)
    Call ZeroMemory(ByVal VarPtr(Party(partyNum)), LenB(Party(partyNum)))
End Sub

' *********
' * Quests *
' *********
Public Sub SaveQuests()
Dim i As Long

    For i = 1 To MAX_QUESTS
        Call SaveQuest(i)
    Next
End Sub

Public Sub SaveQuest(ByVal QuestNum As Long)
Dim FileName As String
Dim F As Long

    FileName = App.Path & "\data\quests\quest" & QuestNum & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Quest(QuestNum)
    Close #F
End Sub

Public Sub LoadQuests()
Dim FileName As String
Dim i As Long
Dim F As Long

    Call CheckQuests

    For i = 1 To MAX_QUESTS
        FileName = App.Path & "\data\quests\quest" & i & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, , Quest(i)
        Close #F
    Next
End Sub

Public Sub CheckQuests()
Dim i As Long

    For i = 1 To MAX_QUESTS
        If Not FileExist(App.Path & "\data\quests\quest" & i & ".dat") Then
            Call SaveQuest(i)
        End If
    Next
End Sub

Public Sub ClearQuest(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Quest(index)), LenB(Quest(index)))
End Sub

Public Sub ClearQuests()
Dim i As Long

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next
End Sub

