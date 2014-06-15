Attribute VB_Name = "modDirectDraw7"
Option Explicit

' DirectDraw7 Object
Public DD As DirectDraw7

' Clipper object
Public DD_Clip As DirectDrawClipper

' Primary surface
Public DDS_Primary As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

' Back buffer
Public DDS_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

' Used for pre-rendering
Public DDS_Map As DirectDrawSurface7
Public DDSD_Map As DDSURFACEDESC2

' Graphics buffers
Public DDS_Item() As DirectDrawSurface7 ' arrays
Public DDS_Character() As DirectDrawSurface7
Public DDS_Paperdoll() As DirectDrawSurface7
Public DDS_Tileset() As DirectDrawSurface7
Public DDS_Resource() As DirectDrawSurface7
Public DDS_Animation() As DirectDrawSurface7
Public DDS_SpellIcon() As DirectDrawSurface7
Public DDS_Face() As DirectDrawSurface7

Public DDS_Blood As DirectDrawSurface7 ' singles
Public DDS_Misc As DirectDrawSurface7
Public DDS_Direction As DirectDrawSurface7
Public DDS_Target As DirectDrawSurface7
Public DDS_Bars As DirectDrawSurface7

' Descriptions
Public DDSD_Temp As DDSURFACEDESC2 ' arrays
Public DDSD_Item() As DDSURFACEDESC2
Public DDSD_Character() As DDSURFACEDESC2
Public DDSD_Paperdoll() As DDSURFACEDESC2
Public DDSD_Tileset() As DDSURFACEDESC2
Public DDSD_Resource() As DDSURFACEDESC2
Public DDSD_Animation() As DDSURFACEDESC2
Public DDSD_SpellIcon() As DDSURFACEDESC2
Public DDSD_Face() As DDSURFACEDESC2

Public DDSD_Blood As DDSURFACEDESC2 ' singles
Public DDSD_Misc As DDSURFACEDESC2
Public DDSD_Direction As DDSURFACEDESC2
Public DDSD_Target As DDSURFACEDESC2
Public DDSD_Bars As DDSURFACEDESC2

' Timers
Public Const SurfaceTimerMax As Long = 10000
Public CharacterTimer() As Long
Public PaperdollTimer() As Long
Public ItemTimer() As Long
Public ResourceTimer() As Long
Public AnimationTimer() As Long
Public SpellIconTimer() As Long
Public FaceTimer() As Long

' Number of graphic files
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public NumItems As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long
Public NumFaces As Long

' ********************
' ** Initialization **
' ********************
Public Function InitDirectDraw() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Clear DD7
    Call DestroyDirectDraw
    
    ' Init Direct Draw
    Set DD = DX7.DirectDrawCreate(vbNullString)
    
    ' Windowed
    DD.SetCooperativeLevel frmMain.hWnd, DDSCL_NORMAL

    ' Init type and set the primary surface
    With DDSD_Primary
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
        .lBackBufferCount = 1
    End With
    Set DDS_Primary = DD.CreateSurface(DDSD_Primary)
    
    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)
    
    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMain.picScreen.hWnd
    
    ' Have the blits to the screen clipped to the picture box
    DDS_Primary.SetClipper DD_Clip
    
    ' Initialise the surfaces
    InitSurfaces
    
    ' We're done
    InitDirectDraw = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InitDirectDraw", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub InitSurfaces()
Dim rec As DXVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' DirectDraw Surface memory management setting
    DDSD_Temp.lFlags = DDSD_CAPS
    DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    
    ' clear out everything for re-init
    Set DDS_BackBuffer = Nothing

    ' Initialize back buffer
    With DDSD_BackBuffer
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (MAX_MAPX + 3) * PIC_X
        .lHeight = (MAX_MAPY + 3) * PIC_Y
    End With
    Set DDS_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
    
    ' load persistent surfaces
    If FileExist(App.Path & "\data files\graphics\direction.bmp") Then Call InitDDSurf("direction", DDSD_Direction, DDS_Direction)
    If FileExist(App.Path & "\data files\graphics\target.bmp") Then Call InitDDSurf("target", DDSD_Target, DDS_Target)
    If FileExist(App.Path & "\data files\graphics\misc.bmp") Then Call InitDDSurf("misc", DDSD_Misc, DDS_Misc)
    If FileExist(App.Path & "\data files\graphics\blood.bmp") Then Call InitDDSurf("blood", DDSD_Blood, DDS_Blood)
    If FileExist(App.Path & "\data files\graphics\bars.bmp") Then Call InitDDSurf("bars", DDSD_Bars, DDS_Bars)
    
    ' count the blood sprites
    BloodCount = DDSD_Blood.lWidth / 32
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitSurfaces", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' This sub gets the mask color from the surface loaded from a bitmap image
Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal X As Long, ByVal Y As Long)
Dim TmpR As RECT
Dim TmpDDSD As DDSURFACEDESC2
Dim TmpColorKey As DDCOLORKEY

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With TmpR
        .Left = X
        .top = Y
        .Right = X
        .Bottom = Y
    End With

    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

    With TmpColorKey
        .Low = TheSurface.GetLockedPixel(X, Y)
        .High = .Low
    End With

    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey
    TheSurface.Unlock TmpR
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetMaskColorFromPixel", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Initializing a surface, using a bitmap
Public Sub InitDDSurf(FileName As String, ByRef SurfDesc As DDSURFACEDESC2, ByRef Surf As DirectDrawSurface7)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Set path
    FileName = App.Path & GFX_PATH & FileName & GFX_EXT

    ' Destroy surface if it exist
    If Not Surf Is Nothing Then
        Set Surf = Nothing
        Call ZeroMemory(ByVal VarPtr(SurfDesc), LenB(SurfDesc))
    End If

    ' set flags
    SurfDesc.lFlags = DDSD_CAPS
    SurfDesc.ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
    
    ' init object
    Set Surf = DD.CreateSurfaceFromFile(FileName, SurfDesc)
    
    ' Set mask
    Call SetMaskColorFromPixel(Surf, 0, 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitDDSurf", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function CheckSurfaces() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if we need to restore surfaces
    If Not DD.TestCooperativeLevel = DD_OK Then
        CheckSurfaces = False
    Else
        CheckSurfaces = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "CheckSurfaces", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub ReInitDD()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call InitDirectDraw
    LoadTilesets
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ReInitDD", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyDirectDraw()
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Unload DirectDraw
    Set DDS_Misc = Nothing
    
    For i = 1 To NumTileSets
        Set DDS_Tileset(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i))
    Next

    For i = 1 To NumItems
        Set DDS_Item(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Item(i)), LenB(DDSD_Item(i))
    Next

    For i = 1 To NumCharacters
        Set DDS_Character(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Character(i)), LenB(DDSD_Character(i))
    Next
    
    For i = 1 To NumPaperdolls
        Set DDS_Paperdoll(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Paperdoll(i)), LenB(DDSD_Paperdoll(i))
    Next
    
    For i = 1 To NumResources
        Set DDS_Resource(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Resource(i)), LenB(DDSD_Resource(i))
    Next
    
    For i = 1 To NumAnimations
        Set DDS_Animation(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Animation(i)), LenB(DDSD_Animation(i))
    Next
    
    For i = 1 To NumSpellIcons
        Set DDS_SpellIcon(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_SpellIcon(i)), LenB(DDSD_SpellIcon(i))
    Next
    
    For i = 1 To NumFaces
        Set DDS_Face(i) = Nothing
        ZeroMemory ByVal VarPtr(DDSD_Face(i)), LenB(DDSD_Face(i))
    Next
    
    Set DDS_Blood = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Blood), LenB(DDSD_Blood)

    Set DDS_Direction = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Direction), LenB(DDSD_Direction)
    
    Set DDS_Target = Nothing
    ZeroMemory ByVal VarPtr(DDSD_Target), LenB(DDSD_Target)

    Set DDS_BackBuffer = Nothing
    Set DDS_Primary = Nothing
    Set DD_Clip = Nothing
    Set DD = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DestroyDirectDraw", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **************
' ** Blitting **
' **************
Public Sub Engine_BltFast(ByVal dx As Long, ByVal dy As Long, ByRef ddS As DirectDrawSurface7, srcRECT As RECT, trans As CONST_DDBLTFASTFLAGS)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not ddS Is Nothing Then
        Call DDS_BackBuffer.BltFast(dx, dy, ddS, srcRECT, trans)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Engine_BltFast", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function Engine_BltToDC(ByRef Surface As DirectDrawSurface7, sRECT As DXVBLib.RECT, dRECT As DXVBLib.RECT, ByRef picBox As VB.PictureBox, Optional Clear As Boolean = True) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Clear Then
        picBox.Cls
    End If

    Call Surface.BltToDC(picBox.hdc, sRECT, dRECT)
    picBox.Refresh
    Engine_BltToDC = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "Engine_BltToDC", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub BltDirection(ByVal X As Long, ByVal Y As Long)
Dim rec As DXVBLib.RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' render grid
    rec.top = 24
    rec.Left = 0
    rec.Right = rec.Left + 32
    rec.Bottom = rec.top + 32
    Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Direction, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' render dir blobs
    For i = 1 To 4
        rec.Left = (i - 1) * 8
        rec.Right = rec.Left + 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Map.Tile(X, Y).DirBlock, CByte(i)) Then
            rec.top = 8
        Else
            rec.top = 16
        End If
        rec.Bottom = rec.top + 8
        'render!
        Call Engine_BltFast(ConvertMapX(X * PIC_X) + DirArrowX(i), ConvertMapY(Y * PIC_Y) + DirArrowY(i), DDS_Direction, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltDirection", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltTarget(ByVal X As Long, ByVal Y As Long)
Dim sRECT As DXVBLib.RECT
Dim width As Long, height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DDS_Target Is Nothing Then Exit Sub
    
    width = DDSD_Target.lWidth / 2
    height = DDSD_Target.lHeight

    With sRECT
        .top = 0
        .Bottom = height
        .Left = 0
        .Right = width
    End With
    
    X = X - ((width - 32) / 2)
    Y = Y - (height / 2)
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' clipping
    If Y < 0 Then
        With sRECT
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + height - DDSD_BackBuffer.lHeight)
    End If

    If X + width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + width - DDSD_BackBuffer.lWidth)
    End If
    
    Call Engine_BltFast(X, Y, DDS_Target, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTarget", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltHover(ByVal tType As Long, ByVal target As Long, ByVal X As Long, ByVal Y As Long)
Dim sRECT As DXVBLib.RECT
Dim width As Long, height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DDS_Target Is Nothing Then Exit Sub
    
    width = DDSD_Target.lWidth / 2
    height = DDSD_Target.lHeight

    With sRECT
        .top = 0
        .Bottom = height
        .Left = width
        .Right = .Left + width
    End With
    
    X = X - ((width - 32) / 2)
    Y = Y - (height / 2)

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' clipping
    If Y < 0 Then
        With sRECT
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + height - DDSD_BackBuffer.lHeight)
    End If

    If X + width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + width - DDSD_BackBuffer.lWidth)
    End If
    
    Call Engine_BltFast(X, Y, DDS_Target, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHover", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isAutotileMatch(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    isAutotileMatch = False
    ' end of the map, pretend there's a tile there :3
    If x1 < 0 Or x1 > Map.MaxX Or x2 < 0 Or x2 > Map.MaxX Or y1 < 0 Or y1 > Map.MaxY Or y2 < 0 Or y2 > Map.MaxY Then
        isAutotileMatch = True
    Else
        If Map.Tile(x1, y1).Layer(MapLayer.Autotile).X = Map.Tile(x2, y2).Layer(MapLayer.Autotile).X Then
            If Map.Tile(x1, y1).Layer(MapLayer.Autotile).Y = Map.Tile(x2, y2).Layer(MapLayer.Autotile).Y Then
                If Map.Tile(x1, y1).Layer(MapLayer.Autotile).Tileset = Map.Tile(x2, y2).Layer(MapLayer.Autotile).Tileset Then
                    isAutotileMatch = True
                End If
            End If
        End If
    End If
End Function
Public Sub BltAutotile(ByVal X As Long, ByVal Y As Long, Optional ByVal ScreenShot As Boolean = False)
Dim rec As DXVBLib.RECT
Dim AutoSet As Long
Dim bTop As Long
Dim bLeft As Long
Dim TileSetWidth As Long
Dim TileSetHeight As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    AutoSet = Map.Tile(X, Y).Layer(MapLayer.Autotile).Tileset
    
    ' Check for subscript out of range
    If AutoSet < 1 Or AutoSet > NumTileSets Then Exit Sub
    
    bTop = Map.Tile(X, Y).Layer(MapLayer.Autotile).Y * 32
    bLeft = Map.Tile(X, Y).Layer(MapLayer.Autotile).X * 32
    TileSetWidth = DDSD_Tileset(AutoSet).lWidth
    TileSetHeight = DDSD_Tileset(AutoSet).lHeight
    
    ' Check the height of the autoset to make sure it's not out of bounds
    If bTop + 32 > TileSetHeight Or bTop + 64 > TileSetHeight Or bTop + 96 > TileSetHeight Or bTop + 128 > TileSetHeight Then Exit Sub
    
    ' Check the width of the autoset to make sure it's not out of bounds
    If bLeft + 32 > TileSetWidth Or bLeft + 64 > TileSetWidth Or bLeft + 96 > TileSetWidth Then Exit Sub
    
    ' Animate if width is 384
    If DDSD_Tileset(AutoSet).lWidth = 384 Then bLeft = autoAnim * 96
    
    ' ############
    ' # top left #
    ' ############
    With rec
        .top = bTop + 32
        .Bottom = .top + 16
        .Left = bLeft + 0
        .Right = .Left + 16
    End With
    ' 1-dir horizontal join
    If isAutotileMatch(X, Y, X - 1, Y) Then
        With rec
            .top = bTop + 32
            .Bottom = .top + 16
            .Left = bLeft + 64
            .Right = .Left + 16
        End With
        ' 2-dir horizontal join
        If isAutotileMatch(X, Y, X + 1, Y) Then
            With rec
                .top = bTop + 32
                .Bottom = .top + 16
                .Left = bLeft + 32
                .Right = .Left + 16
            End With
        End If
    End If
    ' 1-dir vertical join
    If isAutotileMatch(X, Y, X, Y - 1) Then
        With rec
            .top = bTop + 96
            .Bottom = .top + 16
            .Left = bLeft + 0
            .Right = .Left + 16
        End With
        ' 2-dir vertical join
        If isAutotileMatch(X, Y, X, Y + 1) Then
            With rec
                .top = bTop + 64
                .Bottom = .top + 16
                .Left = bLeft + 0
                .Right = .Left + 16
            End With
        End If
    End If
    ' 1-dir corner join
    If isAutotileMatch(X, Y, X - 1, Y) And isAutotileMatch(X, Y, X, Y - 1) Then
        With rec
            .top = bTop + 0
            .Bottom = .top + 16
            .Left = bLeft + 64
            .Right = .Left + 16
        End With
        ' 2-dir corner join
        If isAutotileMatch(X, Y, X - 1, Y - 1) Then
            With rec
                .top = bTop + 96
                .Bottom = .top + 16
                .Left = bLeft + 64
                .Right = .Left + 16
            End With
        End If
    End If
    
    ' Surrounded to the left or top
    If (isAutotileMatch(X, Y, X - 1, Y) And isAutotileMatch(X, Y, X, Y - 1) And _
        isAutotileMatch(X, Y, X, Y + 1) And isAutotileMatch(X, Y, X - 1, Y - 1) And _
        isAutotileMatch(X, Y, X - 1, Y + 1)) Or (isAutotileMatch(X, Y, X - 1, Y) And _
        isAutotileMatch(X, Y, X + 1, Y) And isAutotileMatch(X, Y, X, Y - 1) And _
        isAutotileMatch(X, Y, X + 1, Y - 1) And isAutotileMatch(X, Y, X - 1, Y - 1)) Then
        
        With rec
            .top = bTop + 64
            .Bottom = .top + 16
            .Left = bLeft + 32
            .Right = .Left + 16
        End With
    End If
    
    If ScreenShot Then
        Call DDS_Map.BltFast(X * PIC_X, Y * PIC_Y, DDS_Tileset(AutoSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        Call Engine_BltFast(ConvertMapX((X * PIC_X)), ConvertMapY((Y * PIC_Y)), DDS_Tileset(AutoSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If

    '#############
    '# top right #
    '#############
    With rec
        .top = bTop + 32
        .Bottom = .top + 16
        .Left = bLeft + 80
        .Right = .Left + 16
    End With
    ' 1-dir horizontal join
    If isAutotileMatch(X, Y, X + 1, Y) Then
        With rec
            .top = bTop + 32
            .Bottom = .top + 16
            .Left = bLeft + 16
            .Right = .Left + 16
        End With
        ' 2-dir horizontal join
        If isAutotileMatch(X, Y, X - 1, Y) Then
            With rec
                .top = bTop + 32
                .Bottom = .top + 16
                .Left = bLeft + 48
                .Right = .Left + 16
            End With
        End If
    End If
    ' 1-dir vertical join
    If isAutotileMatch(X, Y, X, Y - 1) Then
        With rec
            .top = bTop + 96
            .Bottom = .top + 16
            .Left = bLeft + 80
            .Right = .Left + 16
        End With
        ' 2-dir vertical join
        If isAutotileMatch(X, Y, X, Y + 1) Then
            With rec
                .top = bTop + 64
                .Bottom = .top + 16
                .Left = bLeft + 80
                .Right = .Left + 16
            End With
        End If
    End If
    ' 1-dir corner join
    If isAutotileMatch(X, Y, X + 1, Y) And isAutotileMatch(X, Y, X, Y - 1) Then
        With rec
            .top = bTop + 0
            .Bottom = .top + 16
            .Left = bLeft + 80
            .Right = .Left + 16
        End With
        ' 2-dir corner join
        If isAutotileMatch(X, Y, X + 1, Y - 1) Then
            With rec
                .top = bTop + 96
                .Bottom = .top + 16
                .Left = bLeft + 16
                .Right = .Left + 16
            End With
        End If
    End If
    
    ' Surrounded to the right or top
    If (isAutotileMatch(X, Y, X + 1, Y) And isAutotileMatch(X, Y, X, Y - 1) And _
        isAutotileMatch(X, Y, X, Y + 1) And isAutotileMatch(X, Y, X + 1, Y - 1) And _
        isAutotileMatch(X, Y, X + 1, Y + 1)) Or (isAutotileMatch(X, Y, X - 1, Y) And _
        isAutotileMatch(X, Y, X + 1, Y) And isAutotileMatch(X, Y, X, Y - 1) And _
        isAutotileMatch(X, Y, X + 1, Y - 1) And isAutotileMatch(X, Y, X - 1, Y - 1)) Then
        
        With rec
            .top = bTop + 64
            .Bottom = .top + 16
            .Left = bLeft + 48
            .Right = .Left + 16
        End With
    End If
    
    If ScreenShot Then
        Call DDS_Map.BltFast((X * PIC_X) + 16, (Y * PIC_Y), DDS_Tileset(AutoSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        Call Engine_BltFast(ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y)), DDS_Tileset(AutoSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
    '###############
    '# bottom left #
    '###############
    With rec
        .top = bTop + 112
        .Bottom = .top + 16
        .Left = bLeft + 0
        .Right = .Left + 16
    End With
    ' 1-dir horizontal join
    If isAutotileMatch(X, Y, X - 1, Y) Then
        With rec
            .top = bTop + 112
            .Bottom = .top + 16
            .Left = bLeft + 64
            .Right = .Left + 16
        End With
        ' 2-dir horizontal join
        If isAutotileMatch(X, Y, X + 1, Y) Then
            With rec
                .top = bTop + 112
                .Bottom = .top + 16
                .Left = bLeft + 32
                .Right = .Left + 16
            End With
        End If
    End If
    ' 1-dir vertical join
    If isAutotileMatch(X, Y, X, Y + 1) Then
        With rec
            .top = bTop + 48
            .Bottom = .top + 16
            .Left = bLeft + 0
            .Right = .Left + 16
        End With
        ' 2-dir vertical join
        If isAutotileMatch(X, Y, X, Y - 1) Then
            With rec
                .top = bTop + 80
                .Bottom = .top + 16
                .Left = bLeft + 0
                .Right = .Left + 16
            End With
        End If
    End If
    ' 1-dir corner join
    If isAutotileMatch(X, Y, X - 1, Y) And isAutotileMatch(X, Y, X, Y + 1) Then
        With rec
            .top = bTop + 16
            .Bottom = .top + 16
            .Left = bLeft + 64
            .Right = .Left + 16
        End With
        ' 2-dir corner join
        If isAutotileMatch(X, Y, X - 1, Y + 1) Then
            With rec
                .top = bTop + 48
                .Bottom = .top + 16
                .Left = bLeft + 64
                .Right = .Left + 16
            End With
        End If
    End If
    
    ' Surrounded to the left or top
    If (isAutotileMatch(X, Y, X - 1, Y) And isAutotileMatch(X, Y, X, Y - 1) And _
        isAutotileMatch(X, Y, X, Y + 1) And isAutotileMatch(X, Y, X - 1, Y - 1) And _
        isAutotileMatch(X, Y, X - 1, Y + 1)) Or (isAutotileMatch(X, Y, X - 1, Y) And _
        isAutotileMatch(X, Y, X + 1, Y) And isAutotileMatch(X, Y, X, Y + 1) And _
        isAutotileMatch(X, Y, X + 1, Y + 1) And isAutotileMatch(X, Y, X - 1, Y + 1)) Then
        
        With rec
            .top = bTop + 80
            .Bottom = .top + 16
            .Left = bLeft + 64
            .Right = .Left + 16
        End With
    End If
    
    If ScreenShot Then
        Call DDS_Map.BltFast(X * PIC_X, (Y * PIC_Y) + 16, DDS_Tileset(AutoSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        Call Engine_BltFast(ConvertMapX((X * PIC_X)), ConvertMapY((Y * PIC_Y) + 16), DDS_Tileset(AutoSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
    '################
    '# bottom right #
    '################
    With rec
        .top = bTop + 112
        .Bottom = .top + 16
        .Left = bLeft + 80
        .Right = .Left + 16
    End With
    ' 1-dir horizontal join
    If isAutotileMatch(X, Y, X + 1, Y) Then
        With rec
            .top = bTop + 112
            .Bottom = .top + 16
            .Left = bLeft + 16
            .Right = .Left + 16
        End With
        ' 2-dir horizontal join
        If isAutotileMatch(X, Y, X - 1, Y) Then
            With rec
                .top = bTop + 112
                .Bottom = .top + 16
                .Left = bLeft + 48
                .Right = .Left + 16
            End With
        End If
    End If
    ' 1-dir vertical join
    If isAutotileMatch(X, Y, X, Y + 1) Then
        With rec
            .top = bTop + 48
            .Bottom = .top + 16
            .Left = bLeft + 80
            .Right = .Left + 16
        End With
        ' 2-dir vertical join
        If isAutotileMatch(X, Y, X, Y - 1) Then
            With rec
                .top = bTop + 80
                .Bottom = .top + 16
                .Left = bLeft + 80
                .Right = .Left + 16
            End With
        End If
    End If
    ' 1-dir corner join
    If isAutotileMatch(X, Y, X + 1, Y) And isAutotileMatch(X, Y, X, Y + 1) Then
        With rec
            .top = bTop + 16
            .Bottom = .top + 16
            .Left = bLeft + 80
            .Right = .Left + 16
        End With
        ' 2-dir corner join
        If isAutotileMatch(X, Y, X + 1, Y + 1) Then
            With rec
                .top = bTop + 48
                .Bottom = .top + 16
                .Left = bLeft + 16
                .Right = .Left + 16
            End With
        End If
    End If
    
    ' Surrounded to the left or top
    If (isAutotileMatch(X, Y, X + 1, Y) And isAutotileMatch(X, Y, X, Y - 1) And _
        isAutotileMatch(X, Y, X, Y + 1) And isAutotileMatch(X, Y, X + 1, Y - 1) And _
        isAutotileMatch(X, Y, X + 1, Y + 1)) Or (isAutotileMatch(X, Y, X - 1, Y) And _
        isAutotileMatch(X, Y, X + 1, Y) And isAutotileMatch(X, Y, X, Y + 1) And _
        isAutotileMatch(X, Y, X + 1, Y + 1) And isAutotileMatch(X, Y, X - 1, Y + 1)) Then
        
        With rec
            .top = bTop + 80
            .Bottom = .top + 16
            .Left = bLeft + 48
            .Right = .Left + 16
        End With
    End If
    
    If ScreenShot Then
        Call DDS_Map.BltFast((X * PIC_X) + 16, (Y * PIC_Y) + 16, DDS_Tileset(AutoSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        Call Engine_BltFast(ConvertMapX((X * PIC_X) + 16), ConvertMapY((Y * PIC_Y) + 16), DDS_Tileset(AutoSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
' Error handler
errorhandler:
    Exit Sub
    HandleError "BltAutotile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltMapTile(ByVal X As Long, ByVal Y As Long)
Dim rec As DXVBLib.RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Ground To MapLayer.Mask2
            ' skip tile?
            If i = MapLayer.Autotile Then
                BltAutotile X, Y
            Else
                If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                    ' sort out rec
                    rec.top = .Layer(i).Y * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    rec.Left = .Layer(i).X * PIC_X
                    rec.Right = rec.Left + PIC_X
                    ' render
                    Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If
        Next
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltMapTile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltMapFringeTile(ByVal X As Long, ByVal Y As Long)
Dim rec As DXVBLib.RECT
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(X, Y)
        For i = MapLayer.Fringe To MapLayer.Fringe2
            ' skip tile if tileset isn't set
            If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                ' sort out rec
                rec.top = .Layer(i).Y * PIC_Y
                rec.Bottom = rec.top + PIC_Y
                rec.Left = .Layer(i).X * PIC_X
                rec.Right = rec.Left + PIC_X
                ' render
                Call Engine_BltFast(ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        Next
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltMapFringeTile", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltBlood(ByVal Index As Long)
Dim rec As DXVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Blood(Index)
        ' check if we should be seeing it
        If .Timer + 20000 < timeGetTime Then Exit Sub
        
        rec.top = 0
        rec.Bottom = PIC_Y
        rec.Left = (.Sprite - 1) * PIC_X
        rec.Right = rec.Left + PIC_X
        
        Engine_BltFast ConvertMapX(.X * PIC_X), ConvertMapY(.Y * PIC_Y), DDS_Blood, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBlood", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltAnimation(ByVal Index As Long, ByVal Layer As Long)
Dim Sprite As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT
Dim i As Long
Dim width As Long, height As Long
Dim looptime As Long
Dim FrameCount As Long
Dim X As Long, Y As Long
Dim lockindex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer)
    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    AnimationTimer(Sprite) = timeGetTime + SurfaceTimerMax
    
    If DDS_Animation(Sprite) Is Nothing Then
        Call InitDDSurf("animations\" & Sprite, DDSD_Animation(Sprite), DDS_Animation(Sprite))
    End If
    
    ' total width divided by frame count
    width = DDSD_Animation(Sprite).lWidth / FrameCount
    height = DDSD_Animation(Sprite).lHeight
    
    sRECT.top = 0
    sRECT.Bottom = height
    sRECT.Left = (AnimInstance(Index).FrameIndex(Layer) - 1) * width
    sRECT.Right = sRECT.Left + width
    
    ' change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    X = (GetPlayerX(lockindex) * PIC_X) + 16 - (width / 2) + Player(lockindex).XOffset
                    Y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (height / 2) + Player(lockindex).YOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    X = (MapNpc(lockindex).X * PIC_X) + 16 - (width / 2) + MapNpc(lockindex).XOffset
                    Y = (MapNpc(lockindex).Y * PIC_Y) + 16 - (height / 2) + MapNpc(lockindex).YOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        X = (AnimInstance(Index).X * 32) + 16 - (width / 2)
        Y = (AnimInstance(Index).Y * 32) + 16 - (height / 2)
    End If
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    ' Clip to screen
    If Y < 0 Then
        With sRECT
            .top = .top - Y
        End With

        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With

        X = 0
    End If

    If Y + height > DDSD_BackBuffer.lHeight Then
        sRECT.Bottom = sRECT.Bottom - (Y + height - DDSD_BackBuffer.lHeight)
    End If

    If X + width > DDSD_BackBuffer.lWidth Then
        sRECT.Right = sRECT.Right - (X + width - DDSD_BackBuffer.lWidth)
    End If
    
    Call Engine_BltFast(X, Y, DDS_Animation(Sprite), sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltAnimation", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltItem(ByVal itemNum As Long)
Dim PicNum As Long
Dim rec As DXVBLib.RECT
Dim MaxFrames As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if it's not us then don't render
    If MapItem(itemNum).playerName <> vbNullString Then
        If MapItem(itemNum).playerName <> Trim$(GetPlayerName(MyIndex)) Then Exit Sub
    End If
    
    ' get the picture
    PicNum = Item(MapItem(itemNum).num).Pic

    If PicNum < 1 Or PicNum > NumItems Then Exit Sub
    ItemTimer(PicNum) = timeGetTime + SurfaceTimerMax

    If DDS_Item(PicNum) Is Nothing Then
        Call InitDDSurf("items\" & PicNum, DDSD_Item(PicNum), DDS_Item(PicNum))
    End If

    If DDSD_Item(PicNum).lWidth > 64 Then ' has more than 1 frame
        With rec
            .top = 0
            .Bottom = 32
            .Left = (MapItem(itemNum).Frame * 32)
            .Right = .Left + 32
        End With
    Else
        With rec
            .top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    End If

    Call Engine_BltFast(ConvertMapX(MapItem(itemNum).X * PIC_X), ConvertMapY(MapItem(itemNum).Y * PIC_Y), DDS_Item(PicNum), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ScreenshotMap()
Dim X As Long, Y As Long, i As Long, rec As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Clear the surface
    Set DDS_Map = Nothing
    
    ' Initialize it
    With DDSD_Map
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSD_Temp.ddsCaps.lCaps
        .lWidth = (Map.MaxX + 1) * 32
        .lHeight = (Map.MaxY + 1) * 32
    End With
    Set DDS_Map = DD.CreateSurface(DDSD_Map)
    
    ' Render the tiles
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            With Map.Tile(X, Y)
                For i = MapLayer.Ground To MapLayer.Fringe2
                    ' Skip tile
                    If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                        If .Layer(i).Tileset >= 1 And .Layer(i).Tileset <= NumTileSets Then
                            If Not i = MapLayer.Autotile Then
                                ' Sort out rec
                                rec.top = .Layer(i).Y * PIC_Y
                                rec.Bottom = rec.top + PIC_Y
                                rec.Left = .Layer(i).X * PIC_X
                                rec.Right = rec.Left + PIC_X
                            End If
                        End If
                        ' Render
                        DDS_Map.BltFast X * PIC_X, Y * PIC_Y, DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    End If
                    
                    If i = MapLayer.Autotile Then
                        BltAutotile X, Y, True
                    End If
                Next
            End With
        Next
    Next
    
    ' Render the resources
    For Y = 0 To Map.MaxY
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).Y = Y Then
                            Call BltMapResource(i, True)
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' Render the tiles
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            With Map.Tile(X, Y)
                For i = MapLayer.Fringe To MapLayer.Fringe2
                    ' Skip tile
                    If (.Layer(i).Tileset > 0 And .Layer(i).Tileset <= NumTileSets) And (.Layer(i).X > 0 Or .Layer(i).Y > 0) Then
                        If Not i = MapLayer.Autotile Then
                            ' Sort out rec
                            rec.top = .Layer(i).Y * PIC_Y
                            rec.Bottom = rec.top + PIC_Y
                            rec.Left = .Layer(i).X * PIC_X
                            rec.Right = rec.Left + PIC_X
                            ' Render
                            DDS_Map.BltFast X * PIC_X, Y * PIC_Y, DDS_Tileset(.Layer(i).Tileset), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                        End If
                    End If
                Next
            End With
        Next
    Next
    
    ' Dump and save
    frmMain.picSSMap.width = DDSD_Map.lWidth
    frmMain.picSSMap.height = DDSD_Map.lHeight
    rec.top = 0
    rec.Left = 0
    rec.Bottom = DDSD_Map.lHeight
    rec.Right = DDSD_Map.lWidth
    Engine_BltToDC DDS_Map, rec, rec, frmMain.picSSMap
    SavePicture frmMain.picSSMap.Image, App.Path & "\map" & GetPlayerMap(MyIndex) & ".jpg"
    
    ' let them know we did it
    AddText "Screenshot of map #" & GetPlayerMap(MyIndex) & " saved.", BrightGreen
    Exit Sub
    
' Error handler
errorhandler:
    Exit Sub
    HandleError "ScreenshotMap", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltMapResource(ByVal Resource_num As Long, Optional ByVal ScreenShot As Boolean = False)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim rec As DXVBLib.RECT
Dim X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' make sure it's not out of map
    If MapResource(Resource_num).X > Map.MaxX Then Exit Sub
    If MapResource(Resource_num).Y > Map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_master = Map.Tile(MapResource(Resource_num).X, MapResource(Resource_num).Y).Data1
    
    If Resource_master = 0 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If
    
    ' cut down everything if we're editing
    If InMapEditor Then
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' Load early
    If DDS_Resource(Resource_sprite) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource_sprite, DDSD_Resource(Resource_sprite), DDS_Resource(Resource_sprite))
    End If

    ' src rect
    With rec
        .top = 0
        .Bottom = DDSD_Resource(Resource_sprite).lHeight
        .Left = 0
        .Right = DDSD_Resource(Resource_sprite).lWidth
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_num).X * PIC_X) - (DDSD_Resource(Resource_sprite).lWidth / 2) + 16
    Y = (MapResource(Resource_num).Y * PIC_Y) - DDSD_Resource(Resource_sprite).lHeight + 32
    
    ' render it
    If Not ScreenShot Then
        Call BltResource(Resource_sprite, X, Y, rec)
    Else
        Call ScreenshotResource(Resource_sprite, X, Y, rec)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltMapResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltResource(ByVal Resource As Long, ByVal dx As Long, dy As Long, rec As DXVBLib.RECT)
Dim X As Long
Dim Y As Long
Dim width As Long
Dim height As Long
Dim destRECT As DXVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub
    
    ResourceTimer(Resource) = timeGetTime + SurfaceTimerMax

    If DDS_Resource(Resource) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource, DDSD_Resource(Resource), DDS_Resource(Resource))
    End If

    X = ConvertMapX(dx)
    Y = ConvertMapY(dy)
    
    width = (rec.Right - rec.Left)
    height = (rec.Bottom - rec.top)

    If Y < 0 Then
        With rec
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + height - DDSD_BackBuffer.lHeight)
    End If

    If X + width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + width - DDSD_BackBuffer.lWidth)
    End If

    ' End clipping
    Call Engine_BltFast(X, Y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub ScreenshotResource(ByVal Resource As Long, ByVal X As Long, Y As Long, rec As DXVBLib.RECT)
Dim width As Long
Dim height As Long
Dim destRECT As DXVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub
    
    ResourceTimer(Resource) = timeGetTime + SurfaceTimerMax

    If DDS_Resource(Resource) Is Nothing Then
        Call InitDDSurf("Resources\" & Resource, DDSD_Resource(Resource), DDS_Resource(Resource))
    End If
    
    width = (rec.Right - rec.Left)
    height = (rec.Bottom - rec.top)

    If Y < 0 Then
        With rec
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + height > DDSD_Map.lHeight Then
        rec.Bottom = rec.Bottom - (Y + height - DDSD_Map.lHeight)
    End If

    If X + width > DDSD_Map.lWidth Then
        rec.Right = rec.Right - (X + width - DDSD_Map.lWidth)
    End If

    ' End clipping
    DDS_Map.BltFast X, Y, DDS_Resource(Resource), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScreenshotResource", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltBars()
Dim tmpY As Long, tmpX As Long
Dim sWidth As Long, sHeight As Long
Dim sRECT As RECT
Dim barWidth As Long
Dim i As Long, NpcNum As Long, partyIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' dynamic bar calculations
    sWidth = DDSD_Bars.lWidth
    sHeight = DDSD_Bars.lHeight / 4
    
    ' render health bars
    For i = 1 To MAX_MAP_NPCS
        NpcNum = MapNpc(i).num
        ' exists?
        If NpcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < Npc(NpcNum).HP Then
                ' lock to npc
                tmpX = MapNpc(i).X * PIC_X + MapNpc(i).XOffset + 16 - (sWidth / 2)
                tmpY = MapNpc(i).Y * PIC_Y + MapNpc(i).YOffset + 35
                
                ' calculate the width to fill
                barWidth = ((MapNpc(i).Vital(Vitals.HP) / sWidth) / (Npc(NpcNum).HP / sWidth)) * sWidth
                
                ' draw bar background
                With sRECT
                    .top = sHeight * 1 ' HP bar background
                    .Left = 0
                    .Right = .Left + sWidth
                    .Bottom = .top + sHeight
                End With
                Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                
                ' draw the bar proper
                With sRECT
                    .top = 0 ' HP bar
                    .Left = 0
                    .Right = .Left + barWidth
                    .Bottom = .top + sHeight
                End With
                Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).YOffset + 35 + sHeight + 1
            
            ' calculate the width to fill
            barWidth = (timeGetTime - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer)).CastTime * 1000)) * sWidth
            
            ' draw bar background
            With sRECT
                .top = sHeight * 3 ' cooldown bar background
                .Left = 0
                .Right = sWidth
                .Bottom = .top + sHeight
            End With
            Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
            
            ' draw the bar proper
            With sRECT
                .top = sHeight * 2 ' cooldown bar
                .Left = 0
                .Right = barWidth
                .Bottom = .top + sHeight
            End With
            Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        End If
    End If
    
    ' draw own health bar
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).YOffset + 35
       
        ' calculate the width to fill
        barWidth = ((GetPlayerVital(MyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / sWidth)) * sWidth
       
        ' draw bar background
        With sRECT
            .top = sHeight * 1 ' HP bar background
            .Left = 0
            .Right = .Left + sWidth
            .Bottom = .top + sHeight
        End With
        Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
       
        ' draw the bar proper
        With sRECT
            .top = 0 ' HP bar
            .Left = 0
            .Right = .Left + barWidth
            .Bottom = .top + sHeight
        End With
        Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End If
    
    ' draw party health bars
    If Party.Leader > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            partyIndex = Party.Member(i)
            If (partyIndex > 0) And (partyIndex <> MyIndex) And (GetPlayerMap(partyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(partyIndex, Vitals.HP) > 0 And GetPlayerVital(partyIndex, Vitals.HP) < GetPlayerMaxVital(partyIndex, Vitals.HP) Then
                    ' lock to Player
                    tmpX = GetPlayerX(partyIndex) * PIC_X + Player(partyIndex).XOffset + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(partyIndex) * PIC_X + Player(partyIndex).YOffset + 35
                    
                    ' calculate the width to fill
                    barWidth = ((GetPlayerVital(partyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(partyIndex, Vitals.HP) / sWidth)) * sWidth
                    
                    ' draw bar background
                    With sRECT
                        .top = sHeight * 1 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .top + sHeight
                    End With
                    Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    
                    ' draw the bar proper
                    With sRECT
                        .top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + barWidth
                        .Bottom = .top + sHeight
                    End With
                    Engine_BltFast ConvertMapX(tmpX), ConvertMapY(tmpY), DDS_Bars, sRECT, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                End If
            End If
        Next
    End If
                    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBars", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltHotbar()
Dim sRECT As RECT, dRECT As RECT, i As Long, num As String, n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmMain.picHotbar.Cls
    For i = 1 To MAX_HOTBAR
        With dRECT
            .top = HotbarTop
            .Left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
            .Bottom = .top + 32
            .Right = .Left + 32
        End With
        
        With sRECT
            .top = 0
            .Left = 32
            .Bottom = 32
            .Right = 64
        End With
        
        Select Case Hotbar(i).sType
            Case 1 ' inventory
                If Len(Item(Hotbar(i).Slot).name) > 0 Then
                    If Item(Hotbar(i).Slot).Pic > 0 Then
                        If DDS_Item(Item(Hotbar(i).Slot).Pic) Is Nothing Then
                            Call InitDDSurf("Items\" & Item(Hotbar(i).Slot).Pic, DDSD_Item(Item(Hotbar(i).Slot).Pic), DDS_Item(Item(Hotbar(i).Slot).Pic))
                        End If
                        Engine_BltToDC DDS_Item(Item(Hotbar(i).Slot).Pic), sRECT, dRECT, frmMain.picHotbar, False
                    End If
                End If
            Case 2 ' spell
                With sRECT
                    .top = 0
                    .Left = 0
                    .Bottom = 32
                    .Right = 32
                End With
                If Len(Spell(Hotbar(i).Slot).name) > 0 Then
                    If Spell(Hotbar(i).Slot).Icon > 0 Then
                        If DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon) Is Nothing Then
                            Call InitDDSurf("Spellicons\" & Spell(Hotbar(i).Slot).Icon, DDSD_SpellIcon(Spell(Hotbar(i).Slot).Icon), DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon))
                        End If
                        ' check for cooldown
                        For n = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(n) = Hotbar(i).Slot Then
                                ' has spell
                                If Not SpellCD(i) = 0 Then
                                    sRECT.Left = 32
                                    sRECT.Right = 64
                                End If
                            End If
                        Next
                        Engine_BltToDC DDS_SpellIcon(Spell(Hotbar(i).Slot).Icon), sRECT, dRECT, frmMain.picHotbar, False
                    End If
                End If
        End Select
        
        ' render the numbers
        num = Str(i)
        Select Case i
            Case 10: num = Str(0)
            Case 11: num = "-"
            Case 12: num = "="
        End Select
        
        DrawText frmMain.picHotbar.hdc, dRECT.Left + 2, dRECT.top + 16, num, QBColor(White)
    Next
    frmMain.picHotbar.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltHotbar", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte, i As Long, X As Long, Y As Long
Dim Sprite As Long, spritetop As Long
Dim rec As DXVBLib.RECT
Dim attackspeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = GetPlayerSprite(Index)

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    
    CharacterTimer(Sprite) = timeGetTime + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    ' speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Index, Weapon)).Speed
    Else
        attackspeed = 1000
    End If

    ' Reset frame
    If Player(Index).Step = 3 Then
        Anim = 0
    ElseIf Player(Index).Step = 1 Then
        Anim = 2
    End If
    
    ' Check for attacking animation
    If Player(Index).AttackTimer + (attackspeed / 2) > timeGetTime Then
        If Player(Index).Attacking = 1 Then
            Anim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).YOffset > 8) Then Anim = Player(Index).Step
            Case DIR_DOWN
                If (Player(Index).YOffset < -8) Then Anim = Player(Index).Step
            Case DIR_LEFT
                If (Player(Index).XOffset > 8) Then Anim = Player(Index).Step
            Case DIR_RIGHT
                If (Player(Index).XOffset < -8) Then Anim = Player(Index).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + attackspeed < timeGetTime Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .top = spritetop * (DDSD_Character(Sprite).lHeight / 4)
        .Bottom = .top + (DDSD_Character(Sprite).lHeight / 4)
        .Left = Anim * (DDSD_Character(Sprite).lWidth / 4)
        .Right = .Left + (DDSD_Character(Sprite).lWidth / 4)
    End With

    ' Calculate the X
    X = GetPlayerX(Index) * PIC_X + Player(Index).XOffset - ((DDSD_Character(Sprite).lWidth / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (DDSD_Character(Sprite).lHeight) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - ((DDSD_Character(Sprite).lHeight / 4) - 32)
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset
    End If

    ' render the actual sprite
    Call BltSprite(Sprite, X, Y, rec)
    
    ' check for paperdolling
    For i = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                Call BltPaperdoll(X, Y, Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, Anim, spritetop)
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPlayer", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte, i As Long, X As Long, Y As Long, Sprite As Long, spritetop As Long
Dim rec As DXVBLib.RECT
Dim attackspeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MapNpc(MapNpcNum).num = 0 Then Exit Sub ' no npc set
    Sprite = Npc(MapNpc(MapNpcNum).num).Sprite
    
    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    CharacterTimer(Sprite) = timeGetTime + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    attackspeed = 1000

    ' Reset frame
    Anim = 0
    ' Check for attacking animation
    If MapNpc(MapNpcNum).AttackTimer + (attackspeed / 2) > timeGetTime Then
        If MapNpc(MapNpcNum).Attacking = 1 Then
            Anim = 3
        End If
    Else
        ' If not attacking, walk normally
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset > 8) Then Anim = MapNpc(MapNpcNum).Step
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < -8) Then Anim = MapNpc(MapNpcNum).Step
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset > 8) Then Anim = MapNpc(MapNpcNum).Step
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < -8) Then Anim = MapNpc(MapNpcNum).Step
        End Select
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + attackspeed < timeGetTime Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNpc(MapNpcNum).Dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .top = (DDSD_Character(Sprite).lHeight / 4) * spritetop
        .Bottom = .top + DDSD_Character(Sprite).lHeight / 4
        .Left = Anim * (DDSD_Character(Sprite).lWidth / 4)
        .Right = .Left + (DDSD_Character(Sprite).lWidth / 4)
    End With

    ' Calculate the X
    X = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset - ((DDSD_Character(Sprite).lWidth / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (DDSD_Character(Sprite).lHeight / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset - ((DDSD_Character(Sprite).lHeight / 4) - 32)
    Else
        ' Proceed as normal
        Y = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset
    End If

    Call BltSprite(Sprite, X, Y, rec)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltNpc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltPaperdoll(ByVal x2 As Long, ByVal y2 As Long, ByVal Sprite As Long, ByVal Anim As Long, ByVal spritetop As Long)
Dim rec As DXVBLib.RECT
Dim X As Long, Y As Long
Dim width As Long, height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub
    
    If DDS_Paperdoll(Sprite) Is Nothing Then
        Call InitDDSurf("Paperdolls\" & Sprite, DDSD_Paperdoll(Sprite), DDS_Paperdoll(Sprite))
    End If
    
    With rec
        .top = spritetop * (DDSD_Paperdoll(Sprite).lHeight / 4)
        .Bottom = .top + (DDSD_Paperdoll(Sprite).lHeight / 4)
        .Left = Anim * (DDSD_Paperdoll(Sprite).lWidth / 4)
        .Right = .Left + (DDSD_Paperdoll(Sprite).lWidth / 4)
    End With
    
    ' clipping
    X = ConvertMapX(x2)
    Y = ConvertMapY(y2)
    width = (rec.Right - rec.Left)
    height = (rec.Bottom - rec.top)

    ' Clip to screen
    If Y < 0 Then
        With rec
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + height - DDSD_BackBuffer.lHeight)
    End If

    If X + width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + width - DDSD_BackBuffer.lWidth)
    End If
    
    Call Engine_BltFast(X, Y, DDS_Paperdoll(Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPaperdoll", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub BltSprite(ByVal Sprite As Long, ByVal x2 As Long, y2 As Long, rec As DXVBLib.RECT)
Dim X As Long
Dim Y As Long
Dim width As Long
Dim height As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    X = ConvertMapX(x2)
    Y = ConvertMapY(y2)
    width = (rec.Right - rec.Left)
    height = (rec.Bottom - rec.top)

    ' clipping
    If Y < 0 Then
        With rec
            .top = .top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If

    If Y + height > DDSD_BackBuffer.lHeight Then
        rec.Bottom = rec.Bottom - (Y + height - DDSD_BackBuffer.lHeight)
    End If

    If X + width > DDSD_BackBuffer.lWidth Then
        rec.Right = rec.Right - (X + width - DDSD_BackBuffer.lWidth)
    End If
    
    Call Engine_BltFast(X, Y, DDS_Character(Sprite), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltAnimatedInvItems()
Dim i As Long
Dim itemNum As Long, itemPic As Long
Dim X As Long, Y As Long
Dim MaxFrames As Byte
Dim Amount As Long
Dim rec As RECT, positionRec As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    ' check for map animation changes#
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(i).num > 0 Then
            itemPic = Item(MapItem(i).num).Pic

            If itemPic < 1 Or itemPic > NumItems Then Exit Sub
            MaxFrames = (DDSD_Item(itemPic).lWidth / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

            If MapItem(i).Frame < MaxFrames - 1 Then
                MapItem(i).Frame = MapItem(i).Frame + 1
            Else
                MapItem(i).Frame = 0
            End If
        End If
    Next

    For i = 1 To MAX_INV
        itemNum = GetPlayerInvItemNum(MyIndex, i)

        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            itemPic = Item(itemNum).Pic

            If itemPic > 0 And itemPic <= NumItems Then
                If DDSD_Item(itemPic).lWidth > 64 Then
                    MaxFrames = (DDSD_Item(itemPic).lWidth / 2) / 32 ' Work out how many frames there are. /2 because of inventory icons as well as ingame

                    If InvItemFrame(i) < MaxFrames - 1 Then
                        InvItemFrame(i) = InvItemFrame(i) + 1
                    Else
                        InvItemFrame(i) = 0
                    End If

                    With rec
                        .top = 0
                        .Bottom = 32
                        .Left = (DDSD_Item(itemPic).lWidth / 2) + (InvItemFrame(i) * 32) ' middle to get the start of inv gfx, then +32 for each frame
                        .Right = .Left + 32
                    End With

                    With positionRec
                        .top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    ' Load item if not loaded, and reset timer
                    ItemTimer(itemPic) = timeGetTime + SurfaceTimerMax

                    If DDS_Item(itemPic) Is Nothing Then
                        Call InitDDSurf("Items\" & itemPic, DDSD_Item(itemPic), DDS_Item(itemPic))
                    End If

                    ' We'll now re-blt the item, and place the currency value over it again :P
                    Engine_BltToDC DDS_Item(itemPic), rec, positionRec, frmMain.picInventory, False

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = positionRec.top + 22
                        X = positionRec.Left - 4
                        Amount = CStr(GetPlayerInvItemValue(MyIndex, i))
                        ' Draw currency but with k, m, b etc. using a convertion function
                        DrawText frmMain.picInventory.hdc, X, Y, ConvertCurrency(Amount), QBColor(Yellow)

                        ' Check if it's gold, and update the label
                        If GetPlayerInvItemNum(MyIndex, i) = 1 Then '1 = gold :P
                            frmMain.lblGold.Caption = Format$(Amount, "#,###,###,###") & "g"
                        End If
                    End If
                End If
            End If
        End If
    Next

    ' refresh it to keep up
    frmMain.picInventory.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltAnimatedInvItems", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltFace()
Dim rec As RECT, positionRec As RECT, faceNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumFaces = 0 Then Exit Sub
    frmMain.picFace.Cls
    
    faceNum = GetPlayerSprite(MyIndex)
    If faceNum <= 0 Or faceNum > NumFaces Then Exit Sub

    With rec
        .top = 0
        .Bottom = 100
        .Left = 0
        .Right = 100
    End With

    With positionRec
        .top = 0
        .Bottom = 100
        .Left = 0
        .Right = 100
    End With

    ' Load face if not loaded, and reset timer
    FaceTimer(faceNum) = timeGetTime + SurfaceTimerMax

    If DDS_Face(faceNum) Is Nothing Then
        Call InitDDSurf("Faces\" & faceNum, DDSD_Face(faceNum), DDS_Face(faceNum))
    End If

    Engine_BltToDC DDS_Face(faceNum), rec, positionRec, frmMain.picFace, False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltFace", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltEquipment()
Dim i As Long, itemNum As Long, itemPic As Long
Dim itemPicRec As RECT, positionRec As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumItems = 0 Then Exit Sub
    frmMain.picCharacter.Cls

    For i = 1 To Equipment.Equipment_Count - 1
        itemNum = GetPlayerEquipment(MyIndex, i)

        If itemNum > 0 Then
            itemPic = Item(itemNum).Pic

            With itemPicRec
                .top = 0
                .Bottom = 32
                .Left = 32
                .Right = 64
            End With

            ' Change this if you've edited your character screen
            With positionRec
                .top = EqTop
                .Bottom = .top + PIC_Y
                .Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            ' Load item if not loaded, and reset timer
            ItemTimer(itemPic) = timeGetTime + SurfaceTimerMax

            If DDS_Item(itemPic) Is Nothing Then
                Call InitDDSurf("Items\" & itemPic, DDSD_Item(itemPic), DDS_Item(itemPic))
            End If

            Engine_BltToDC DDS_Item(itemPic), itemPicRec, positionRec, frmMain.picCharacter, False
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltEquipment", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltInventory()
Dim i As Long, X As Long, Y As Long, itemNum As Long, itemPic As Long
Dim Amount As Long
Dim itemPicRec As RECT, positionRec As RECT
Dim colour As Long
Dim tmpItem As Long, amountModifier As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    ' reset gold label
    frmMain.lblGold.Caption = "0g"
    frmMain.picInventory.Cls

    For i = 1 To MAX_INV
        itemNum = GetPlayerInvItemNum(MyIndex, i)

        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            itemPic = Item(itemNum).Pic
            amountModifier = 0
            
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For X = 1 To MAX_INV
                    tmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(X).num)
                    If TradeYourOffer(X).num = i Then
                        ' check if currency
                        If Not Item(tmpItem).Type = ITEM_TYPE_CURRENCY Then
                            ' normal item, exit out
                            GoTo NextLoop
                        Else
                            ' if amount = all currency, remove from inventory
                            If TradeYourOffer(X).value = GetPlayerInvItemValue(MyIndex, i) Then
                                GoTo NextLoop
                            Else
                                ' not all, change modifier to show change in currency count
                                amountModifier = TradeYourOffer(X).value
                            End If
                        End If
                    End If
                Next
            End If

            If itemPic > 0 And itemPic <= NumItems Then
                If DDSD_Item(itemPic).lWidth <= 64 Then ' more than 1 frame is handled by anim sub

                    With itemPicRec
                        .top = 0
                        .Bottom = 32
                        .Left = 32
                        .Right = 64
                    End With
                    
                    ' Change this if you've edited your inv screen
                    With positionRec
                        .top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    ' Load item if not loaded, and reset timer
                    ItemTimer(itemPic) = timeGetTime + SurfaceTimerMax

                    If DDS_Item(itemPic) Is Nothing Then
                        Call InitDDSurf("Items\" & itemPic, DDSD_Item(itemPic), DDS_Item(itemPic))
                    End If
                    Engine_BltToDC DDS_Item(itemPic), itemPicRec, positionRec, frmMain.picInventory, False

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = positionRec.top + 22
                        X = positionRec.Left - 4
                        
                        Amount = GetPlayerInvItemValue(MyIndex, i) - amountModifier
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If Amount < 1000000 Then
                            colour = QBColor(White)
                        ElseIf Amount > 1000000 And Amount < 10000000 Then
                            colour = QBColor(Yellow)
                        ElseIf Amount > 10000000 Then
                            colour = QBColor(BrightGreen)
                        End If
                        
                        ' Draw the amount in the stack
                        DrawText frmMain.picInventory.hdc, X, Y, Format$(ConvertCurrency(Str(Amount)), "#,###,###,###"), colour

                        ' Check if it's gold, and update the label
                        If GetPlayerInvItemNum(MyIndex, i) = 1 Then ' 1 = gold
                            frmMain.lblGold.Caption = Format$(Amount, "#,###,###,###") & "g"
                        End If
                    End If
                End If
            End If
        End If
NextLoop:
    Next
    
    ' Update animated items
    BltAnimatedInvItems
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventory", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltTrade()
Dim i As Long, X As Long, Y As Long, itemNum As Long, itemPic As Long
Dim Amount As Long
Dim rec As RECT, positionRec As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    frmMain.picYourTrade.Cls
    frmMain.picTheirTrade.Cls
    
    For i = 1 To MAX_INV
        ' blt your own offer
        itemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)

        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            itemPic = Item(itemNum).Pic

            If itemPic > 0 And itemPic <= NumItems Then
                With rec
                    .top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With

                With positionRec
                    .top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .top + PIC_Y
                    .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With

                ' Load item if not loaded, and reset timer
                ItemTimer(itemPic) = timeGetTime + SurfaceTimerMax

                If DDS_Item(itemPic) Is Nothing Then
                    Call InitDDSurf("Items\" & itemPic, DDSD_Item(itemPic), DDS_Item(itemPic))
                End If

                Engine_BltToDC DDS_Item(itemPic), rec, positionRec, frmMain.picYourTrade, False

                ' If item is a stack - draw the amount you have
                If TradeYourOffer(i).value > 1 Then
                    Y = positionRec.top + 22
                    X = positionRec.Left - 4
                    
                    Amount = TradeYourOffer(i).value
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        colour = QBColor(White)
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf Amount > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If
                    
                    DrawText frmMain.picYourTrade.hdc, X, Y, ConvertCurrency(Str(Amount)), colour
                End If
            End If
        End If
            
        ' blt their offer
        itemNum = TradeTheirOffer(i).num

        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            itemPic = Item(itemNum).Pic

            If itemPic > 0 And itemPic <= NumItems Then
                With rec
                    .top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With

                With positionRec
                    .top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .top + PIC_Y
                    .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With

                ' Load item if not loaded, and reset timer
                ItemTimer(itemPic) = timeGetTime + SurfaceTimerMax

                If DDS_Item(itemPic) Is Nothing Then
                    Call InitDDSurf("Items\" & itemPic, DDSD_Item(itemPic), DDS_Item(itemPic))
                End If

                Engine_BltToDC DDS_Item(itemPic), rec, positionRec, frmMain.picTheirTrade, False

                ' If item is a stack - draw the amount you have
                If TradeTheirOffer(i).value > 1 Then
                    Y = positionRec.top + 22
                    X = positionRec.Left - 4
                    
                    Amount = TradeTheirOffer(i).value
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        colour = QBColor(White)
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf Amount > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If
                    
                    DrawText frmMain.picTheirTrade.hdc, X, Y, ConvertCurrency(Str(Amount)), colour
                End If
            End If
        End If
    Next
    
    frmMain.picYourTrade.Refresh
    frmMain.picTheirTrade.Refresh
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTrade", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltPlayerSpells()
Dim i As Long, X As Long, Y As Long, spellnum As Long, spellicon As Long
Dim Amount As String
Dim rec As RECT, positionRec As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    frmMain.picSpells.Cls

    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(i)

        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            spellicon = Spell(spellnum).Icon

            If spellicon > 0 And spellicon <= NumSpellIcons Then
            
                With rec
                    .top = 0
                    .Bottom = 32
                    .Left = 0
                    .Right = 32
                End With
                
                If Not SpellCD(i) = 0 Then
                    rec.Left = 32
                    rec.Right = 64
                End If

                With positionRec
                    .top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                    .Bottom = .top + PIC_Y
                    .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                    .Right = .Left + PIC_X
                End With

                ' Load spellicon if not loaded, and reset timer
                SpellIconTimer(spellicon) = timeGetTime + SurfaceTimerMax

                If DDS_SpellIcon(spellicon) Is Nothing Then
                    Call InitDDSurf("SpellIcons\" & spellicon, DDSD_SpellIcon(spellicon), DDS_SpellIcon(spellicon))
                End If

                Engine_BltToDC DDS_SpellIcon(spellicon), rec, positionRec, frmMain.picSpells, False
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltPlayerSpells", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltShop()
Dim i As Long, X As Long, Y As Long, itemNum As Long, itemPic As Long
Dim Amount As String
Dim rec As RECT, positionRec As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    frmMain.picShopItems.Cls

    For i = 1 To MAX_TRADES
        itemNum = Shop(InShop).TradeItem(i).Item 'GetPlayerInvItemNum(MyIndex, i)
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            itemPic = Item(itemNum).Pic
            If itemPic > 0 And itemPic <= NumItems Then
            
                With rec
                    .top = 0
                    .Bottom = 32
                    .Left = 32
                    .Right = 64
                End With
                
                With positionRec
                    .top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                    .Bottom = .top + PIC_Y
                    .Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                    .Right = .Left + PIC_X
                End With
                
                ' Load item if not loaded, and reset timer
                ItemTimer(itemPic) = timeGetTime + SurfaceTimerMax
                
                If DDS_Item(itemPic) Is Nothing Then
                    Call InitDDSurf("Items\" & itemPic, DDSD_Item(itemPic), DDS_Item(itemPic))
                End If
                
                Engine_BltToDC DDS_Item(itemPic), rec, positionRec, frmMain.picShopItems, False
                
                ' If item is a stack - draw the amount you have
                If Shop(InShop).TradeItem(i).ItemValue > 1 Then
                    Y = positionRec.top + 22
                    X = positionRec.Left - 4
                    Amount = CStr(Shop(InShop).TradeItem(i).ItemValue)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = QBColor(White)
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If
                    
                    DrawText frmMain.picShopItems.hdc, X, Y, ConvertCurrency(Amount), colour
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltShop", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltInventoryItem(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT, positionRec As RECT
Dim itemNum As Long, itemPic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)

    If itemNum > 0 And itemNum <= MAX_ITEMS Then
        itemPic = Item(itemNum).Pic
        
        If itemPic = 0 Then Exit Sub

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .Left = DDSD_Item(itemPic).lWidth / 2
            .Right = .Left + PIC_X
        End With

        With positionRec
            .top = 2
            .Bottom = .top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        ' Load item if not loaded, and reset timer
        ItemTimer(itemPic) = timeGetTime + SurfaceTimerMax

        If DDS_Item(itemPic) Is Nothing Then
            Call InitDDSurf("Items\" & itemPic, DDSD_Item(itemPic), DDS_Item(itemPic))
        End If

        Engine_BltToDC DDS_Item(itemPic), rec, positionRec, frmMain.picTempInv, False

        With frmMain.picTempInv
            .top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventoryItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltDraggedSpell(ByVal X As Long, ByVal Y As Long)
Dim rec As RECT, positionRec As RECT
Dim spellnum As Long, spellpic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spellnum = PlayerSpells(DragSpell)

    If spellnum > 0 And spellnum <= MAX_SPELLS Then
        spellpic = Spell(spellnum).Icon
        
        If spellpic = 0 Then Exit Sub

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With

        With positionRec
            .top = 2
            .Bottom = .top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        ' Load item if not loaded, and reset timer
        SpellIconTimer(spellpic) = timeGetTime + SurfaceTimerMax

        If DDS_SpellIcon(spellpic) Is Nothing Then
            Call InitDDSurf("Spellicons\" & spellpic, DDSD_SpellIcon(spellpic), DDS_SpellIcon(spellpic))
        End If

        Engine_BltToDC DDS_SpellIcon(spellpic), rec, positionRec, frmMain.picTempSpell, False

        With frmMain.picTempSpell
            .top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltInventoryItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltItemDesc(ByVal itemNum As Long)
Dim rec As RECT, positionRec As RECT
Dim itemPic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmMain.picItemDescPic.Cls
    
    If itemNum > 0 And itemNum <= MAX_ITEMS Then
        itemPic = Item(itemNum).Pic

        If itemPic = 0 Then Exit Sub
        
        ' Load item if not loaded, and reset timer
        ItemTimer(itemPic) = timeGetTime + SurfaceTimerMax

        If DDS_Item(itemPic) Is Nothing Then
            Call InitDDSurf("Items\" & itemPic, DDSD_Item(itemPic), DDS_Item(itemPic))
        End If

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .Left = DDSD_Item(itemPic).lWidth / 2
            .Right = .Left + PIC_X
        End With

        With positionRec
            .top = 0
            .Bottom = 64
            .Left = 0
            .Right = 64
        End With
        Engine_BltToDC DDS_Item(itemPic), rec, positionRec, frmMain.picItemDescPic, False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltItemDesc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltSpellDesc(ByVal spellnum As Long)
Dim rec As RECT, positionRec As RECT
Dim spellpic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmMain.picSpellDescPic.Cls

    If spellnum > 0 And spellnum <= MAX_SPELLS Then
        spellpic = Spell(spellnum).Icon

        If spellpic <= 0 Or spellpic > NumSpellIcons Then Exit Sub
        
        ' Load item if not loaded, and reset timer
        SpellIconTimer(spellpic) = timeGetTime + SurfaceTimerMax

        If DDS_SpellIcon(spellpic) Is Nothing Then
            Call InitDDSurf("SpellIcons\" & spellpic, DDSD_SpellIcon(spellpic), DDS_SpellIcon(spellpic))
        End If

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .Left = 0
            .Right = .Left + PIC_X
        End With

        With positionRec
            .top = 0
            .Bottom = 64
            .Left = 0
            .Right = 64
        End With
        Engine_BltToDC DDS_SpellIcon(spellpic), rec, positionRec, frmMain.picSpellDescPic, False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltSpellDesc", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ******************
' ** Game Editors **
' ******************
Public Sub EditorMap_BltTileset()
Dim height As Long
Dim width As Long
Dim Tileset As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.value
    
    ' exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTileSets Then Exit Sub
    
    ' make sure it's loaded
    If DDS_Tileset(Tileset) Is Nothing Then
        Call InitDDSurf("tilesets\" & Tileset, DDSD_Tileset(Tileset), DDS_Tileset(Tileset))
    End If
    
    height = DDSD_Tileset(Tileset).lHeight
    width = DDSD_Tileset(Tileset).lWidth
    
    dRECT.top = 0
    dRECT.Bottom = height
    dRECT.Left = 0
    dRECT.Right = width
    
    frmEditor_Map.picBackSelect.height = height
    frmEditor_Map.picBackSelect.width = width
    
    Call Engine_BltToDC(DDS_Tileset(Tileset), sRECT, dRECT, frmEditor_Map.picBackSelect)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltTileset", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltTileOutline()
Dim rec As DXVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optBlock.value Then Exit Sub

    With rec
        .top = 0
        .Bottom = .top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    Call Engine_BltFast(ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), DDS_Misc, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTileOutline", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NewCharacterBltSprite()
Dim Sprite As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT
Dim width As Long, height As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    If frmMenu.optMale.value = True Then
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).MaleSprite(newCharSprite)
    Else
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).FemaleSprite(newCharSprite)
    End If
    
    If Sprite < 1 Or Sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If
    
    CharacterTimer(Sprite) = timeGetTime + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("Characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If
    
    width = DDSD_Character(Sprite).lWidth / 4
    height = DDSD_Character(Sprite).lHeight / 4
    
    frmMenu.picSprite.width = width
    frmMenu.picSprite.height = height
    
    sRECT.top = 0
    sRECT.Bottom = sRECT.top + height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + width
    
    dRECT.top = 0
    dRECT.Bottom = height
    dRECT.Left = 0
    dRECT.Right = width
    
    Call Engine_BltToDC(DDS_Character(Sprite), sRECT, dRECT, frmMenu.picSprite)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NewCharacterBltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_BltMapItem()
Dim itemNum As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = Item(frmEditor_Map.scrlMapItem.value).Pic

    If itemNum < 1 Or itemNum > NumItems Then
        frmEditor_Map.picMapItem.Cls
        Exit Sub
    End If

    ItemTimer(itemNum) = timeGetTime + SurfaceTimerMax

    If DDS_Item(itemNum) Is Nothing Then
        Call InitDDSurf("Items\" & itemNum, DDSD_Item(itemNum), DDS_Item(itemNum))
    End If

    sRECT.top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    Call Engine_BltToDC(DDS_Item(itemNum), sRECT, dRECT, frmEditor_Map.picMapItem)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltMapItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_BltKey()
Dim itemNum As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = Item(frmEditor_Map.scrlMapKey.value).Pic

    If itemNum < 1 Or itemNum > NumItems Then
        frmEditor_Map.picMapKey.Cls
        Exit Sub
    End If

    ItemTimer(itemNum) = timeGetTime + SurfaceTimerMax

    If DDS_Item(itemNum) Is Nothing Then
        Call InitDDSurf("Items\" & itemNum, DDSD_Item(itemNum), DDS_Item(itemNum))
    End If

    sRECT.top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    Call Engine_BltToDC(DDS_Item(itemNum), sRECT, dRECT, frmEditor_Map.picMapKey)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_BltKey", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_BltItem()
Dim itemNum As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = frmEditor_Item.scrlPic.value

    If itemNum < 1 Or itemNum > NumItems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If

    ItemTimer(itemNum) = timeGetTime + SurfaceTimerMax

    If DDS_Item(itemNum) Is Nothing Then
        Call InitDDSurf("Items\" & itemNum, DDSD_Item(itemNum), DDS_Item(itemNum))
    End If

    ' rect for source
    sRECT.top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' same for destination as source
    dRECT = sRECT
    Call Engine_BltToDC(DDS_Item(itemNum), sRECT, dRECT, frmEditor_Item.picItem)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_BltItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_BltPaperdoll()
Dim Sprite As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmEditor_Item.picPaperdoll.Cls
    
    Sprite = frmEditor_Item.scrlPaperdoll.value

    If Sprite < 1 Or Sprite > NumPaperdolls Then
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    PaperdollTimer(Sprite) = timeGetTime + SurfaceTimerMax

    If DDS_Paperdoll(Sprite) Is Nothing Then
        Call InitDDSurf("paperdolls\" & Sprite, DDSD_Paperdoll(Sprite), DDS_Paperdoll(Sprite))
    End If

    ' rect for source
    sRECT.top = 0
    sRECT.Bottom = DDSD_Paperdoll(Sprite).lHeight
    sRECT.Left = 0
    sRECT.Right = DDSD_Paperdoll(Sprite).lWidth
    ' same for destination as source
    dRECT = sRECT
    
    Call Engine_BltToDC(DDS_Paperdoll(Sprite), sRECT, dRECT, frmEditor_Item.picPaperdoll)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_BltPaperdoll", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_BltIcon()
Dim iconnum As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    iconnum = frmEditor_Spell.scrlIcon.value
    
    If iconnum < 1 Or iconnum > NumSpellIcons Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If
    
    SpellIconTimer(iconnum) = timeGetTime + SurfaceTimerMax
    
    If DDS_SpellIcon(iconnum) Is Nothing Then
        Call InitDDSurf("SpellIcons\" & iconnum, DDSD_SpellIcon(iconnum), DDS_SpellIcon(iconnum))
    End If
    
    sRECT.top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRECT.top = 0
    dRECT.Bottom = PIC_Y
    dRECT.Left = 0
    dRECT.Right = PIC_X
    
    Call Engine_BltToDC(DDS_SpellIcon(iconnum), sRECT, dRECT, frmEditor_Spell.picSprite)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorSpell_BltIcon", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorAnim_BltAnim()
Dim Animationnum As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT
Dim i As Long
Dim width As Long, height As Long
Dim looptime As Long
Dim FrameCount As Long
Dim ShouldRender As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).value
        
        If Animationnum < 1 Or Animationnum > NumAnimations Then
            frmEditor_Animation.picSprite(i).Cls
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= timeGetTime Then
                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                AnimEditorTimer(i) = timeGetTime
                ShouldRender = True
            End If
        
            If ShouldRender Then
                frmEditor_Animation.picSprite(i).Cls
            
                AnimationTimer(Animationnum) = timeGetTime + SurfaceTimerMax
                
                If DDS_Animation(Animationnum) Is Nothing Then
                    Call InitDDSurf("animations\" & Animationnum, DDSD_Animation(Animationnum), DDS_Animation(Animationnum))
                End If
                
                If frmEditor_Animation.scrlFrameCount(i).value > 0 Then
                    ' total width divided by frame count
                    width = DDSD_Animation(Animationnum).lWidth / frmEditor_Animation.scrlFrameCount(i).value
                    height = DDSD_Animation(Animationnum).lHeight
                    
                    sRECT.top = 0
                    sRECT.Bottom = height
                    sRECT.Left = (AnimEditorFrame(i) - 1) * width
                    sRECT.Right = sRECT.Left + width
                    
                    dRECT.top = 0
                    dRECT.Bottom = height
                    dRECT.Left = 0
                    dRECT.Right = width
                    
                    Call Engine_BltToDC(DDS_Animation(Animationnum), sRECT, dRECT, frmEditor_Animation.picSprite(i))
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorAnim_BltAnim", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorNpc_BltSprite()
Dim Sprite As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Sprite = frmEditor_NPC.scrlSprite.value

    If Sprite < 1 Or Sprite > NumCharacters Then
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    CharacterTimer(Sprite) = timeGetTime + SurfaceTimerMax

    If DDS_Character(Sprite) Is Nothing Then
        Call InitDDSurf("characters\" & Sprite, DDSD_Character(Sprite), DDS_Character(Sprite))
    End If

    sRECT.top = 0
    sRECT.Bottom = SIZE_Y
    sRECT.Left = PIC_X * 3 ' facing down
    sRECT.Right = sRECT.Left + SIZE_X
    dRECT.top = 0
    dRECT.Bottom = SIZE_Y
    dRECT.Left = 0
    dRECT.Right = SIZE_X
    Call Engine_BltToDC(DDS_Character(Sprite), sRECT, dRECT, frmEditor_NPC.picSprite)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorNpc_BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorResource_BltSprite()
Dim Sprite As Long
Dim sRECT As DXVBLib.RECT
Dim dRECT As DXVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' normal sprite
    Sprite = frmEditor_Resource.scrlNormalPic.value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        ResourceTimer(Sprite) = timeGetTime + SurfaceTimerMax
        If DDS_Resource(Sprite) Is Nothing Then
            Call InitDDSurf("Resources\" & Sprite, DDSD_Resource(Sprite), DDS_Resource(Sprite))
        End If
        sRECT.top = 0
        sRECT.Bottom = DDSD_Resource(Sprite).lHeight
        sRECT.Left = 0
        sRECT.Right = DDSD_Resource(Sprite).lWidth
        dRECT.top = 0
        dRECT.Bottom = DDSD_Resource(Sprite).lHeight
        dRECT.Left = 0
        dRECT.Right = DDSD_Resource(Sprite).lWidth
        Call Engine_BltToDC(DDS_Resource(Sprite), sRECT, dRECT, frmEditor_Resource.picNormalPic)
    End If

    ' exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        ResourceTimer(Sprite) = timeGetTime + SurfaceTimerMax
        If DDS_Resource(Sprite) Is Nothing Then
            Call InitDDSurf("Resources\" & Sprite, DDSD_Resource(Sprite), DDS_Resource(Sprite))
        End If
        sRECT.top = 0
        sRECT.Bottom = DDSD_Resource(Sprite).lHeight
        sRECT.Left = 0
        sRECT.Right = DDSD_Resource(Sprite).lWidth
        dRECT.top = 0
        dRECT.Bottom = DDSD_Resource(Sprite).lHeight
        dRECT.Left = 0
        dRECT.Right = DDSD_Resource(Sprite).lWidth
        Call Engine_BltToDC(DDS_Resource(Sprite), sRECT, dRECT, frmEditor_Resource.picExhaustedPic)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorResource_BltSprite", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Render_Graphics()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim rec As DXVBLib.RECT
Dim positionRec As DXVBLib.RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check if automation is screwed
    If Not CheckSurfaces Then
        ' exit out and let them know we need to re-init
        ReInitSurfaces = True
        Exit Sub
    Else
        ' if we need to fix the surfaces then do so
        If ReInitSurfaces Then
            ReInitSurfaces = False
            ReInitDD
        End If
    End If
    
    ' don't render
    If frmMain.WindowState = vbMinimized Then Exit Sub
    If GettingMap Then Exit Sub
    
    ' update the viewpoint
    UpdateCamera
    
    ' update animation editor
    If Editor = EDITOR_ANIMATION Then
        EditorAnim_BltAnim
    End If
    
    ' fill it with black
    DDS_BackBuffer.BltColorFill positionRec, 0
    
    ' blit lower tiles
    If NumTileSets > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    Call BltMapTile(X, Y)
                End If
            Next
        Next
    End If

    ' render the decals
    For i = 1 To MAX_BLOOD
        Call BltBlood(i)
    Next

    ' Blit out the items
    If NumItems > 0 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).num > 0 Then
                Call BltItem(i)
            End If
        Next
    End If
    
    ' draw animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(0) Then
                BltAnimation i, 0
            End If
        Next
    End If

    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    For Y = 0 To Map.MaxY
        If NumCharacters > 0 Then
            ' Players
            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Player(i).Y = Y Then
                        Call BltPlayer(i)
                    End If
                End If
            Next
        
            ' Npcs
            For i = 1 To Npc_HighIndex
                If MapNpc(i).Y = Y Then
                    Call BltNpc(i)
                End If
            Next
        End If
        
        ' Resources
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).Y = Y Then
                            Call BltMapResource(i)
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(1) Then
                BltAnimation i, 1
            End If
        Next
    End If

    ' blit out upper tiles
    If NumTileSets > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    Call BltMapFringeTile(X, Y)
                End If
            Next
        Next
    End If
    
    ' blit out a square at mouse cursor
    If InMapEditor Then
        If frmEditor_Map.optBlock.value = True Then
            For X = TileView.Left To TileView.Right
                For Y = TileView.top To TileView.Bottom
                    If IsValidMapPoint(X, Y) Then
                        Call BltDirection(X, Y)
                    End If
                Next
            Next
        End If
        Call BltTileOutline
    End If
    
    ' Render the bars
    BltBars
    
    ' Blt the target icon
    If myTarget > 0 Then
        If myTargetType = TARGET_TYPE_PLAYER Then
            BltTarget (Player(myTarget).X * 32) + Player(myTarget).XOffset, (Player(myTarget).Y * 32) + Player(myTarget).YOffset
        ElseIf myTargetType = TARGET_TYPE_NPC Then
            BltTarget (MapNpc(myTarget).X * 32) + MapNpc(myTarget).XOffset, (MapNpc(myTarget).Y * 32) + MapNpc(myTarget).YOffset
        End If
    End If
    
    ' blt the hover icon
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).Map = Player(MyIndex).Map Then
                If CurX = Player(i).X And CurY = Player(i).Y Then
                    If myTargetType = TARGET_TYPE_PLAYER And myTarget = i Then
                        ' dont render lol
                    Else
                        BltHover TARGET_TYPE_PLAYER, i, (Player(i).X * 32) + Player(i).XOffset, (Player(i).Y * 32) + Player(i).YOffset
                    End If
                End If
            End If
        End If
    Next
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            If CurX = MapNpc(i).X And CurY = MapNpc(i).Y Then
                If myTargetType = TARGET_TYPE_NPC And myTarget = i Then
                    ' dont render lol
                Else
                    BltHover TARGET_TYPE_NPC, i, (MapNpc(i).X * 32) + MapNpc(i).XOffset, (MapNpc(i).Y * 32) + MapNpc(i).YOffset
                End If
            End If
        End If
    Next

    ' Lock the backbuffer so we can draw text and names
    TexthDC = DDS_BackBuffer.GetDC

    ' draw FPS
    If BFPS Then
        Call DrawText(TexthDC, Camera.Right - (Len("FPS: " & GameFPS) * 8), Camera.top + 1, Trim$("FPS: " & GameFPS), QBColor(Yellow))
    End If

    ' draw cursor, player X and Y locations
    If BLoc Then
        Call DrawText(TexthDC, Camera.Left, Camera.top + 1, Trim$("cur x: " & CurX & " y: " & CurY), QBColor(Yellow))
        Call DrawText(TexthDC, Camera.Left, Camera.top + 15, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), QBColor(Yellow))
        Call DrawText(TexthDC, Camera.Left, Camera.top + 27, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), QBColor(Yellow))
    End If

    ' draw player names
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            Call DrawPlayerName(i)
        End If
    Next
    
    ' draw npc names
    For i = 1 To Npc_HighIndex
        If MapNpc(i).num > 0 Then
            Call DrawNpcName(i)
        End If
    Next
    
    For i = 1 To Action_HighIndex
        Call BltActionMsg(i)
    Next i

    ' Blit out map attributes
    If InMapEditor Then
        Call BltMapAttributes
    End If

    ' Draw map name
    Call DrawText(TexthDC, DrawMapNameX, DrawMapNameY, Map.name, DrawMapNameColor)

    ' Release DC
    DDS_BackBuffer.ReleaseDC TexthDC
    
    ' Get rec
    With rec
        .top = Camera.top
        .Bottom = .top + ScreenY
        .Left = Camera.Left
        .Right = .Left + ScreenX
    End With
    
    ' positionRec
    With positionRec
        .Bottom = ((MAX_MAPY + 1) * PIC_Y)
        .Right = ((MAX_MAPX + 1) * PIC_X)
    End With
    
    ' Flip and render
    DX7.GetWindowRect frmMain.picScreen.hWnd, positionRec
    DDS_Primary.Blt positionRec, DDS_BackBuffer, rec, DDBLT_WAIT
    
    ' Error handler
    Exit Sub
    
errorhandler:
    HandleError "Render_Graphics", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateCamera()
Dim offsetX As Long
Dim offsetY As Long
Dim StartX As Long
Dim StartY As Long
Dim EndX As Long
Dim EndY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    offsetX = Player(MyIndex).XOffset + PIC_X
    offsetY = Player(MyIndex).YOffset + PIC_Y

    StartX = GetPlayerX(MyIndex) - StartXValue
    StartY = GetPlayerY(MyIndex) - StartYValue
    If StartX < 0 Then
        offsetX = 0
        If StartX = -1 Then
            If Player(MyIndex).XOffset > 0 Then
                offsetX = Player(MyIndex).XOffset
            End If
        End If
        StartX = 0
    End If
    If StartY < 0 Then
        offsetY = 0
        If StartY = -1 Then
            If Player(MyIndex).YOffset > 0 Then
                offsetY = Player(MyIndex).YOffset
            End If
        End If
        StartY = 0
    End If
    
    EndX = StartX + EndXValue
    EndY = StartY + EndYValue
    If EndX > Map.MaxX Then
        offsetX = 32
        If EndX = Map.MaxX + 1 Then
            If Player(MyIndex).XOffset < 0 Then
                offsetX = Player(MyIndex).XOffset + PIC_X
            End If
        End If
        EndX = Map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If
    If EndY > Map.MaxY Then
        offsetY = 32
        If EndY = Map.MaxY + 1 Then
            If Player(MyIndex).YOffset < 0 Then
                offsetY = Player(MyIndex).YOffset + PIC_Y
            End If
        End If
        EndY = Map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .top = offsetY
        .Bottom = .top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With
    
    UpdateDrawMapName

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateCamera", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapX = X - (TileView.Left * PIC_X)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapX", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapY = Y - (TileView.top * PIC_Y)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapY", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function InViewPort(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InViewPort = False

    If X < TileView.Left Then Exit Function
    If Y < TileView.top Then Exit Function
    If X > TileView.Right Then Exit Function
    If Y > TileView.Bottom Then Exit Function
    InViewPort = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InViewPort", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsValidMapPoint = False

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
        
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsValidMapPoint", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub LoadTilesets()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim tilesetInUse() As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim tilesetInUse(0 To NumTileSets)
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                ' check exists
                If Map.Tile(X, Y).Layer(i).Tileset > 0 And Map.Tile(X, Y).Layer(i).Tileset <= NumTileSets Then
                    tilesetInUse(Map.Tile(X, Y).Layer(i).Tileset) = True
                End If
            Next
        Next
    Next
    
    For i = 1 To NumTileSets
        If tilesetInUse(i) Then
            ' load tileset
            If DDS_Tileset(i) Is Nothing Then
                Call InitDDSurf("tilesets\" & i, DDSD_Tileset(i), DDS_Tileset(i))
            End If
        Else
            ' unload tileset
            Call ZeroMemory(ByVal VarPtr(DDSD_Tileset(i)), LenB(DDSD_Tileset(i)))
            Set DDS_Tileset(i) = Nothing
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTilesets", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BltBank()
Dim i As Long, X As Long, Y As Long, itemNum As Long
Dim Amount As String
Dim sRECT As RECT, dRECT As RECT
Dim Sprite As Long, colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMain.picBank.Visible = True Then
        frmMain.picBank.Cls
                
        For i = 1 To MAX_BANK
            itemNum = GetBankItemNum(i)
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
            
                Sprite = Item(itemNum).Pic
                
                If Sprite <= 0 Or Sprite > NumItems Then Exit Sub
                
                If DDS_Item(Sprite) Is Nothing Then
                    Call InitDDSurf("Items\" & Sprite, DDSD_Item(Sprite), DDS_Item(Sprite))
                End If
            
                With sRECT
                    .top = 0
                    .Bottom = .top + PIC_Y
                    .Left = DDSD_Item(Sprite).lWidth / 2
                    .Right = .Left + PIC_X
                End With
                
                With dRECT
                    .top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                    .Bottom = .top + PIC_Y
                    .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                    .Right = .Left + PIC_X
                End With
                
                Engine_BltToDC DDS_Item(Sprite), sRECT, dRECT, frmMain.picBank, False

                ' If item is a stack - draw the amount you have
                If GetBankItemValue(0, i) > 1 Then
                    Y = dRECT.top + 22
                    X = dRECT.Left - 4
                
                    Amount = CStr(GetBankItemValue(0, i))
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = QBColor(White)
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = QBColor(Yellow)
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = QBColor(BrightGreen)
                    End If
                    DrawText frmMain.picBank.hdc, X, Y, ConvertCurrency(Amount), colour
                End If
            End If
        Next
    
        frmMain.picBank.Refresh
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBank", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BltBankItem(ByVal X As Long, ByVal Y As Long)
Dim sRECT As RECT, dRECT As RECT
Dim itemNum As Long
Dim Sprite As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = GetBankItemNum(DragBankSlotNum)
    Sprite = Item(GetBankItemNum(DragBankSlotNum)).Pic
    
    If DDS_Item(Sprite) Is Nothing Then
        Call InitDDSurf("Items\" & Sprite, DDSD_Item(Sprite), DDS_Item(Sprite))
    End If
    
    If itemNum > 0 Then
        If itemNum <= MAX_ITEMS Then
            With sRECT
                .top = 0
                .Bottom = .top + PIC_Y
                .Left = DDSD_Item(Sprite).lWidth / 2
                .Right = .Left + PIC_X
            End With
        End If
    End If
    
    With dRECT
        .top = 2
        .Bottom = .top + PIC_Y
        .Left = 2
        .Right = .Left + PIC_X
    End With

    Engine_BltToDC DDS_Item(Sprite), sRECT, dRECT, frmMain.picTempBank
    
    With frmMain.picTempBank
        .top = Y
        .Left = X
        .Visible = True
        .ZOrder (0)
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltBankItem", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
