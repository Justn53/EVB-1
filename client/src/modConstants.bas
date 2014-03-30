Attribute VB_Name = "modConstants"
Option Explicit

' API Declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lparam As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

' Version constants
Public Const VERSION_MAJOR As Byte = 2
Public Const VERSION_MINOR As Byte = 5
Public Const VERSION_REVISION As Byte = 1

' animated buttons
Public Const MAX_MENUBUTTONS As Long = 4
Public Const MAX_MAINBUTTONS As Long = 7
Public Const MENUBUTTON_PATH As String = "\data files\graphics\gui\menu\buttons\"
Public Const MAINBUTTON_PATH As String = "\data files\graphics\gui\main\buttons\"

' Hotbar
Public Const HotbarTop As Long = 2
Public Const HotbarLeft As Long = 2
Public Const HotbarOffsetX As Long = 8

' Inventory constants
Public Const InvTop As Long = 24
Public Const InvLeft As Long = 12
Public Const InvOffsetY As Long = 3
Public Const InvOffsetX As Long = 3
Public Const InvColumns As Long = 5

' Bank constants
Public Const BankTop As Long = 38
Public Const BankLeft As Long = 42
Public Const BankOffsetY As Long = 3
Public Const BankOffsetX As Long = 4
Public Const BankColumns As Long = 11

' spells constants
Public Const SpellTop As Long = 24
Public Const SpellLeft As Long = 12
Public Const SpellOffsetY As Long = 3
Public Const SpellOffsetX As Long = 3
Public Const SpellColumns As Long = 5

' shop constants
Public Const ShopTop As Long = 6
Public Const ShopLeft As Long = 8
Public Const ShopOffsetY As Long = 2
Public Const ShopOffsetX As Long = 4
Public Const ShopColumns As Long = 5

' character consts
Public Const EqTop As Long = 224
Public Const EqLeft As Long = 18
Public Const EqOffsetX As Long = 10
Public Const EqColumns As Long = 4

' timeout for server connection (miliseconds)
Public Const TIMEOUT As Long = 5000

' data type values
Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

' path constants
Public Const SOUND_PATH As String = "\data files\sound\"
Public Const MUSIC_PATH As String = "\data files\music\"

' font
Public Const FONT_NAME As String = "Georgia"
Public Const FONT_SIZE As Byte = 14

' log path
Public Const LOG_DEBUG As String = "debug.txt"
Public Const LOG_PATH As String = "\data files\logs\"

' map path
Public Const MAP_PATH As String = "\data files\maps\"
Public Const MAP_EXT As String = ".map"

' graphics path
Public Const GFX_PATH As String = "\data files\graphics\"
Public Const GFX_EXT As String = ".bmp"

' key constants
Public Const VK_UP As Long = &H26
Public Const VK_DOWN As Long = &H28
Public Const VK_LEFT As Long = &H25
Public Const VK_RIGHT As Long = &H27
Public Const VK_SHIFT As Long = &H10
Public Const VK_RETURN As Long = &HD
Public Const VK_CONTROL As Long = &H11

' menu states
Public Const MENU_STATE_NEWACCOUNT As Byte = 0
Public Const MENU_STATE_DELACCOUNT As Byte = 1
Public Const MENU_STATE_LOGIN As Byte = 2
Public Const MENU_STATE_GETCHARS As Byte = 3
Public Const MENU_STATE_NEWCHAR As Byte = 4
Public Const MENU_STATE_ADDCHAR As Byte = 5
Public Const MENU_STATE_DELCHAR As Byte = 6
Public Const MENU_STATE_USECHAR As Byte = 7
Public Const MENU_STATE_INIT As Byte = 8

' movement speed
Public Const WALK_SPEED As Byte = 4
Public Const RUN_SPEED As Byte = 6

' tile size constants
Public Const PIC_X As Long = 32
Public Const PIC_Y As Long = 32

' sprite, item, spell size constants
Public Const SIZE_X As Long = 32
Public Const SIZE_Y As Long = 32

' ********************************************************
' * The values below must match with the server's values *
' ********************************************************

' General constants
Public Const MAX_PLAYERS As Long = 50
Public Const MAX_ITEMS As Long = 200
Public Const MAX_NPCS As Long = 200
Public Const MAX_ANIMATIONS As Long = 200
Public Const MAX_INV As Long = 35
Public Const MAX_MAP_ITEMS As Long = 100
Public Const MAX_MAP_NPCS As Long = 30
Public Const MAX_SHOPS As Long = 50
Public Const MAX_PLAYER_SPELLS As Long = 35
Public Const MAX_SPELLS As Long = 200
Public Const MAX_TRADES As Long = 30
Public Const MAX_RESOURCES As Long = 100
Public Const MAX_LEVELS As Long = 60
Public Const MAX_BANK As Long = 99
Public Const MAX_HOTBAR As Byte = 12
Public Const MAX_PARTYS As Long = 35
Public Const MAX_PARTY_MEMBERS As Byte = 4
Public Const MAX_NPC_DROPS As Byte = 20
Public Const MAX_CONVS As Long = 200
Public Const MAX_CONV_CHATS As Byte = 20
Public Const MAX_QUESTS As Long = 200
Public Const MAX_QUEST_OBJECTIVES As Long = 3
Public Const MAX_BLOOD As Long = 100

' Website
Public Const GAME_WEBSITE As String = "http://www.indierising.net"

' text color constants
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = BrightBlue
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = BrightGreen
Public Const EmoteColor As Byte = BrightCyan
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = BrightBlue
Public Const WhoColor As Byte = BrightBlue
Public Const JoinLeftColor As Byte = DarkGrey
Public Const NpcColor As Byte = Brown
Public Const AlertColor As Byte = Red
Public Const NewMapColor As Byte = BrightBlue

' String constants
Public Const NAME_LENGTH As Byte = 20
Public Const ACCOUNT_LENGTH As Byte = 12

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
Public Const MAX_MAPS As Long = 100
Public Const MAX_MAPX As Byte = 14
Public Const MAX_MAPY As Byte = 11
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_RESOURCE As Byte = 7
Public Const TILE_TYPE_NPCSPAWN As Byte = 8
Public Const TILE_TYPE_SHOP As Byte = 9
Public Const TILE_TYPE_BANK As Byte = 10
Public Const TILE_TYPE_HEAL As Byte = 11
Public Const TILE_TYPE_TRAP As Byte = 12
Public Const TILE_TYPE_SLIDE As Byte = 13

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_HELMET As Byte = 3
Public Const ITEM_TYPE_SHIELD As Byte = 4
Public Const ITEM_TYPE_CONSUME As Byte = 5
Public Const ITEM_TYPE_KEY As Byte = 6
Public Const ITEM_TYPE_CURRENCY As Byte = 7
Public Const ITEM_TYPE_SPELL As Byte = 8

' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

' player movement in tiles per second
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' admin types
Public Const ADMIN_MONITOR As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' npc types
Public Const NPC_BEHAVIOUR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOUR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOUR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOUR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOUR_GUARD As Byte = 4

' spell types
Public Const SPELL_TYPE_DAMAGEHP As Byte = 0
Public Const SPELL_TYPE_DAMAGEMP As Byte = 1
Public Const SPELL_TYPE_HEALHP As Byte = 2
Public Const SPELL_TYPE_HEALMP As Byte = 3
Public Const SPELL_TYPE_WARP As Byte = 4

' game editors
Public Const EDITOR_ITEM As Byte = 1
Public Const EDITOR_NPC As Byte = 2
Public Const EDITOR_SPELL As Byte = 3
Public Const EDITOR_SHOP As Byte = 4
Public Const EDITOR_RESOURCE As Byte = 5
Public Const EDITOR_ANIMATION As Byte = 6
Public Const EDITOR_CONV As Byte = 7
Public Const EDITOR_QUEST As Byte = 8

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2

' Dialogue box constants
Public Const DIALOGUE_TYPE_NONE As Byte = 0
Public Const DIALOGUE_TYPE_TRADE As Byte = 1
Public Const DIALOGUE_TYPE_FORGET As Byte = 2
Public Const DIALOGUE_TYPE_PARTY As Byte = 3

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Long = 0
Public Const ACTIONMSG_SCROLL As Long = 1
Public Const ACTIONMSG_SCREEN As Long = 2

' stuffs
Public Const HalfX As Long = ((MAX_MAPX + 1) / 2) * PIC_X
Public Const HalfY As Long = ((MAX_MAPY + 1) / 2) * PIC_Y
Public Const ScreenX As Long = (MAX_MAPX + 1) * PIC_X
Public Const ScreenY As Long = (MAX_MAPY + 1) * PIC_Y
Public Const StartXValue As Long = ((MAX_MAPX + 1) / 2)
Public Const StartYValue As Long = ((MAX_MAPY + 1) / 2)
Public Const EndXValue As Long = (MAX_MAPX + 1) + 1
Public Const EndYValue As Long = (MAX_MAPY + 1) + 1
Public Const Half_PIC_X As Long = PIC_X / 2
Public Const Half_PIC_Y As Long = PIC_Y / 2
