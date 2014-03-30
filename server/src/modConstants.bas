Attribute VB_Name = "modConstants"
Option Explicit

' API
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

' Version constants
Public Const VERSION_MAJOR As Byte = 2
Public Const VERSION_MINOR As Byte = 5
Public Const VERSION_REVISION As Byte = 1

Public Const MAX_LINES As Long = 500 ' Used for frmServer.txtText

' ********************************************************
' * The values below must match with the client's values *
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

' server-side stuff
Public Const MIN_PVP_LEVEL As Long = 5
Public Const ITEM_SPAWN_TIME As Long = 30000 ' 30 seconds
Public Const ITEM_DESPAWN_TIME As Long = 90000 ' 1:30 seconds
Public Const MAX_DOTS As Long = 30

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

' string constants
Public Const NAME_LENGTH As Byte = 20
Public Const ACCOUNT_LENGTH As Byte = 12

' sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' map constants
Public Const MAX_MAPS As Long = 100
Public Const MAX_MAPX As Byte = 14
Public Const MAX_MAPY As Byte = 11
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1

' tile consants
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

' item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_HELMET As Byte = 3
Public Const ITEM_TYPE_SHIELD As Byte = 4
Public Const ITEM_TYPE_CONSUME As Byte = 5
Public Const ITEM_TYPE_KEY As Byte = 6
Public Const ITEM_TYPE_CURRENCY As Byte = 7
Public Const ITEM_TYPE_SPELL As Byte = 8

' direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

' constants for player movement
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

' target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Long = 0
Public Const ACTIONMSG_SCROLL As Long = 1
Public Const ACTIONMSG_SCREEN As Long = 2
