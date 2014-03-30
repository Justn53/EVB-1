VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_NPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   746
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Quest Data"
      Height          =   5535
      Left            =   7320
      TabIndex        =   53
      Top             =   120
      Width           =   3735
      Begin VB.HScrollBar scrlQuestIndex 
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   1440
         Width           =   3495
      End
      Begin VB.HScrollBar scrlPostQuestChat 
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   4560
         Width           =   3495
      End
      Begin VB.HScrollBar scrlQuestInProgressChat 
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   3360
         Width           =   3495
      End
      Begin VB.HScrollBar scrlQuestStartChat 
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Label Label10 
         Caption         =   $"frmEditor_NPC.frx":3332
         Height          =   735
         Left            =   120
         TabIndex        =   67
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblQuestIndex 
         Caption         =   "Quest Index: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label Label9 
         Caption         =   "(Shown to the player next time they talk to the NPC after finishing the quest.) "
         Height          =   495
         Left            =   120
         TabIndex        =   64
         Top             =   4920
         Width           =   3495
      End
      Begin VB.Label lblPostQuestChat 
         Caption         =   "Post-quest chat: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   4320
         Width           =   3495
      End
      Begin VB.Label Label8 
         Caption         =   "(Shown to the player while they're on the quest. Gives them the option to hand it in.)"
         Height          =   615
         Left            =   120
         TabIndex        =   61
         Top             =   3720
         Width           =   3495
      End
      Begin VB.Label lblQuestInProgressChat 
         Caption         =   "Quest in-progress chat: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   3120
         Width           =   3495
      End
      Begin VB.Label Label6 
         Caption         =   "(Shown to the player if they haven't started or finished the quest yet. Gives them the option to start the quest.)"
         Height          =   615
         Left            =   120
         TabIndex        =   58
         Top             =   2400
         Width           =   3495
      End
      Begin VB.Label lblQuestStartChat 
         Caption         =   "Quest start chat: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1800
         Width           =   3495
      End
   End
   Begin VB.Frame fraDrop 
      Caption         =   "Drop (1/20)"
      Height          =   1815
      Left            =   120
      TabIndex        =   41
      Top             =   5040
      Width           =   4215
      Begin VB.TextBox txtChance 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2880
         TabIndex        =   45
         Text            =   "0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   2880
         Max             =   255
         TabIndex        =   44
         Top             =   1080
         Width           =   1215
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   43
         Top             =   1440
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDropIndex 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   42
         Top             =   240
         Value           =   1
         Width           =   3975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chance (percent):"
         Height          =   180
         Left            =   120
         TabIndex        =   48
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   1365
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "Item: None"
         Height          =   180
         Left            =   120
         TabIndex        =   47
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "Value: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stats"
      Height          =   1815
      Left            =   4440
      TabIndex        =   30
      Top             =   5040
      Width           =   2775
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   5
         Left            =   1440
         Max             =   255
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   4
         Left            =   120
         Max             =   255
         TabIndex        =   34
         Top             =   720
         Width           =   1215
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   3
         Left            =   120
         Max             =   255
         TabIndex        =   33
         Top             =   1200
         Width           =   1215
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   2
         Left            =   1440
         Max             =   255
         TabIndex        =   32
         Top             =   240
         Width           =   1215
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   1
         Left            =   120
         Max             =   255
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   1440
         TabIndex        =   40
         Top             =   960
         Width           =   480
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   465
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   1440
         TabIndex        =   37
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4320
      TabIndex        =   17
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "NPC Properties"
      Height          =   4815
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox chkQuestGiver 
         Caption         =   "Quest giver?"
         Height          =   255
         Left            =   3000
         TabIndex        =   55
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtAttackSay 
         Height          =   255
         Left            =   960
         TabIndex        =   54
         Top             =   600
         Width           =   3375
      End
      Begin VB.HScrollBar scrlFace 
         Height          =   255
         Left            =   2520
         TabIndex        =   52
         Top             =   2040
         Width           =   1815
      End
      Begin VB.HScrollBar scrlConv 
         Height          =   255
         Left            =   2520
         TabIndex        =   50
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   29
         Text            =   "0"
         Top             =   4320
         Width           =   1335
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   3120
         TabIndex        =   23
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox txtDamage 
         Height          =   285
         Left            =   960
         TabIndex        =   22
         Top             =   3600
         Width           =   1215
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   3960
         Width           =   1695
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3840
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   9
         Top             =   1080
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   8
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":33DC
         Left            =   1320
         List            =   "frmEditor_NPC.frx":33EF
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2400
         Width           =   3015
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   5
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblFace 
         Caption         =   "Face: "
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblConv 
         Caption         =   "Conversation: "
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Spawn Rate (seconds):"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   4440
         UseMnemonic     =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Damage:"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   3600
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   180
         Left            =   2520
         TabIndex        =   24
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Anim: None"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Behaviour:"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         UseMnemonic     =   0   'False
         Width           =   810
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         Caption         =   "Say:"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   345
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Exp:"
         Height          =   180
         Left            =   2520
         TabIndex        =   11
         Top             =   3240
         Width           =   345
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Health:"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "NPC List"
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.ListBox lstIndex 
         Height          =   4380
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkQuestGiver_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If chkQuestGiver.value = vbUnchecked Then
        Npc(EditorIndex).questGiver = 0
        Me.width = 7395
    ElseIf chkQuestGiver.value = vbChecked Then
        Npc(EditorIndex).questGiver = 1
        Me.width = 11265
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkQuestGiver_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbBehaviour_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Npc(EditorIndex).Behaviour = cmbBehaviour.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBehaviour_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearNPC EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlSprite.max = NumCharacters
    scrlAnimation.max = MAX_ANIMATIONS
    Me.width = 7395
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call NpcEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call NpcEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlAnimation.value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.value).name)
    lblAnimation.Caption = "Anim: " & sString
    Npc(EditorIndex).Animation = scrlAnimation.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlConv_Change()
    Npc(EditorIndex).Conv = scrlConv.value
    lblConv.Caption = "Conversation: " & scrlConv.value
End Sub

Private Sub scrlDropIndex_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    fraDrop.Caption = "Drop (" & scrlDropIndex.value & "/" & MAX_NPC_DROPS & ")"
    txtChance.Text = CStr(Npc(EditorIndex).DropChance(scrlDropIndex.value))
    scrlNum.value = Npc(EditorIndex).DropItem(scrlDropIndex.value)
    scrlValue.value = Npc(EditorIndex).DropItemValue(scrlDropIndex.value)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDropIndex_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFace_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Npc(EditorIndex).Face = scrlFace.value
    lblFace = "Face: " & scrlFace.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlFace_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlQuestIndex_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblQuestIndex.Caption = "Quest Index: " & scrlQuestIndex.value
    Npc(EditorIndex).questIndex = scrlQuestIndex.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlQuestIndex_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlQuestStartChat_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblQuestStartChat.Caption = "Quest start chat: " & scrlQuestStartChat.value
    Npc(EditorIndex).questStartChat = scrlQuestStartChat.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlQuestStartChat_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlQuestInProgressChat_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblQuestInProgressChat.Caption = "Quest in-progress chat: " & scrlQuestInProgressChat.value
    Npc(EditorIndex).questInProgressChat = scrlQuestInProgressChat.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlQuestInProgressChat_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPostQuestChat_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblPostQuestChat.Caption = "Post-quest chat: " & scrlPostQuestChat.value
    Npc(EditorIndex).postQuestChat = scrlPostQuestChat.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPostQuestChat_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSprite_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSprite.Caption = "Sprite: " & scrlSprite.value
    Call EditorNpc_BltSprite
    Npc(EditorIndex).Sprite = scrlSprite.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRange.Caption = "Range: " & scrlRange.value
    Npc(EditorIndex).Range = scrlRange.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNum_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlNum.value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(Item(scrlNum.value).name)
    End If
    
    Npc(EditorIndex).DropItem(scrlDropIndex.value) = scrlNum.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStat_Change(index As Integer)
Dim Prefix As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case index
        Case 1: Prefix = "Str: "
        Case 2: Prefix = "End: "
        Case 3: Prefix = "Int: "
        Case 4: Prefix = "Agi: "
        Case 5: Prefix = "Will: "
    End Select
    
    lblStat(index).Caption = Prefix & scrlStat(index).value
    Npc(EditorIndex).Stat(index) = scrlStat(index).value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStat_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlValue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblValue.Caption = "Value: " & scrlValue.value
    Npc(EditorIndex).DropItemValue(scrlDropIndex.value) = scrlValue.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlValue_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtAttackSay_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Npc(EditorIndex).AttackSay = txtAttackSay.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtAttackSay_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtChance_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtChance.Text) > 0 Then Exit Sub
    If IsNumeric(txtChance.Text) Then Npc(EditorIndex).DropChance(scrlDropIndex.value) = Val(txtChance.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtChance_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtDamage.Text) > 0 Then Exit Sub
    If IsNumeric(txtDamage.Text) Then Npc(EditorIndex).Damage = Val(txtDamage.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDamage_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtEXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtEXP.Text) > 0 Then Exit Sub
    If IsNumeric(txtEXP.Text) Then Npc(EditorIndex).EXP = Val(txtEXP.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtEXP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtHP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtHP.Text) > 0 Then Exit Sub
    If IsNumeric(txtHP.Text) Then Npc(EditorIndex).HP = Val(txtHP.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtHP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtLevel_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Len(txtLevel.Text) <= 0 Then Exit Sub
    If IsNumeric(txtLevel.Text) Then Npc(EditorIndex).Level = Val(txtLevel.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtlevel_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Npc(EditorIndex).name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtSpawnSecs_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtSpawnSecs.Text) > 0 Then Exit Sub
    Npc(EditorIndex).SpawnSecs = Val(txtSpawnSecs.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtSpawnSecs_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Npc(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Npc(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
