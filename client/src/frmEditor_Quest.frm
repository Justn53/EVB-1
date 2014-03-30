VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest Editor"
   ClientHeight    =   10395
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraRequirements 
      Caption         =   "Requirements"
      Height          =   1575
      Left            =   2640
      TabIndex        =   15
      Top             =   1080
      Width           =   5175
      Begin VB.HScrollBar scrlQuestReq 
         Height          =   255
         LargeChange     =   10
         Left            =   1320
         Max             =   99
         TabIndex        =   20
         Top             =   720
         Width           =   3735
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   315
         ItemData        =   "frmEditor_Quest.frx":0000
         Left            =   1320
         List            =   "frmEditor_Quest.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1080
         Width           =   3735
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   1320
         Max             =   99
         TabIndex        =   16
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label lblQuestReq 
         AutoSize        =   -1  'True
         Caption         =   "Quest req: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Class Req:"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level req: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5640
      TabIndex        =   8
      Top             =   9840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   9840
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   10215
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2415
      Begin VB.ListBox lstIndex 
         Height          =   9810
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   855
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Objectives"
      Height          =   6855
      Left            =   2640
      TabIndex        =   1
      Top             =   2760
      Width           =   5175
      Begin VB.HScrollBar scrlDataValue 
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   36
         Top             =   5400
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDataIndex 
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   35
         Top             =   5040
         Width           =   2895
      End
      Begin VB.ComboBox cmbObjectiveType 
         Height          =   315
         Index           =   3
         ItemData        =   "frmEditor_Quest.frx":0004
         Left            =   120
         List            =   "frmEditor_Quest.frx":0017
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   4680
         Width           =   4935
      End
      Begin VB.HScrollBar scrlDataValue 
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   31
         Top             =   3960
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDataIndex 
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   30
         Top             =   3600
         Width           =   2895
      End
      Begin VB.ComboBox cmbObjectiveType 
         Height          =   315
         Index           =   2
         ItemData        =   "frmEditor_Quest.frx":0045
         Left            =   120
         List            =   "frmEditor_Quest.frx":0058
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3240
         Width           =   4935
      End
      Begin VB.HScrollBar scrlDataValue 
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   25
         Top             =   2520
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDataIndex 
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   24
         Top             =   2160
         Width           =   2895
      End
      Begin VB.ComboBox cmbObjectiveType 
         Height          =   315
         Index           =   1
         ItemData        =   "frmEditor_Quest.frx":0086
         Left            =   120
         List            =   "frmEditor_Quest.frx":0099
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1800
         Width           =   4935
      End
      Begin VB.TextBox txtDescription 
         Height          =   855
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   4935
      End
      Begin VB.HScrollBar scrlXP 
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   6120
         Width           =   2175
      End
      Begin VB.HScrollBar scrlAmount 
         Height          =   255
         Left            =   840
         Min             =   1
         TabIndex        =   12
         Top             =   6480
         Value           =   1
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.HScrollBar scrlReward 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   6120
         Value           =   1
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Objective 1:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Objective 2:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Objective 3:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   5040
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Label lblDataValue 
         Caption         =   "Data Value: 0"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   38
         Top             =   5400
         Width           =   2055
      End
      Begin VB.Label lblDataIndex 
         Caption         =   "Data Index: 0"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   37
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label lblDataValue 
         Caption         =   "Data Value: 0"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label lblDataIndex 
         Caption         =   "Data Index: 0"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblDataValue 
         Caption         =   "Data Value: 0"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label lblDataIndex 
         Caption         =   "Data Index: 0"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label lblXP 
         Caption         =   "XP Reward: None"
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   5880
         Width           =   1695
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   6480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblReward 
         Caption         =   "Reward: None"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   5880
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   9840
      Width           =   1335
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbClassReq_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_QUESTS Then Exit Sub
    Quest(EditorIndex).ClassReq = cmbClassReq.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClassReq_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbObjectiveType_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Quest(lstIndex.ListIndex + 1).Objective(Index).ObjectiveType = cmbObjectiveType(Index).ListIndex
    
    scrlDataIndex(Index).Visible = False
    scrlDataValue(Index).Visible = False
    lblDataIndex(Index).Visible = False
    lblDataValue(Index).Visible = False
    
    Select Case cmbObjectiveType(Index).ListIndex
        Case 1 ' conversation
            scrlDataIndex(Index).Visible = True
            lblDataIndex(Index).Visible = True
            lblDataIndex(Index).Caption = "NPC Index: " & scrlDataIndex(Index).value
        Case 2 ' kill
            scrlDataIndex(Index).Visible = True
            scrlDataValue(Index).Visible = True
            lblDataIndex(Index).Visible = True
            lblDataValue(Index).Visible = True
            lblDataIndex(Index).Caption = "NPC Index: " & scrlDataIndex(Index).value
            lblDataValue(Index).Caption = "NPC Value: " & scrlDataValue(Index).value
        Case 3 ' item
            scrlDataIndex(Index).Visible = True
            scrlDataValue(Index).Visible = True
            lblDataIndex(Index).Visible = True
            lblDataValue(Index).Visible = True
            lblDataIndex(Index).Caption = "Item Index: " & scrlDataIndex(Index).value
            lblDataValue(Index).Caption = "Item Value: " & scrlDataValue(Index).value
        Case 4 ' resource
            scrlDataIndex(Index).Visible = True
            scrlDataValue(Index).Visible = True
            lblDataIndex(Index).Visible = True
            lblDataValue(Index).Visible = True
            lblDataIndex(Index).Caption = "Resource Index: " & scrlDataIndex(Index).value
            lblDataValue(Index).Caption = "Resource Value: " & scrlDataValue(Index).value
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbObjectiveType_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call QuestEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call QuestEditorOK
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Re-init the editor
    If QuestEditorLoaded = True Then QuestEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAmount_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblReward.Caption = "Reward: " & scrlAmount.value & "x " & Trim$(Item(scrlReward.value).name)
    Quest(lstIndex.ListIndex + 1).RewardAmount = scrlAmount.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAmount_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDataIndex_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblDataIndex(Index).Caption = "Data Index: " & scrlDataIndex(Index).value
    Quest(lstIndex.ListIndex + 1).Objective(Index).DataIndex = scrlDataIndex(Index).value

    Select Case cmbObjectiveType(Index).ListIndex
        Case 1 ' conversation
           lblDataIndex(Index).Caption = "NPC Index: " & scrlDataIndex(Index).value
        Case 2 ' kill
            lblDataIndex(Index).Caption = "NPC Index: " & scrlDataIndex(Index).value
        Case 3 ' item
            lblDataIndex(Index).Caption = "Item Index: " & scrlDataIndex(Index).value
        Case 4 ' resource
            lblDataIndex(Index).Caption = "Resource Index: " & scrlDataIndex(Index).value
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDataIndex_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDataValue_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblDataValue(Index).Caption = "Data Value: " & scrlDataValue(Index).value
    Quest(lstIndex.ListIndex + 1).Objective(Index).DataAmount = scrlDataValue(Index).value
    
    Select Case cmbObjectiveType(Index).ListIndex
        Case 2 ' kill
            lblDataValue(Index).Caption = "NPC Value: " & scrlDataValue(Index).value
        Case 3 ' item
            lblDataValue(Index).Caption = "Item Value: " & scrlDataValue(Index).value
        Case 4 ' resource
            lblDataIndex(Index).Caption = "Resource Index: " & scrlDataIndex(Index).value
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDataValue_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevelReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_QUESTS Then Exit Sub
    lblLevelReq.Caption = "Level req: " & scrlLevelReq
    Quest(EditorIndex).LevelReq = scrlLevelReq.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlQuestReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_QUESTS Then Exit Sub
    lblQuestReq.Caption = "Quest req: " & scrlQuestReq
    Quest(EditorIndex).QuestReq = scrlQuestReq.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlQuestReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlReward_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlReward.value <> 0 Then
        If Item(scrlReward.value).Type = ITEM_TYPE_CURRENCY Then
            scrlAmount.Visible = True
            lblAmount.Visible = True
            lblReward.Caption = "Reward: " & scrlAmount.value & "x " & Trim$(Item(scrlReward.value).name)
        Else
            lblReward.Caption = "Reward: " & Trim$(Item(scrlReward.value).name)
            Quest(lstIndex.ListIndex + 1).RewardAmount = 1
            scrlAmount.Visible = False
            lblAmount.Visible = False
        End If
    Else
        lblReward.Caption = "Reward: None"
    End If
    Quest(lstIndex.ListIndex + 1).Reward = scrlReward.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlReward_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlXP.value > 0 Then
        lblXP.Caption = "XP Reward: " & scrlXP.value
    Else
        lblXP.Caption = "XP Reward: None"
    End If
    
    Quest(EditorIndex).XPReward = scrlXP.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlXP_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDescription_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Quest(EditorIndex).Description = txtDescription.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDescription_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Quest(EditorIndex).name = txtName.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

