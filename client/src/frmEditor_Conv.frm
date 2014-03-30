VERSION 5.00
Begin VB.Form frmEditor_Conv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversation Editor"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   24
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4320
      TabIndex        =   23
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   20
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Frame fraConv 
      Caption         =   "Conversation (1/50)"
      Height          =   6015
      Left            =   2880
      TabIndex        =   7
      Top             =   1920
      Width           =   5655
      Begin VB.ComboBox cmbSound 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2160
         Width           =   5415
      End
      Begin VB.HScrollBar scrlData3 
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   5640
         Width           =   3350
      End
      Begin VB.HScrollBar scrlData2 
         Height          =   255
         Left            =   2160
         TabIndex        =   26
         Top             =   5280
         Width           =   3350
      End
      Begin VB.HScrollBar scrlData1 
         Height          =   255
         Left            =   2160
         TabIndex        =   25
         Top             =   4920
         Width           =   3350
      End
      Begin VB.ComboBox cmbEvent 
         Height          =   315
         ItemData        =   "frmEditor_Conv.frx":0000
         Left            =   120
         List            =   "frmEditor_Conv.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   4560
         Width           =   5415
      End
      Begin VB.ComboBox cmbToConv 
         Height          =   315
         Index           =   4
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3960
         Width           =   1455
      End
      Begin VB.ComboBox cmbToConv 
         Height          =   315
         Index           =   3
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3600
         Width           =   1455
      End
      Begin VB.ComboBox cmbToConv 
         Height          =   315
         Index           =   2
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3240
         Width           =   1455
      End
      Begin VB.ComboBox cmbToConv 
         Height          =   315
         Index           =   1
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   3960
         Width           =   3855
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   3600
         Width           =   3855
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   3855
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   3855
      End
      Begin VB.TextBox txtConvText 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   840
         Width           =   5415
      End
      Begin VB.HScrollBar scrlCurChat 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   8
         Top             =   240
         Value           =   1
         Width           =   5415
      End
      Begin VB.Label Label2 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblData3 
         Caption         =   "Data3: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   5640
         Width           =   1935
      End
      Begin VB.Label lblData2 
         Caption         =   "Data2: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   5280
         Width           =   1935
      End
      Begin VB.Label lblData1 
         Caption         =   "Data1: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Event:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Replies:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   1695
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.HScrollBar scrlChatCount 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   6
         Top             =   1200
         Value           =   1
         Width           =   5415
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label lblChatCount 
         Caption         =   "Chat count:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Name: "
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conversation List"
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.ListBox lstIndex 
         Height          =   7860
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmEditor_Conv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbEvent_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Make sure it's loaded
    If ConvEditorLoaded = False Then Exit Sub
    
    Conv(EditorIndex).Chat(scrlCurChat.value).Event = cmbEvent.ListIndex
    If ConvEditorLoaded Then InitEventData

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbEvent_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Make sure it's loaded
    If ConvEditorLoaded = False Then Exit Sub
    
    If cmbSound.ListIndex >= 0 Then
        Conv(EditorIndex).Chat(scrlCurChat.value).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Conv(EditorIndex).Chat(scrlCurChat.value).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbSound_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbToConv_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Make sure it's loaded
    If ConvEditorLoaded = False Then Exit Sub
    Conv(EditorIndex).Chat(scrlCurChat.value).ReplyConvTo(Index) = cmbToConv(Index).ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbToConv_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_CONVS Then Exit Sub
    ClearConv EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Conv(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Refresh the conv editor
    ConvEditorInit

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ConvEditorOK
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ConvEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Re-init the editor
    If ConvEditorLoaded = True Then
        ConvEditorLoaded = False
        ConvEditorInit
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlChatCount_Change()
Dim curIndex(1 To 4) As Byte, i As Long, j As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    If ConvEditorLoaded = False Then Exit Sub
    
    Conv(EditorIndex).ChatCount = scrlChatCount.value
    scrlCurChat.max = Conv(EditorIndex).ChatCount
    lblChatCount.Caption = "Chat count: " & scrlChatCount.value
    fraConv.Caption = "Conversation: (" & scrlCurChat.value & "/" & scrlChatCount.value & ")"
    
    ' Reset the conv-to boxes
    For i = 1 To 4
        If cmbToConv(i).ListIndex > 0 Then
            curIndex(i) = cmbToConv(i).ListIndex
        Else
            curIndex(i) = 0
        End If
        
        cmbToConv(i).Clear
        cmbToConv(i).AddItem "None", 0
        
        For j = 1 To scrlChatCount.value
            cmbToConv(i).AddItem CStr(j), j
        Next
        
        ' reset the list index
        If curIndex(i) > scrlChatCount.value Then curIndex(i) = scrlChatCount.value
        cmbToConv(i).ListIndex = curIndex(i)
    Next
    
    ' Reset the data
    If Conv(EditorIndex).Chat(scrlChatCount.value).Data1 = 0 Then Conv(EditorIndex).Chat(scrlChatCount.value).Data1 = 1
    If Conv(EditorIndex).Chat(scrlChatCount.value).Data2 = 0 Then Conv(EditorIndex).Chat(scrlChatCount.value).Data2 = 1
    If Conv(EditorIndex).Chat(scrlChatCount.value).Data3 = 0 Then Conv(EditorIndex).Chat(scrlChatCount.value).Data3 = 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlChatCount_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCurChat_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    If ConvEditorLoaded = False Then Exit Sub
    
    fraConv.Caption = "Conversation: (" & scrlCurChat.value & "/" & Conv(EditorIndex).ChatCount & ")"
    ConvEditorLoaded = False
    ConvEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlData1_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Make sure it's loaded
    If ConvEditorLoaded = False Then Exit Sub
   
    ' Change it based on what it is
    Select Case cmbEvent.ListIndex
        Case 2
            Conv(EditorIndex).Chat(scrlCurChat.value).Data1 = scrlData1.value
            lblData1.Caption = "Shop Index: " & scrlData1.value
        Case 3, 4
            Conv(EditorIndex).Chat(scrlCurChat.value).Data1 = scrlData1.value
            lblData1.Caption = "Item Index: " & scrlData1.value
        Case 5
            Conv(EditorIndex).Chat(scrlCurChat.value).Data1 = scrlData1.value
            lblData1.Caption = "Map Index: " & scrlData1.value
        Case 6
            Conv(EditorIndex).Chat(scrlCurChat.value).Data1 = scrlData1.value
            lblData1.Caption = "Heal amount: " & scrlData1.value
        Case 7
            Conv(EditorIndex).Chat(scrlCurChat.value).Data1 = scrlData1.value
            lblData1.Caption = "Quest Index: " & scrlData1.value
        Case 8
            Conv(EditorIndex).Chat(scrlCurChat.value).Data1 = scrlData1.value
            lblData1.Caption = "Quest Index: " & scrlData1.value
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlData1_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlData2_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
   
    ' Make sure it's loaded
    If ConvEditorLoaded = False Then Exit Sub
   
    ' Change it based on what it is
    Select Case cmbEvent.ListIndex
        Case 3, 4
            Conv(EditorIndex).Chat(scrlCurChat.value).Data2 = scrlData2.value
            lblData2.Caption = "Value: " & scrlData2.value
        Case 5
            Conv(EditorIndex).Chat(scrlCurChat.value).Data2 = scrlData2.value
            lblData2.Caption = "X: " & scrlData2.value
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlData2_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlData3_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Make sure it's loaded
    If ConvEditorLoaded = False Then Exit Sub
   
    ' Change it based on what it is
    Select Case cmbEvent.ListIndex
        Case 5
            Conv(EditorIndex).Chat(scrlCurChat.value).Data3 = scrlData3.value
            lblData3.Caption = "Y: " & scrlData3.value
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlData3_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtConvText_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Conv(EditorIndex).Chat(scrlCurChat.value).Text = Trim$(txtConvText.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtConvText_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Conv(EditorIndex).name = txtName.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtReply_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Conv(EditorIndex).Chat(scrlCurChat.value).ReplyText(Index) = txtReply(Index).Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Change", "frmEditor_Conv", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
