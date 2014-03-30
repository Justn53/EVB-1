VERSION 5.00
Begin VB.Form frmEditors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Editors"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmEditors.frx":0000
   ScaleHeight     =   230
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Image imgClose 
      Height          =   420
      Left            =   2235
      Top             =   2760
      Width           =   1545
   End
   Begin VB.Image imgEditor 
      Height          =   540
      Index           =   8
      Left            =   4020
      Top             =   2010
      Width           =   1545
   End
   Begin VB.Image imgEditor 
      Height          =   540
      Index           =   7
      Left            =   2235
      Top             =   2010
      Width           =   1545
   End
   Begin VB.Image imgEditor 
      Height          =   540
      Index           =   6
      Left            =   465
      Top             =   2010
      Width           =   1545
   End
   Begin VB.Image imgEditor 
      Height          =   540
      Index           =   5
      Left            =   4005
      Top             =   1230
      Width           =   1545
   End
   Begin VB.Image imgEditor 
      Height          =   540
      Index           =   4
      Left            =   2235
      Top             =   1230
      Width           =   1545
   End
   Begin VB.Image imgEditor 
      Height          =   540
      Index           =   3
      Left            =   465
      Top             =   1230
      Width           =   1545
   End
   Begin VB.Image imgEditor 
      Height          =   540
      Index           =   2
      Left            =   4005
      Top             =   465
      Width           =   1545
   End
   Begin VB.Image imgEditor 
      Height          =   540
      Index           =   1
      Left            =   2235
      Top             =   465
      Width           =   1545
   End
   Begin VB.Image imgEditor 
      Height          =   540
      Index           =   0
      Left            =   465
      Top             =   465
      Width           =   1545
   End
End
Attribute VB_Name = "frmEditors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub animEditor()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        MsgBox ("You need to be a developer to do this.")
        Exit Sub
    End If
    
    SendRequestEditAnimation
End Sub

Private Sub convEditor()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        MsgBox ("You need to be a developer to do this.")
        Exit Sub
    End If
    
    SendRequestEditConv
End Sub

Private Sub itemEditor()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        MsgBox ("You need to be a developer to do this.")
        Exit Sub
    End If
    
    SendRequestEditItem
End Sub

Private Sub mapEditor()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        MsgBox ("You need to be a mapper to do this.")
        Exit Sub
    End If
    
    SendRequestEditMap
End Sub

Private Sub npcEditor()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        MsgBox ("You need to be a developer to do this.")
        Exit Sub
    End If
    
    SendRequestEditNpc
End Sub

Private Sub questEditor()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        MsgBox ("You need to be a developer to do this.")
        Exit Sub
    End If
    
    SendRequestEditQuest
End Sub

Private Sub resourceEditor()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        MsgBox ("You need to be a developer to do this.")
        Exit Sub
    End If
    
    SendRequestEditResource
End Sub

Private Sub shopEditor()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        MsgBox ("You need to be a developer to do this.")
        Exit Sub
    End If
    
    SendRequestEditShop
End Sub

Private Sub spellEditor()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        MsgBox ("You need to be a developer to do this.")
        Exit Sub
    End If
    
    SendRequestEditSpell
End Sub

Private Sub imgEditor_Click(Index As Integer)
    Select Case Index
        Case 0: animEditor
        Case 1: convEditor
        Case 2: itemEditor
        Case 3: mapEditor
        Case 4: npcEditor
        Case 5: questEditor
        Case 6: resourceEditor
        Case 7: shopEditor
        Case 8: spellEditor
    End Select
End Sub

Private Sub imgClose_Click()
    Me.Visible = False
End Sub
