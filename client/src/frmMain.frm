VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   14070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15180
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   938
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1012
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00B5B5B5&
      ForeColor       =   &H80000008&
      Height          =   7410
      Left            =   12000
      ScaleHeight     =   492
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   2865
      Begin VB.CommandButton cmdAllEditors 
         Caption         =   "Game Editors"
         Height          =   255
         Left            =   240
         TabIndex        =   115
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CommandButton cmdSSMap 
         Caption         =   "Screenshot Map"
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Top             =   6960
         Width           =   2295
      End
      Begin VB.CommandButton cmdLevel 
         Caption         =   "Level Up"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   6600
         Width           =   2295
      End
      Begin VB.CommandButton cmdAAccess 
         Caption         =   "Set Access"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtAAccess 
         Height          =   285
         Left            =   1440
         TabIndex        =   38
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtASprite 
         Height          =   285
         Left            =   2160
         TabIndex        =   36
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton cmdARespawn 
         Caption         =   "Respawn"
         Height          =   255
         Left            =   1440
         TabIndex        =   35
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton cmdASprite 
         Caption         =   "Set Sprite"
         Height          =   255
         Left            =   1440
         TabIndex        =   34
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Spawn Item"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   6120
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAAmount 
         Height          =   255
         Left            =   240
         Min             =   1
         TabIndex        =   32
         Top             =   5760
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   240
         Min             =   1
         TabIndex        =   30
         Top             =   5160
         Value           =   1
         Width           =   2295
      End
      Begin VB.CommandButton cmdADestroy 
         Caption         =   "Del Bans"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMapReport 
         Caption         =   "Map Report"
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdALoc 
         Caption         =   "Loc"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Warp To"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtAMap 
         Height          =   285
         Left            =   960
         TabIndex        =   22
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton cmdAWarpMe2 
         Caption         =   "WarpMe2"
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp2Me 
         Caption         =   "Warp2Me"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Ban"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Kick"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtAName 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.Line Line5 
         X1              =   16
         X2              =   168
         Y1              =   432
         Y2              =   432
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Access:"
         Height          =   255
         Left            =   1440
         TabIndex        =   39
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite#:"
         Height          =   255
         Left            =   1440
         TabIndex        =   37
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: 1"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Item: None"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Line Line4 
         X1              =   16
         X2              =   168
         Y1              =   384
         Y2              =   384
      End
      Begin VB.Line Line3 
         X1              =   16
         X2              =   168
         Y1              =   264
         Y2              =   264
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Editors:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Line Line2 
         X1              =   16
         X2              =   168
         Y1              =   216
         Y2              =   216
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Map#:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   16
         X2              =   168
         Y1              =   160
         Y2              =   160
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Panel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   120
         Width           =   2865
      End
   End
   Begin VB.PictureBox picQuests 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8115
      ScaleHeight     =   4050
      ScaleWidth      =   2910
      TabIndex        =   111
      Top             =   4245
      Visible         =   0   'False
      Width           =   2910
      Begin VB.ListBox lstQuests 
         Height          =   2595
         Left            =   240
         TabIndex        =   112
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblDrop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Drop"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   600
         TabIndex        =   114
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Info"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   600
         TabIndex        =   113
         Top             =   3360
         Width           =   1695
      End
   End
   Begin VB.PictureBox picConv 
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      Height          =   2070
      Left            =   180
      ScaleHeight     =   2070
      ScaleWidth      =   7140
      TabIndex        =   103
      Top             =   6630
      Visible         =   0   'False
      Width           =   7140
      Begin VB.PictureBox picConvFace 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   120
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   98
         TabIndex        =   104
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label lblConvResp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000015&
         BackStyle       =   0  'Transparent
         Caption         =   "Reply4"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   6360
         TabIndex        =   110
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblConvResp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000015&
         BackStyle       =   0  'Transparent
         Caption         =   "Reply3"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   6360
         TabIndex        =   109
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblConvResp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000015&
         BackStyle       =   0  'Transparent
         Caption         =   "Reply2"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   6360
         TabIndex        =   108
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblConvResp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reply1"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   6360
         TabIndex        =   107
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblConvMessage 
         BackColor       =   &H80000006&
         BackStyle       =   0  'Transparent
         Caption         =   "This is a random block of text. Lets test to see how far we can go before breaking it. Okay. Cool."
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   1680
         TabIndex        =   106
         Top             =   480
         Width           =   5415
      End
      Begin VB.Label lblConvNPCName 
         BackColor       =   &H80000015&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   105
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.PictureBox picItemDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   0
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   6
      Top             =   9120
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picItemDescPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1095
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   80
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblItemDesc 
         BackStyle       =   0  'Transparent
         Caption         =   """This is an example of an item's description. It  can be quite big, so we have to keep it at a decent size."""
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1530
         Left            =   240
         TabIndex        =   79
         Top             =   1800
         Width           =   2640
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   7
         Top             =   210
         Width           =   2805
      End
   End
   Begin VB.PictureBox picSpellDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   3240
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   55
      Top             =   9120
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picSpellDescPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1095
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   85
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblSpellDesc 
         BackStyle       =   0  'Transparent
         Caption         =   """This is an example of an item's description. It  can be quite big, so we have to keep it at a decent size."""
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1530
         Left            =   240
         TabIndex        =   84
         Top             =   1800
         Width           =   2640
      End
      Begin VB.Label lblSpellName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   83
         Top             =   210
         Width           =   2805
      End
   End
   Begin VB.PictureBox picDialogue 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   6480
      ScaleHeight     =   2085
      ScaleWidth      =   7140
      TabIndex        =   90
      Top             =   11880
      Visible         =   0   'False
      Width           =   7140
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Okay"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   3285
         TabIndex        =   95
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label lblDialogue_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Robin has requested a trade. Would you like to accept?"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   94
         Top             =   720
         Width           =   6615
      End
      Begin VB.Label lblDialogue_Title 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Trade Request"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   480
         Width           =   6855
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   3375
         TabIndex        =   92
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   3405
         TabIndex        =   91
         Top             =   1560
         Width           =   285
      End
   End
   Begin VB.PictureBox picCurrency 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   6480
      ScaleHeight     =   2085
      ScaleWidth      =   7140
      TabIndex        =   48
      Top             =   9720
      Visible         =   0   'False
      Width           =   7140
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   50
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblCurrencyCancel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3240
         TabIndex        =   52
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblCurrencyOk 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Okay"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3300
         TabIndex        =   51
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblCurrency 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "How many do you want to drop?"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   49
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.PictureBox picTempInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   6480
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   4
      Top             =   9120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTempBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   7080
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   66
      Top             =   9120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTempSpell 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   7680
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   87
      Top             =   9120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picParty 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8115
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   96
      Top             =   4245
      Visible         =   0   'False
      Width           =   2910
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   4
         Left            =   90
         Top             =   3075
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   4
         Left            =   90
         Top             =   2940
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   3
         Left            =   90
         Top             =   2340
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   3
         Left            =   90
         Top             =   2205
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   2
         Left            =   90
         Top             =   1620
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   2
         Left            =   90
         Top             =   1485
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartySpirit 
         Height          =   135
         Index           =   1
         Left            =   90
         Top             =   870
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Image imgPartyHealth 
         Height          =   135
         Index           =   1
         Left            =   90
         Top             =   735
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Label lblPartyLeave 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   102
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblPartyInvite 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   101
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   100
         Top             =   2670
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   99
         Top             =   1935
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   98
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lblPartyMember 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   97
         Top             =   465
         Width           =   2415
      End
   End
   Begin VB.PictureBox picSSMap 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   12360
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   89
      Top             =   8640
      Width           =   255
   End
   Begin VB.PictureBox picCover 
      Appearance      =   0  'Flat
      BackColor       =   &H00181C21&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   12000
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   82
      Top             =   8640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picHotbar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   180
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   476
      TabIndex        =   81
      Top             =   5985
      Width           =   7140
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1800
      Left            =   180
      TabIndex        =   1
      Top             =   6630
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   3175
      _Version        =   393217
      BackColor       =   790032
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":3332
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtMyChat 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   720
      TabIndex        =   2
      Top             =   8475
      Width           =   6600
   End
   Begin VB.PictureBox picSpells 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8115
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   47
      Top             =   4245
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8115
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   56
      Top             =   4245
      Visible         =   0   'False
      Width           =   2910
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   62
         Top             =   1440
         Width           =   1935
         Begin VB.OptionButton optSOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   64
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optSOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   63
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   59
         Top             =   840
         Width           =   1935
         Begin VB.OptionButton optMOff 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Off"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   61
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optMOn 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "On"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   60
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   58
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Music"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   240
         TabIndex        =   57
         Top             =   600
         Width           =   555
      End
   End
   Begin VB.PictureBox picInventory 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8115
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   3
      Top             =   4245
      Width           =   2895
   End
   Begin VB.PictureBox picCharacter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   8115
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   194
      TabIndex        =   5
      Top             =   4245
      Visible         =   0   'False
      Width           =   2910
      Begin VB.PictureBox picFace 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   735
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   78
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label lblPoints 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2250
         TabIndex        =   86
         Top             =   2970
         Width           =   120
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   1440
         TabIndex        =   45
         Top             =   2955
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   5
         Left            =   2550
         TabIndex        =   44
         Top             =   2730
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   1440
         TabIndex        =   43
         Top             =   2730
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   2550
         TabIndex        =   42
         Top             =   2505
         Width           =   105
      End
      Begin VB.Label lblTrainStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   1440
         TabIndex        =   41
         Top             =   2505
         Width           =   105
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   1140
         TabIndex        =   13
         Top             =   2970
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   5
         Left            =   2250
         TabIndex        =   12
         Top             =   2760
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   1140
         TabIndex        =   11
         Top             =   2760
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   4
         Left            =   2250
         TabIndex        =   10
         Top             =   2535
         Width           =   120
      End
      Begin VB.Label lblCharStat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   1140
         TabIndex        =   9
         Top             =   2535
         Width           =   120
      End
      Begin VB.Label lblCharName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   8
         Top             =   495
         Width           =   2640
      End
   End
   Begin VB.PictureBox picTrade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   150
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   67
      Top             =   150
      Visible         =   0   'False
      Width           =   7200
      Begin VB.PictureBox picTheirTrade 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   3855
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   69
         Top             =   465
         Width           =   2895
      End
      Begin VB.PictureBox picYourTrade 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   435
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   68
         Top             =   465
         Width           =   2895
      End
      Begin VB.Image imgDeclineTrade 
         Height          =   435
         Left            =   3675
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Image imgAcceptTrade 
         Height          =   435
         Left            =   2475
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Label lblTradeStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   600
         TabIndex        =   72
         Top             =   5520
         Width           =   5895
      End
      Begin VB.Label lblTheirWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   71
         Top             =   4500
         Width           =   1815
      End
      Begin VB.Label lblYourWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   70
         Top             =   4500
         Width           =   1815
      End
   End
   Begin VB.PictureBox picBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   150
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   65
      Top             =   150
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.PictureBox picShop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5115
      Left            =   1680
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   53
      Top             =   480
      Visible         =   0   'False
      Width           =   4125
      Begin VB.PictureBox picShopItems 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3165
         Left            =   615
         ScaleHeight     =   211
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   54
         Top             =   630
         Width           =   2895
      End
      Begin VB.Image imgLeaveShop 
         Height          =   435
         Left            =   2715
         Top             =   4350
         Width           =   1035
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00181C21&
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
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   150
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   7200
      Begin MSWinsockLib.Winsock Socket 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   7
      Left            =   9000
      Picture         =   "frmMain.frx":33AE
      Top             =   8400
      Width           =   1035
   End
   Begin VB.Label lblEXP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9480
      TabIndex        =   77
      Top             =   1080
      Width           =   1845
   End
   Begin VB.Label lblMP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9480
      TabIndex        =   76
      Top             =   750
      Width           =   1845
   End
   Begin VB.Label lblHP 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   9480
      TabIndex        =   75
      Top             =   420
      Width           =   1845
   End
   Begin VB.Image imgEXPBar 
      Height          =   240
      Left            =   7770
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Image imgMPBar 
      Height          =   240
      Left            =   7770
      Top             =   750
      Width           =   3615
   End
   Begin VB.Image imgHPBar 
      Height          =   240
      Left            =   7770
      Top             =   420
      Width           =   3615
   End
   Begin VB.Label lblPing 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8520
      TabIndex        =   74
      Top             =   1920
      Width           =   450
   End
   Begin VB.Label lblGold 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0g"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8520
      TabIndex        =   73
      Top             =   1515
      Width           =   225
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   6
      Left            =   10245
      Top             =   3450
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   5
      Left            =   9045
      Top             =   3450
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   4
      Left            =   7845
      Top             =   3450
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   3
      Left            =   10245
      Top             =   2850
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   2
      Left            =   9045
      Top             =   2850
      Width           =   1035
   End
   Begin VB.Image imgButton 
      Height          =   435
      Index           =   1
      Left            =   7845
      Top             =   2850
      Width           =   1035
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' **********
' * Events *
' **********
Private MoveForm As Boolean
Private MouseX As Long
Private MouseY As Long
Private PresentX As Long
Private PresentY As Long

Private Sub cmdAllEditors_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestAllEditors
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAllEditors_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdLevel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestLevelUp
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSSMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' render the map temp
    ScreenshotMap
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' move GUI
    Me.width = 11865
    Me.height = 9345
    picAdmin.Left = 544
    picCurrency.Left = txtChat.Left
    picCurrency.top = txtChat.top
    picDialogue.top = txtChat.top
    picDialogue.Left = txtChat.Left
    picCover.top = picScreen.top - 1
    picCover.Left = picScreen.Left - 1
    picCover.height = picScreen.height + 2
    picCover.width = picScreen.width + 2
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Cancel = True
    logoutGame
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' reset all buttons
    resetButtons_Main
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgAcceptTrade_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    AcceptTrade
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgAcceptTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_Click(index As Integer)
Dim Buffer As clsBuffer
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case index
        Case 1
            If Not picInventory.Visible Then
                ' show the window
                picInventory.Visible = True
                picCharacter.Visible = False
                picSpells.Visible = False
                picOptions.Visible = False
                picParty.Visible = False
                picQuests.Visible = False
                BltInventory
            End If
        Case 2
            If Not picSpells.Visible Then
                ' send packet
                Set Buffer = New clsBuffer
                Buffer.WriteLong CSpells
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                ' show the window
                picSpells.Visible = True
                picInventory.Visible = False
                picCharacter.Visible = False
                picOptions.Visible = False
                picParty.Visible = False
                picQuests.Visible = False
            End If
        Case 3
            If Not picCharacter.Visible Then
                ' send packet
                SendRequestPlayerData
                ' show the window
                picCharacter.Visible = True
                picInventory.Visible = False
                picSpells.Visible = False
                picOptions.Visible = False
                picParty.Visible = False
                picQuests.Visible = False
                
                ' Render
                BltEquipment
                BltFace
            End If
        Case 4
            If Not picOptions.Visible Then
                ' show the window
                picCharacter.Visible = False
                picInventory.Visible = False
                picSpells.Visible = False
                picOptions.Visible = True
                picParty.Visible = False
                picQuests.Visible = False
            End If
        Case 5
            If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                SendTradeRequest
            Else
                AddText "Invalid trade target.", BrightRed
            End If
        Case 6
            ' show the window
            picCharacter.Visible = False
            picInventory.Visible = False
            picSpells.Visible = False
            picOptions.Visible = False
            picQuests.Visible = False
            picParty.Visible = True
            picQuests.Visible = False
        Case 7
            ' show the window
            picCharacter.Visible = False
            picInventory.Visible = False
            picSpells.Visible = False
            picOptions.Visible = False
            picQuests.Visible = False
            picParty.Visible = False
            picQuests.Visible = True
            
            ' populate the quest list
            If picQuests.Visible Then
                lstQuests.Clear
                
                For i = 1 To MAX_QUESTS
                    If Player(MyIndex).Quest(i).QuestStatus = 1 Then
                        lstQuests.AddItem Quest(i).name
                    End If
                Next
            End If
    End Select
    
    ' Play the button-click sound
    If index > 0 Then PlaySound "Decision1.wav"
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' reset other buttons
    resetButtons_Main index
    
    ' change the button we're hovering on
    If Not MainButton(index).state = 2 Then ' make sure we're not clicking
        changeButtonState_Main index, 1 ' hover
    End If
    
    ' play sound
    If Not LastButtonSound_Main = index Then
        PlaySound "Cursor1.wav"
        LastButtonSound_Main = index
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
        
    ' reset all buttons
    resetButtons_Main -1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' reset other buttons
    resetButtons_Main index
    
    ' change the button we're hovering on
    changeButtonState_Main index, 2 ' clicked
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgButton_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblConvResp_Click(index As Integer)
    Call ConvResp(index)
End Sub

Private Sub lblCurrencyCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' clear everything
    picCurrency.Visible = False
    txtCurrency.Text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCurrencyCancel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgDeclineTrade_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DeclineTrade
    frmMain.picTrade.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgDeclineTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgLeaveShop_Click()
Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CCloseShop
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    picCover.Visible = False
    picShop.Visible = False
    InShop = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "imgLeaveShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblCurrencyOk_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsNumeric(txtCurrency.Text) Then
        If Val(txtCurrency.Text) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then txtCurrency.Text = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
        Select Case CurrencyMenu
            Case 1 ' drop item
                SendDropItem tmpCurrencyItem, Val(txtCurrency.Text)
            Case 2 ' deposit item
                DepositItem tmpCurrencyItem, Val(txtCurrency.Text)
            Case 3 ' withdraw item
                WithdrawItem tmpCurrencyItem, ValCurrency
            Case 4 ' offer trade item
                TradeItem tmpCurrencyItem, Val(txtCurrency.Text)
        End Select
    Else
        AddText "Please enter a valid amount.", BrightRed
        Exit Sub
    End If
    
    picCurrency.Visible = False
    tmpCurrencyItem = 0
    txtCurrency.Text = vbNullString
    CurrencyMenu = 0 ' clear
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblCurrencyOk_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblDialogue_Button_Click(index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' call the handler
    dialogueHandler index
    
    picDialogue.Visible = False
    dialogueIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblDialogue_Button_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblDrop_Click()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Make sure they got an index
    If lstQuests.ListIndex < 0 Then
        AddText "You need to select a quest!", BrightRed
        Exit Sub
    End If
    
    ' Info request
    SendQuestDrop lstQuests.List(lstQuests.ListIndex)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblDrop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblInfo_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Make sure they got an index
    If lstQuests.ListIndex < 0 Then
        AddText "You need to select a quest!", BrightRed
        Exit Sub
    End If
    
    ' Info request
    SendRequestQuestInfo lstQuests.List(lstQuests.ListIndex)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblInfo_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblPartyInvite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
        SendPartyRequest
    Else
        AddText "Invalid invitation target.", BrightRed
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblPartyLeave_Click()
        ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Party.Leader > 0 Then
        SendPartyLeave
    Else
        AddText "You are not in a party.", BrightRed
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblTrainStat_Click(index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
    SendTrainStat index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lblTrainStat_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optMOff_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Music = 0
    ' stop music playing
    StopMusic
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optMOn_Click()
Dim MusicFile As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' start the music
    Options.Music = 1
    MusicFile = Trim$(Map.Music)
    
    If Not MusicFile = "None." Then
        PlayMusic MusicFile
    Else
        StopMusic
    End If
    
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optMOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSOff_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Sound = 0
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSOff_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSOn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Options.Sound = 1
    ' save to config.ini
    SaveOptions
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "optSOn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picCover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' reset all buttons
    resetButtons_Main
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCover_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picHotbar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim SlotNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SlotNum = IsHotbarSlot(x, y)

    If Button = 1 Then
        If SlotNum <> 0 Then
            If SpellBuffer = SlotNum Then Exit Sub
            SendHotbarUse SlotNum
        End If
    ElseIf Button = 2 Then
        If SlotNum <> 0 Then
            SendHotbarChange 0, 0, SlotNum
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picHotbar_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picHotbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim SlotNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    SlotNum = IsHotbarSlot(x, y)

    If SlotNum <> 0 Then
        If Hotbar(SlotNum).sType = 1 Then ' item
            x = x + picHotbar.Left + 1
            y = y + picHotbar.top - picItemDesc.height - 1
            UpdateDescWindow Hotbar(SlotNum).Slot, x, y
            LastItemDesc = Hotbar(SlotNum).Slot ' set it so you don't re-set values
            Exit Sub
        ElseIf Hotbar(SlotNum).sType = 2 Then ' spell
            x = x + picHotbar.Left + 1
            y = y + picHotbar.top - picSpellDesc.height - 1
            UpdateSpellWindow Hotbar(SlotNum).Slot, x, y
            LastSpellDesc = Hotbar(SlotNum).Slot  ' set it so you don't re-set values
            Exit Sub
        End If
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' no spell was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picHotbar_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InMapEditor Then
        Call MapEditorMouseDown(Button, x, y, False)
    Else
        ' left click
        If Button = vbLeftButton Then
            ' targetting
            Call PlayerSearch(CurX, CurY)
        ' right click
        ElseIf Button = vbRightButton Then
            If ShiftDown Then
                ' admin warp if we're pressing shift and right clicking
                If GetPlayerAccess(MyIndex) >= 2 Then AdminWarp CurX, CurY
            End If
        End If
    End If

    Call SetFocusOnChat
    txtMyChat.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picScreen_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    CurX = TileView.Left + ((x + Camera.Left) \ PIC_X)
    CurY = TileView.top + ((y + Camera.top) \ PIC_Y)

    If InMapEditor Then
        frmEditor_Map.shpLoc.Visible = False

        If Button = vbLeftButton Or Button = vbRightButton Then
            Call MapEditorMouseDown(Button, x, y)
        End If
    End If
    
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' reset all buttons
    resetButtons_Main
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picScreen_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsShopItem(ByVal x As Single, ByVal y As Single) As Long
Dim tempRec As RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsShopItem = 0

    For i = 1 To MAX_TRADES
        If Shop(InShop).TradeItem(i).Item > 0 And Shop(InShop).TradeItem(i).Item <= MAX_ITEMS Then
            With tempRec
                .top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                .Bottom = .top + PIC_Y
                .Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.top And y <= tempRec.Bottom Then
                    IsShopItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsShopItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub picShop_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' reset all buttons
    resetButtons_Main
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picShop_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picShopItems_DblClick()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ShopItem > 0 Then
        BuyItem ShopItem
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picShopItems_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picShopItems_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x2 As Long, y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ShopItem = IsShopItem(x, y)

    If ShopItem > 0 Then
        x2 = x + picShop.Left + picShopItems.Left + 1
        y2 = y + picShop.top + picShopItems.top + 1
        UpdateDescWindow Shop(InShop).TradeItem(ShopItem).Item, x2, y2
        LastItemDesc = Shop(InShop).TradeItem(ShopItem).Item
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picShopItems_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpellDesc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picSpellDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpellDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_DblClick()
Dim spellnum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spellnum = IsPlayerSpell(SpellX, SpellY)

    If spellnum <> 0 Then
        Call CastSpell(spellnum)
        Exit Sub
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim spellnum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spellnum = IsPlayerSpell(SpellX, SpellY)
    If Button = 1 Then ' left click
        If spellnum <> 0 Then
            DragSpell = spellnum
            Exit Sub
        End If
    ElseIf Button = 2 Then ' right click
        If spellnum <> 0 Then
            Dialogue "Forget Spell", "Are you sure you want to forget how to cast " & Trim$(Spell(PlayerSpells(spellnum)).name) & "?", DIALOGUE_TYPE_FORGET, True, spellnum
            Exit Sub
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim spellSlot As Long
Dim x2 As Long, y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellX = x
    SpellY = y
    
    spellSlot = IsPlayerSpell(x, y)
    
    If DragSpell > 0 Then
        Call BltDraggedSpell(x + picSpells.Left, y + picSpells.top)
    Else
        If spellSlot <> 0 Then
            x2 = x + picSpells.Left - picSpellDesc.width - 1
            y2 = y + picSpells.top - picSpellDesc.height - 1
            UpdateSpellWindow PlayerSpells(spellSlot), x2, y2
            LastSpellDesc = PlayerSpells(spellSlot)
            Exit Sub
        End If
    End If
    
    picSpellDesc.Visible = False
    LastSpellDesc = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picSpells_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
Dim recPos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DragSpell > 0 Then
        ' drag + drop
        For i = 1 To MAX_PLAYER_SPELLS
            With recPos
                .top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= recPos.Left And x <= recPos.Right Then
                If y >= recPos.top And y <= recPos.Bottom Then
                    If DragSpell <> i Then
                        SendChangeSpellSlots DragSpell, i
                        Exit For
                    End If
                End If
            End If
        Next
        ' hotbar
        For i = 1 To MAX_HOTBAR
            With recPos
                .top = picHotbar.top - picSpells.top
                .Left = picHotbar.Left - picSpells.Left + (HotbarOffsetX * (i - 1)) + (32 * (i - 1))
                .Right = .Left + 32
                .Bottom = picHotbar.top - picSpells.top + 32
            End With
            
            If x >= recPos.Left And x <= recPos.Right Then
                If y >= recPos.top And y <= recPos.Bottom Then
                    SendHotbarChange 2, DragSpell, i
                    DragSpell = 0
                    picTempSpell.Visible = False
                    Exit Sub
                End If
            End If
        Next
    End If

    DragSpell = 0
    picTempSpell.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picSpells_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picTrade_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picYourTrade_DblClick()
Dim TradeNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeNum = IsTradeItem(TradeX, TradeY, True)
    If TradeNum <> 0 Then UntradeItem TradeNum
      
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picYourTrade_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picYourTrade_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim TradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeX = x
    TradeY = y
    
    TradeNum = IsTradeItem(x, y, True)
    
    If TradeNum <> 0 Then
        x = x + picTrade.Left + picYourTrade.Left + 4
        y = y + picTrade.top + picYourTrade.top + 4
        UpdateDescWindow GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).num), x, y
        LastItemDesc = GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).num) ' set it so you don't re-set values
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picYourTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picTheirTrade_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim TradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeNum = IsTradeItem(x, y, False)
    
    If TradeNum <> 0 Then
        x = x + picTrade.Left + picTheirTrade.Left + 4
        y = y + picTrade.top + picTheirTrade.top + 4
        UpdateDescWindow TradeTheirOffer(TradeNum).num, x, y
        LastItemDesc = TradeTheirOffer(TradeNum).num ' set it so you don't re-set values
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picTheirTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAAmount_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAAmount.Caption = "Amount: " & scrlAAmount.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAAmount_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAItem_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAItem.Caption = "Item: " & Trim$(Item(scrlAItem.value).name)
    If Item(scrlAItem.value).Type = ITEM_TYPE_CURRENCY Then
        scrlAAmount.Enabled = True
        Exit Sub
    End If
    
    scrlAAmount.Enabled = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAItem_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Triggers when we receive data
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsConnected Then Call IncomingData(bytesTotal)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call HandleKeyPresses(KeyAscii)

    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case KeyCode
        Case vbKeyInsert
            If Player(MyIndex).Access > 0 Then
                picAdmin.Visible = Not picAdmin.Visible
            End If
        Case vbKeyReturn
            ' Check to see if we're trying to enable the chatbox
            If frmMain.txtMyChat.Visible = True Then
                frmMain.txtMyChat.Visible = False
            Else
                frmMain.txtMyChat.Text = vbNullString
                frmMain.txtMyChat.Visible = True
                SetFocusOnChat
            End If
    End Select
    
    ' make sure they're not trying to input other stuff
    If InTrade Or InShop Or InBank Then Exit Sub
    If picCurrency.Visible Then Exit Sub
    
    ' hotbar keys
    Select Case KeyCode
        Case 48: ' 0
            SendHotbarUse 10
            Exit Sub
        Case 187: ' -
            SendHotbarUse 12
            Exit Sub
        Case 189: ' =
            SendHotbarUse 11
            Exit Sub
    End Select
    
    ' number keys
    For i = 1 To MAX_HOTBAR - 3
       If KeyCode = 48 + i Then
            SendHotbarUse i
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtChat_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    txtMyChat.Text = vbNullString
    txtMyChat.Visible = True
    SetFocusOnChat
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtChat_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtCurrency_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' fix for withdrawing currencies
    If txtCurrency.Text <> vbNullString And txtCurrency.Text <> "0" Then
        ValCurrency = Val(txtCurrency.Text)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtCurrency_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMyChat_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    MyText = txtMyChat
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMyChat_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtChat_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SetFocusOnChat
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtChat_GotFocus", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' *************
' * Inventory *
' *************
Private Sub picInventory_DblClick()
Dim InvNum As Long
Dim value As Long
Dim multiplier As Double
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DragInvSlotNum = 0
    InvNum = IsInvItem(InvX, InvY)

    If InvNum <> 0 Then
    
        ' are we in a shop?
        If InShop > 0 Then
            SellItem InvNum
            Exit Sub
        End If
        
        ' in bank?
        If InBank Then
            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                CurrencyMenu = 2 ' deposit
                lblCurrency.Caption = "How many do you want to deposit?"
                tmpCurrencyItem = InvNum
                txtCurrency.Text = vbNullString
                picCurrency.Visible = True
                txtCurrency.SetFocus
                Exit Sub
            End If
                
            Call DepositItem(InvNum, 0)
            Exit Sub
        End If
        
        ' in trade?
        If InTrade > 0 Then
            ' exit out if we're offering that item
            For i = 1 To MAX_INV
                If TradeYourOffer(i).num = InvNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).Type = ITEM_TYPE_CURRENCY Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(i).value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                CurrencyMenu = 4 ' offer in trade
                lblCurrency.Caption = "How many do you want to trade?"
                tmpCurrencyItem = InvNum
                txtCurrency.Text = vbNullString
                picCurrency.Visible = True
                txtCurrency.SetFocus
                Exit Sub
            End If
            
            Call TradeItem(InvNum, 0)
            Exit Sub
        End If
        
        ' use item if not doing anything else
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        Call SendUseItem(InvNum)
        Exit Sub
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsEqItem(ByVal x As Single, ByVal y As Single) As Long
Dim tempRec As RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsEqItem = 0

    For i = 1 To Equipment.Equipment_Count - 1

        If GetPlayerEquipment(MyIndex, i) > 0 And GetPlayerEquipment(MyIndex, i) <= MAX_ITEMS Then

            With tempRec
                .top = EqTop
                .Bottom = .top + PIC_Y
                .Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.top And y <= tempRec.Bottom Then
                    IsEqItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsEqItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsInvItem(ByVal x As Single, ByVal y As Single) As Long
Dim tempRec As RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsInvItem = 0

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then

            With tempRec
                .top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.top And y <= tempRec.Bottom Then
                    IsInvItem = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsInvItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsPlayerSpell(ByVal x As Single, ByVal y As Single) As Long
Dim tempRec As RECT
Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsPlayerSpell = 0

    For i = 1 To MAX_PLAYER_SPELLS

        If PlayerSpells(i) > 0 And PlayerSpells(i) <= MAX_SPELLS Then

            With tempRec
                .top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.top And y <= tempRec.Bottom Then
                    IsPlayerSpell = i
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlayerSpell", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Function IsTradeItem(ByVal x As Single, ByVal y As Single, ByVal Yours As Boolean) As Long
Dim tempRec As RECT
Dim i As Long
Dim itemNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsTradeItem = 0

    For i = 1 To MAX_INV
        If Yours Then
            itemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)
        Else
            itemNum = TradeTheirOffer(i).num
        End If

        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            With tempRec
                .top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.top And y <= tempRec.Bottom Then
                    IsTradeItem = i
                    Exit Function
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsTradeItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub picInventory_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim InvNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InvNum = IsInvItem(x, y)

    If Button = 1 Then
        If InvNum <> 0 Then
            If InTrade > 0 Then Exit Sub
            If InBank Or InShop Then Exit Sub
            DragInvSlotNum = InvNum
        End If
    ElseIf Button = 2 Then
        If Not InBank And Not InShop And Not InTrade > 0 Then
            If InvNum <> 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                    If GetPlayerInvItemValue(MyIndex, InvNum) > 0 Then
                        CurrencyMenu = 1 ' drop
                        lblCurrency.Caption = "How many do you want to drop?"
                        tmpCurrencyItem = InvNum
                        txtCurrency.Text = vbNullString
                        picCurrency.Visible = True
                        txtCurrency.SetFocus
                    End If
                Else
                    Call SendDropItem(InvNum, 0)
                End If
            End If
        End If
    End If

    SetFocusOnChat
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picInventory_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim InvNum As Long
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InvX = x
    InvY = y

    If DragInvSlotNum > 0 Then
        If InTrade > 0 Then Exit Sub
        If InBank Or InShop Then Exit Sub
        Call BltInventoryItem(x + picInventory.Left, y + picInventory.top)
    Else
        InvNum = IsInvItem(x, y)

        If InvNum <> 0 Then
            ' exit out if we're offering that item
            If InTrade Then
                For i = 1 To MAX_INV
                    If TradeYourOffer(i).num = InvNum Then
                        ' is currency?
                        If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).Type = ITEM_TYPE_CURRENCY Then
                            ' only exit out if we're offering all of it
                            If TradeYourOffer(i).value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).num) Then
                                Exit Sub
                            End If
                        Else
                            Exit Sub
                        End If
                    End If
                Next
            End If
            x = x + picInventory.Left - picItemDesc.width - 1
            y = y + picInventory.top - picItemDesc.height - 1
            UpdateDescWindow GetPlayerInvItemNum(MyIndex, InvNum), x, y
            LastItemDesc = GetPlayerInvItemNum(MyIndex, InvNum) ' set it so you don't re-set values
            Exit Sub
        End If
    End If

    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picInventory_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    Dim recPos As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If InTrade > 0 Then Exit Sub
    If InBank Or InShop Then Exit Sub

    If DragInvSlotNum > 0 Then
        ' drag + drop
        For i = 1 To MAX_INV
            With recPos
                .top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= recPos.Left And x <= recPos.Right Then
                If y >= recPos.top And y <= recPos.Bottom Then '
                    If DragInvSlotNum <> i Then
                        SendChangeInvSlots DragInvSlotNum, i
                        Exit For
                    End If
                End If
            End If
        Next
        ' hotbar
        For i = 1 To MAX_HOTBAR
            With recPos
                .top = picHotbar.top - picInventory.top
                .Left = picHotbar.Left - picInventory.Left + (HotbarOffsetX * (i - 1)) + (32 * (i - 1))
                .Right = .Left + 32
                .Bottom = picHotbar.top - picInventory.top + 32
            End With
            
            If x >= recPos.Left And x <= recPos.Right Then
                If y >= recPos.top And y <= recPos.Bottom Then
                    SendHotbarChange 1, DragInvSlotNum, i
                    DragInvSlotNum = 0
                    picTempInv.Visible = False
                    BltHotbar
                    Exit Sub
                End If
            End If
        Next
    End If

    DragInvSlotNum = 0
    picTempInv.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picInventory_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picItemDesc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picItemDesc.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picItemDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ***************
' * Char window *
' ***************
Private Sub picCharacter_Click()
Dim EqNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    EqNum = IsEqItem(EqX, EqY)
    If EqNum <> 0 Then SendUnequip EqNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCharacter_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picCharacter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim EqNum As Long
Dim x2 As Long, y2 As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    EqX = x
    EqY = y
    EqNum = IsEqItem(x, y)

    If EqNum <> 0 Then
        y2 = y + picCharacter.top - frmMain.picItemDesc.height - 1
        x2 = x + picCharacter.Left - frmMain.picItemDesc.width - 1
        UpdateDescWindow GetPlayerEquipment(MyIndex, EqNum), x2, y2
        LastItemDesc = GetPlayerEquipment(MyIndex, EqNum) ' set it so you don't re-set values
        Exit Sub
    End If

    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picCharacter_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **************
' * Admin Menu *
' **************
Private Sub cmdALoc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    BLoc = Not BLoc
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdALoc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp2Me_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    If Len(Trim$(txtAName.Text)) < 1 Then Exit Sub
    If IsNumeric(Trim$(txtAName.Text)) Then Exit Sub

    WarpToMe Trim$(txtAName.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp2Me_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarpMe2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    If Len(Trim$(txtAName.Text)) < 1 Then Exit Sub
    If IsNumeric(Trim$(txtAName.Text)) Then Exit Sub

    WarpMeTo Trim$(txtAName.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarpMe2_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp_Click()
Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    If Len(Trim$(txtAMap.Text)) < 1 Then Exit Sub
    If Not IsNumeric(Trim$(txtAMap.Text)) Then Exit Sub

    n = CLng(Trim$(txtAMap.Text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        Call WarpTo(n)
    Else
        Call AddText("Invalid map number.", Red)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAWarp_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASprite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    If Len(Trim$(txtASprite.Text)) < 1 Then Exit Sub
    If Not IsNumeric(Trim$(txtASprite.Text)) Then Exit Sub

    SendSetSprite CLng(Trim$(txtASprite.Text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASprite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAMapReport_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendMapReport
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAMapReport_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdARespawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendMapRespawn
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdARespawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdABan_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    If Len(Trim$(txtAName.Text)) < 1 Then Exit Sub
    
    SendBan Trim$(txtAName.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdABan_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAAccess_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then Exit Sub
    If Len(Trim$(txtAName.Text)) < 2 Then Exit Sub
    If IsNumeric(Trim$(txtAName.Text)) Or Not IsNumeric(Trim$(txtAAccess.Text)) Then Exit Sub

    SendSetAccess Trim$(txtAName.Text), CLng(Trim$(txtAAccess.Text))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAAccess_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdADestroy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then Exit Sub
    SendBanDestroy
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdADestroy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then Exit Sub
    SendSpawnItem scrlAItem.value, scrlAAmount.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdASpawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_DblClick()
Dim bankNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DragBankSlotNum = 0

    bankNum = IsBankItem(BankX, BankY)
    If bankNum <> 0 Then
         If GetBankItemNum(bankNum) = ITEM_TYPE_NONE Then Exit Sub
         
             If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_CURRENCY Then
                CurrencyMenu = 3 ' withdraw
                lblCurrency.Caption = "How many do you want to withdraw?"
                tmpCurrencyItem = bankNum
                'txtCurrency.Text = vbNullString
                picCurrency.Visible = True
                txtCurrency.SetFocus
                Exit Sub
            End If
            
         WithdrawItem bankNum, 0
         Exit Sub
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim bankNum As Long
                        
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    bankNum = IsBankItem(x, y)
    
    If bankNum <> 0 Then
        If Button = 1 Then DragBankSlotNum = bankNum
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
Dim recPos As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' TODO : Add sub to change bankslots client side first so there's no delay in switching
    If DragBankSlotNum > 0 Then
        For i = 1 To MAX_BANK
            With recPos
                .top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= recPos.Left And x <= recPos.Right Then
                If y >= recPos.top And y <= recPos.Bottom Then
                    If DragBankSlotNum <> i Then
                        ChangeBankSlots DragBankSlotNum, i
                        Exit For
                    End If
                End If
            End If
        Next
    End If

    DragBankSlotNum = 0
    picTempBank.Visible = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBank_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim bankNum As Long, itemNum As Long, itemType As Long
Dim x2 As Long, y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    BankX = x
    BankY = y
    
    If DragBankSlotNum > 0 Then
        Call BltBankItem(x + picBank.Left, y + picBank.top)
    Else
        bankNum = IsBankItem(x, y)
        
        If bankNum <> 0 Then
            x2 = x + picBank.Left + 1
            y2 = y + picBank.top + 1
            UpdateDescWindow Bank.Item(bankNum).num, x2, y2
            Exit Sub
        End If
    End If
    
    frmMain.picItemDesc.Visible = False
    LastBankDesc = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "picBank_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsBankItem(ByVal x As Single, ByVal y As Single) As Long
Dim tempRec As RECT
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsBankItem = 0
    
    For i = 1 To MAX_BANK
        If GetBankItemNum(i) > 0 And GetBankItemNum(i) <= MAX_ITEMS Then
        
            With tempRec
                .top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With
            
            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.top And y <= tempRec.Bottom Then
                    
                    IsBankItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsBankItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Private Sub txtMyChat_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Focus on the chat bar
    frmMain.txtMyChat.Text = vbNullString
    txtMyChat.Visible = True

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMyChat_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAKick_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    If Len(Trim$(txtAName.Text)) < 1 Then Exit Sub
    
    ' kick 'em!
    SendKick Trim$(txtAName.Text)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdAKick_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub