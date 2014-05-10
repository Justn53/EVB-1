VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextureManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Textures As Collection
Private D3DX As D3DX8

Private Sub Class_Initialize(ByVal dx As D3DX8)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    D3DX = dx
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Class_Initialize", "clsTextureManager", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadTexture(ByVal filePath As String, ByVal textureName As String)
Dim texture As Direct3DTexture8

    texture = D3DX.CreateTextureFromFile(D3DDevice, filePath)

    Call Textures.Add(texture, textureName)
    
End Sub

