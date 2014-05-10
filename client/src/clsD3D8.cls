VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsD3D8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private D3DX As D3DX8

Public TextureManager As clsTextureManager

Private Sub Class_Initialize()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set D3DX = New D3DX8
    Set TextureManager = New clsTextureManager
    
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Class_Initialize", "clsD3D8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

