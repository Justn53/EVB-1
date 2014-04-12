VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private buffer() As Byte
Private BufferSize As Long
Private WriteHead As Long
Private ReadHead As Long

Private Sub Class_Initialize()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Flush
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Class_Initialize", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PreAllocate(ByVal nLength As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    WriteHead = 0
    ReadHead = 0
    BufferSize = nLength - 1
    ReDim buffer(0 To BufferSize)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PreAllocate", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Allocate(ByVal nLength As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If BufferSize = 0 And nLength > 1 Then nLength = nLength - 1
    BufferSize = BufferSize + nLength
    ReDim Preserve buffer(0 To BufferSize)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Allocate", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Flush()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    WriteHead = 0
    ReadHead = 0
    BufferSize = 0
    ReDim buffer(0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Flush", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Trim()
Dim c As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' If the readhead is past the buffersize, this means everything has been read in the array, flush it
    If ReadHead >= Count Then
        Flush
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Trim", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WriteByte(ByVal nByte As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If WriteHead > BufferSize Then Allocate 1
 
    buffer(WriteHead) = nByte
    WriteHead = WriteHead + 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WriteByte", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WriteBytes(ByRef nByte() As Byte)
Dim nLength As Long
        
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    nLength = (UBound(nByte) - LBound(nByte)) + 1
    
    If WriteHead + nLength - 1 > BufferSize Then Allocate nLength
    
    CopyMemory buffer(WriteHead), nByte(0), nLength
    WriteHead = WriteHead + nLength
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WriteBytes", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WriteLong(ByVal nLong As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If WriteHead + 3 > BufferSize Then Allocate 4
    
    CopyMemory buffer(WriteHead), nLong, 4
    WriteHead = WriteHead + 4
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WriteLong", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WriteString(ByRef nString As String)
Dim sBytes() As Byte
Dim sLength As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    sLength = Len(nString)
    sBytes = StrConv(nString, vbFromUnicode)
    
    WriteLong sLength
    
    If sLength <= 0 Then Exit Sub
    
    If WriteHead + sLength - 1 > BufferSize Then Allocate sLength
      
    CopyMemory buffer(WriteHead), sBytes(0), sLength
    WriteHead = WriteHead + sLength
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "WriteString", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ReadByte(Optional MoveReadHead As Boolean = True) As Byte
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If ReadHead > BufferSize Then Exit Function
    
    ReadByte = buffer(ReadHead)
    If MoveReadHead Then ReadHead = ReadHead + 1
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ReadByte", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function


Public Function ReadCurrency(Optional MoveReadHead As Boolean = True) As Currency

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ReadHead + 7 > BufferSize Then Exit Function
    
    CopyMemory ReadCurrency, buffer(ReadHead), 8
    If MoveReadHead Then ReadHead = ReadHead + 8

' Error handler
    Exit Function
errorhandler:
    HandleError "ReadByte", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ReadSingle(Optional MoveReadHead As Boolean = True) As Single

 ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler


    If ReadHead + 3 > BufferSize Then Exit Function
    
    CopyMemory ReadSingle, buffer(ReadHead), 4
    If MoveReadHead Then ReadHead = ReadHead + 4

' Error handler
    Exit Function
errorhandler:
    HandleError "ReadByte", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ReadDouble(Optional MoveReadHead As Boolean = True) As Double

 ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler


    If ReadHead + 7 > BufferSize Then Exit Function
    
    CopyMemory ReadDouble, buffer(ReadHead), 8
    If MoveReadHead Then ReadHead = ReadHead + 8

' Error handler
    Exit Function
errorhandler:
    HandleError "ReadByte", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ReadBytes(ByVal nLength As Long, Optional MoveReadHead As Boolean = True) As Byte()
Dim data() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If nLength = 0 Then Exit Function
    If ReadHead + nLength - 1 > BufferSize Then Exit Function
        
    ReDim data(nLength - 1)
        
    CopyMemory data(0), buffer(ReadHead), nLength
    If MoveReadHead Then ReadHead = ReadHead + nLength
    
    ReadBytes = data
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ReadBytes", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ReadLong(Optional MoveReadHead As Boolean = True) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If ReadHead + 3 > BufferSize Then Exit Function
    
    CopyMemory ReadLong, buffer(ReadHead), 4
    If MoveReadHead Then ReadHead = ReadHead + 4
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ReadLong", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ReadString(Optional MoveReadHead As Boolean = True) As String
Dim sLength As Long
Dim sBytes() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    sLength = ReadLong(False)
    If sLength <= 0 Then
        If MoveReadHead Then ReadHead = ReadHead + 4
        Exit Function
    End If
    
    ReDim sBytes(sLength - 1)
    
    CopyMemory sBytes(0), buffer(ReadHead + 4), sLength
    
    ReadString = StrConv(sBytes, vbUnicode)
    If MoveReadHead Then ReadHead = ReadHead + sLength + 4
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ReadString", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function Count() As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Count = (UBound(buffer) - LBound(buffer)) + 1
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "Count", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function length() As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    length = Count - ReadHead
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "Length", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ToArray() As Byte()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ToArray = buffer()
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ToArray", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ToString() As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ToString = StrConv(buffer, vbUnicode)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ToString", "clsBuffer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
