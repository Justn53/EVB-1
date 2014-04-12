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

Private Buffer() As Byte
Private BufferSize As Long
Private WriteHead As Long
Private ReadHead As Long

Private Sub Class_Initialize()
    Flush
End Sub

Public Sub PreAllocate(ByVal nLength As Long)
    WriteHead = 0
    ReadHead = 0
    BufferSize = nLength - 1
    ReDim Buffer(0 To BufferSize)
End Sub

Public Sub Allocate(ByVal nLength As Long)
    If BufferSize = 0 And nLength > 1 Then nLength = nLength - 1
    BufferSize = BufferSize + nLength
    ReDim Preserve Buffer(0 To BufferSize)
End Sub

Public Sub Flush()
    WriteHead = 0
    ReadHead = 0
    BufferSize = 0
    ReDim Buffer(0)
End Sub

Public Sub Trim()
    ' If the readhead is past the buffersize, this means everything has been read in the array, flush it
    If ReadHead >= Count Then Flush
End Sub

Public Sub WriteByte(ByVal nByte As Byte)

    If WriteHead > BufferSize Then Allocate 1
 
    Buffer(WriteHead) = nByte
    WriteHead = WriteHead + 1
End Sub

Public Sub WriteBytes(ByRef nByte() As Byte)
Dim nLength As Long
         
    nLength = (UBound(nByte) - LBound(nByte)) + 1
    
    If WriteHead + nLength - 1 > BufferSize Then Allocate nLength
    
    CopyMemory Buffer(WriteHead), nByte(0), nLength
    WriteHead = WriteHead + nLength
End Sub

Public Sub WriteCurrency(ByVal nCurrency As Currency)

    If WriteHead + 7 > BufferSize Then Allocate 8
    
    CopyMemory Buffer(WriteHead), nCurrency, 8
    WriteHead = WriteHead + 8
End Sub

Public Sub WriteDouble(ByVal nDouble As Double)

    If WriteHead + 7 > BufferSize Then Allocate 8
    
    CopyMemory Buffer(WriteHead), nDouble, 8
    WriteHead = WriteHead + 8
End Sub

Public Sub WriteSingle(ByVal nSingle As Single)

    If WriteHead + 3 > BufferSize Then Allocate 4
    
    CopyMemory Buffer(WriteHead), nSingle, 4
    WriteHead = WriteHead + 4
End Sub

Public Sub WriteLong(ByVal nLong As Long)

    If WriteHead + 3 > BufferSize Then Allocate 4
    
    CopyMemory Buffer(WriteHead), nLong, 4
    WriteHead = WriteHead + 4
End Sub

Public Sub WriteString(ByRef nString As String)
Dim sBytes() As Byte
Dim sLength As Long
    
    sLength = Len(nString)
    sBytes = StrConv(nString, vbFromUnicode)
    
    WriteLong sLength
    
    If sLength <= 0 Then Exit Sub
    
    If WriteHead + sLength - 1 > BufferSize Then Allocate sLength
      
    CopyMemory Buffer(WriteHead), sBytes(0), sLength
    WriteHead = WriteHead + sLength
End Sub

Public Function ReadByte(Optional MoveReadHead As Boolean = True) As Byte

    If ReadHead > BufferSize Then Exit Function
    
    ReadByte = Buffer(ReadHead)
    If MoveReadHead Then ReadHead = ReadHead + 1
End Function

Public Function ReadBytes(ByVal nLength As Long, Optional MoveReadHead As Boolean = True) As Byte()
Dim Data() As Byte

    If nLength = 0 Then Exit Function
    If ReadHead + nLength - 1 > BufferSize Then Exit Function
        
    ReDim Data(nLength - 1)
        
    CopyMemory Data(0), Buffer(ReadHead), nLength
    If MoveReadHead Then ReadHead = ReadHead + nLength
    
    ReadBytes = Data
End Function

Public Function ReadCurrency(Optional MoveReadHead As Boolean = True) As Currency

    If ReadHead + 7 > BufferSize Then Exit Function
    
    CopyMemory ReadCurrency, Buffer(ReadHead), 8
    If MoveReadHead Then ReadHead = ReadHead + 8

End Function

Public Function ReadSingle(Optional MoveReadHead As Boolean = True) As Single

    If ReadHead + 3 > BufferSize Then Exit Function
    
    CopyMemory ReadSingle, Buffer(ReadHead), 4
    If MoveReadHead Then ReadHead = ReadHead + 4

End Function

Public Function ReadDouble(Optional MoveReadHead As Boolean = True) As Double

    If ReadHead + 7 > BufferSize Then Exit Function
    
    CopyMemory ReadDouble, Buffer(ReadHead), 8
    If MoveReadHead Then ReadHead = ReadHead + 8

End Function


Public Function ReadLong(Optional MoveReadHead As Boolean = True) As Long

    If ReadHead + 3 > BufferSize Then Exit Function
    
    CopyMemory ReadLong, Buffer(ReadHead), 4
    If MoveReadHead Then ReadHead = ReadHead + 4
End Function

Public Function ReadString(Optional MoveReadHead As Boolean = True) As String
Dim sLength As Long
Dim sBytes() As Byte
  
    sLength = ReadLong(False)
    If sLength <= 0 Then
        If MoveReadHead Then ReadHead = ReadHead + 4
        Exit Function
    End If
    
    ReDim sBytes(sLength - 1)
    
    CopyMemory sBytes(0), Buffer(ReadHead + 4), sLength
    
    ReadString = StrConv(sBytes, vbUnicode)
    If MoveReadHead Then ReadHead = ReadHead + sLength + 4
End Function

Public Function Count() As Long
    Count = (UBound(Buffer) - LBound(Buffer)) + 1
End Function

Public Function Length() As Long
    Length = Count - ReadHead
End Function

Public Function ToArray() As Byte()
    ToArray = Buffer()
End Function

Public Function ToString() As String
    ToString = StrConv(Buffer, vbUnicode)
End Function
