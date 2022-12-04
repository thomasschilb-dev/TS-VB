VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataQueueItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CpyMem Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDst As Any, pSrc As Any, ByVal cBytes As Long _
)

Private m_udtMemory         As Memory
Private m_lngDataPosition   As Long

Public Sub Initialize(ByVal pData As Long, ByVal DataLen As Long)
    m_udtMemory = AllocMemory(DataLen)
    CpyMem ByVal m_udtMemory.address, ByVal pData, DataLen
End Sub

Public Sub Free()
    If m_udtMemory.address <> 0 Then
        FreeMemory m_udtMemory
    End If
End Sub

Public Property Get DataPointer() As Long
    DataPointer = m_udtMemory.address
End Property

Public Property Get DataSize() As Long
    DataSize = m_udtMemory.bytes
End Property

Public Property Get DataPosition() As Long
    DataPosition = m_lngDataPosition
End Property

Public Property Let DataPosition(ByVal value As Long)
    m_lngDataPosition = value
End Property

Public Property Get DataLeft() As Long
    DataLeft = DataSize - DataPosition
End Property

Public Property Get EndOfBuffer() As Boolean
    If m_udtMemory.bytes <= 0 Then
        EndOfBuffer = True
    Else
        EndOfBuffer = DataLeft = 0
    End If
End Property

Private Sub Class_Terminate()
    Free
End Sub
