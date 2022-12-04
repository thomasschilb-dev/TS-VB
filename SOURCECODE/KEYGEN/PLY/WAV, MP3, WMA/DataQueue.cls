VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_clsItems()            As DataQueueItem
Private m_lngItemCount          As Long

Public Sub Clear()
    Dim i   As Long
    
    If Count > 0 Then
        For i = 0 To m_lngItemCount
            Set m_clsItems(i) = Nothing
        Next
    End If
    
    m_lngItemCount = 0
End Sub

Public Property Get Count() As Long
    Count = m_lngItemCount
End Property

Public Function Dequeue() As DataQueueItem
    Dim i   As Long

    If m_lngItemCount > 0 Then
        Set Dequeue = m_clsItems(0)
    
        For i = 0 To m_lngItemCount - 2
            Set m_clsItems(i) = m_clsItems(i + 1)
        Next
    
        Set m_clsItems(m_lngItemCount - 1) = Nothing
    
        m_lngItemCount = m_lngItemCount - 1
    Else
        Dequeue = Nothing
    End If
End Function

Public Sub Enqueue(item As DataQueueItem)
    m_lngItemCount = m_lngItemCount + 1
    
    If UBound(m_clsItems) + 1 < m_lngItemCount Then
        ReDim Preserve m_clsItems(UBound(m_clsItems) + 100)
    End If
    
    Set m_clsItems(m_lngItemCount - 1) = item
End Sub

Private Sub Class_Initialize()
    ReDim m_clsItems(100) As DataQueueItem
End Sub

Private Sub Class_Terminate()
    Dim i   As Long
    
    For i = 0 To m_lngItemCount - 1
        Set m_clsItems(i) = Nothing
    Next
End Sub
