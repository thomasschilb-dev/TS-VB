Attribute VB_Name = "modDirectX"
Option Explicit

Private m_clsDirectX            As DirectX8
Private m_intDirectXRefCount    As Integer

Public Property Get DirectX() As DirectX8
    Set DirectX = m_clsDirectX
End Property

Public Function InitializeDirectX() As Boolean
    If m_clsDirectX Is Nothing Then
        On Error GoTo ErrorHandler
            Set m_clsDirectX = New DirectX8
        On Error GoTo 0
    End If
    
    m_intDirectXRefCount = m_intDirectXRefCount + 1
    InitializeDirectX = True

ErrorHandler:
End Function

Public Function DeinitializeDirectX() As Boolean
    If m_clsDirectX Is Nothing Then
        DeinitializeDirectX = False
    Else
        m_intDirectXRefCount = m_intDirectXRefCount - 1
        
        If m_intDirectXRefCount = 0 Then
            Set m_clsDirectX = Nothing
        End If
        
        DeinitializeDirectX = True
    End If
End Function
