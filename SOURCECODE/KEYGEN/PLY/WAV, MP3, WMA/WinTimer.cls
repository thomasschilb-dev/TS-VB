VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WinTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetTimer Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long _
) As Long

Private Declare Function KillTimer Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIDEvent As Long _
) As Long

Public Event Tick()

Private Const WM_TIMER  As Long = &H113&

Private m_lngInterval   As Long
Private m_hTimer        As Long
Private m_blnEnabled    As Boolean

Private m_udtCallback   As MethodDelegate

' needs to be the first public method in this interface!
Public Sub TimerProc( _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
)

    If uMsg = WM_TIMER Then
        RaiseEvent Tick
    End If
End Sub

Public Sub RaiseTimer()
    TimerProc 0, WM_TIMER, 0, 0
End Sub

Public Property Get Interval() As Long
    Interval = m_lngInterval
End Property

Public Property Let Interval(ByVal value As Long)
    m_lngInterval = value
    
    If Enabled Then
        Enabled = False
        Enabled = True
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_blnEnabled
End Property

Public Property Let Enabled(ByVal value As Boolean)
    If Not value = m_blnEnabled Then
        m_blnEnabled = value
    
        If m_blnEnabled Then
            m_hTimer = SetTimer(0, 0, m_lngInterval, m_udtCallback.addr)
        Else
            KillTimer 0, m_hTimer
            m_hTimer = 0
        End If
    End If
End Property

Private Sub Class_Initialize()
    m_udtCallback = CreateDelegate(Me, 0, 4)
End Sub

Private Sub Class_Terminate()
    Enabled = False
    FreeDelegate m_udtCallback
End Sub
