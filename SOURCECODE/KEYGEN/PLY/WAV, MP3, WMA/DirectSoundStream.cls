VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectSoundStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CpyMem Lib "kernel32" Alias "RtlMoveMemory" ( _
    pDst As Any, pSrc As Any, ByVal cBytes As Long _
)

Private Declare Sub FillMem Lib "kernel32" Alias "RtlFillMemory" ( _
    pDst As Any, ByVal cBytes As Long, ByVal value As Byte _
)

Public Enum PlaybackStatus
    PlaybackStopped
    PlaybackPlaying
    PlaybackPausing
End Enum

Private Type AudioBuffer
    pData       As Long
    DataLen     As Long
End Type

Private Const TimerIntervalMs   As Long = 5

Private WithEvents m_clsTimer   As WinTimer
Attribute m_clsTimer.VB_VarHelpID = -1

Private m_clsBufSec             As DirectSoundSecondaryBuffer8
Private m_udeStatus             As PlaybackStatus
Private m_udtFormat             As WAVEFORMATEX
Private m_lngDSBufBytes         As Long
Private m_lngDSBufMs            As Long
Private m_lngWriteCursor        As Long

Private m_blnEndOfStream        As Boolean
Private m_lngEOSPosition        As Long
Private m_lngEOSDataLeft        As Long
Private m_lngEOSDataLeftMarker  As Long

Private m_lngMSPlayed           As Long
Private m_lngMSPlayedMarker     As Long

Private m_lngAudioQueueLengthMs As Long

Private m_clsBuffers            As DataQueue
Private m_clsCurBuf             As DataQueueItem

Public Event StatusChanged(ByVal status As PlaybackStatus)
Public Event BufferDone()
Public Event NoDataLeft()


''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Various

Public Function BytesFromMs(ByVal ms As Long) As Long
    BytesFromMs = MsToBytes(m_udtFormat, ms)
End Function

' data left to play in milliseconds
Public Property Get AudioBufferQueueLength() As Long
    Dim udtPos  As DSCURSORS
    Dim lngLen  As Long
    
    On Error Resume Next
        m_clsBufSec.GetCurrentPosition udtPos
    On Error GoTo 0
    
    lngLen = RingDistance(udtPos.lPlay, m_lngWriteCursor)
    lngLen = lngLen + m_lngAudioQueueLengthMs
    lngLen = lngLen + BytesToMs(m_udtFormat, m_clsCurBuf.DataLeft)
    
    AudioBufferQueueLength = lngLen
End Property

Public Property Get Elapsed() As Long
    Elapsed = m_lngMSPlayed
End Property

Public Property Let Elapsed(ByVal value As Long)
    m_lngMSPlayed = value
End Property

Public Property Get Volume() As Long
    Volume = m_clsBufSec.GetVolume
End Property

Public Property Let Volume(ByVal value As Long)
    If value < -10000 Then value = -10000
    If value > 0 Then value = 0

    m_clsBufSec.SetVolume value
End Property

Public Property Get Balance() As Long
    Balance = m_clsBufSec.GetPan
End Property

Public Property Let Balance(ByVal value As Long)
    If value < -10000 Then value = -10000
    If value > 10000 Then value = 10000
    
    m_clsBufSec.SetPan value
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Playback Handling

Public Function PlaybackStop() As Boolean
    If PlaybackStatus = PlaybackStopped Then
        PlaybackStop = True
    Else
        m_clsTimer.Enabled = False
        m_clsBufSec.Stop
        m_clsBufSec.SetCurrentPosition 0
        m_lngWriteCursor = 0
        m_lngMSPlayed = 0
        m_lngMSPlayedMarker = 0
        m_blnEndOfStream = False
        
        m_clsCurBuf.Free
        Set m_clsCurBuf = Nothing
        m_clsBuffers.Clear
        
        PlaybackStop = True
        SetPlaybackStatus PlaybackStopped
    End If
End Function

Public Function PlaybackPause() As Boolean
    If PlaybackStatus = PlaybackPausing Then
        PlaybackPause = True
    Else
        m_clsTimer.Enabled = False
        m_clsBufSec.Stop
        PlaybackPause = True
        SetPlaybackStatus PlaybackPausing
    End If
End Function

Public Function PlaybackStart() As Boolean
    If PlaybackStatus = PlaybackPlaying Then
        PlaybackStart = True
    Else
        If m_clsBuffers.Count > 0 Then
            If PlaybackStatus = PlaybackStopped Then
                m_clsBufSec.SetCurrentPosition 0
                m_lngWriteCursor = 0
                m_lngMSPlayedMarker = 0
                FillDSBuffer m_lngDSBufBytes
            End If

            m_clsBufSec.Play DSBPLAY_LOOPING
            m_clsTimer.Enabled = True
        ElseIf PlaybackStatus = PlaybackPausing Then
            m_clsBufSec.Play DSBPLAY_LOOPING
            m_clsTimer.Enabled = True
        End If
        
        SetPlaybackStatus PlaybackPlaying
        PlaybackStart = True
    End If
End Function

Public Property Get PlaybackStatus() As PlaybackStatus
    PlaybackStatus = m_udeStatus
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Audio Data Buffering

Private Sub m_clsTimer_Tick()
    Dim udtCursors      As DSCURSORS
    Dim lngDistance     As Long
    Dim lngRead         As Long
    
    On Error GoTo ErrorHandler
        m_clsBufSec.GetCurrentPosition udtCursors
    On Error GoTo 0

    On Error Resume Next
        Err.Clear
        ' 25 days should fit in m_lngMSPlayed
        m_lngMSPlayed = m_lngMSPlayed + BytesToMs(m_udtFormat, RingDistance(m_lngMSPlayedMarker, udtCursors.lPlay))
        If Err.Number = 6 Then m_lngMSPlayed = 0    ' Overflow
    On Error GoTo 0
    
    m_lngMSPlayedMarker = udtCursors.lPlay

    If m_blnEndOfStream Then
        m_lngEOSDataLeft = m_lngEOSDataLeft - RingDistance(m_lngEOSDataLeftMarker, udtCursors.lPlay)
        m_lngEOSDataLeftMarker = udtCursors.lPlay
        
        If m_lngEOSDataLeft <= 0 Then
            m_clsTimer.Enabled = False
            m_clsBufSec.Stop
            RaiseEvent NoDataLeft
        End If
    Else
        lngDistance = RingDistance(m_lngWriteCursor, udtCursors.lPlay)
    
        If lngDistance >= MsToBytes(m_udtFormat, 100) Then
            lngRead = FillDSBuffer(lngDistance)
            If lngRead < lngDistance Then
                If m_clsBuffers.Count = 0 Then
                    m_blnEndOfStream = True
                    m_lngEOSPosition = m_lngWriteCursor
                    m_lngEOSDataLeft = RingDistance(udtCursors.lPlay, m_lngEOSPosition)
                    m_lngEOSDataLeftMarker = udtCursors.lPlay
                End If
            End If
        End If
    End If

ErrorHandler:
End Sub

Public Property Get NoDataLeft() As Boolean
    If m_blnEndOfStream Then
        If m_lngEOSDataLeft <= 0 Then
            NoDataLeft = True
        End If
    End If
End Property

Public Function AudioBufferAdd(ByVal DataPointer As Long, ByVal DataSize As Long) As Boolean
    Dim clsBufferItem   As New DataQueueItem

    clsBufferItem.Initialize DataPointer, DataSize
    m_clsBuffers.Enqueue clsBufferItem

    m_lngAudioQueueLengthMs = m_lngAudioQueueLengthMs + BytesToMs(m_udtFormat, DataSize)

    ' if playback is running but there are no buffers,
    ' the DirectSound buffer will be stopped.
    ' When there is data, playback will be resumed
    If m_blnEndOfStream Then
        If m_clsBuffers.Count = 1 Then
            m_clsBufSec.SetCurrentPosition 0
            m_lngWriteCursor = 0
            
            ' idea: > PreBuffer MS?
            If FillDSBuffer(m_lngDSBufBytes) > 0 Then
                m_blnEndOfStream = False
                
                On Error GoTo ErrorHandler
                    m_clsBufSec.Play DSBPLAY_LOOPING
                On Error GoTo 0
                
                m_clsTimer.Enabled = True
            End If
            
        End If
    End If
    
    AudioBufferAdd = True
    Exit Function
    
ErrorHandler:
    AudioBufferAdd = False
End Function

Public Sub AudioBuffersClear()
    m_clsCurBuf.Free
    m_clsCurBuf = Nothing
    
    m_clsBuffers.Clear
End Sub

Public Property Get AudioBufferCount() As Long
    AudioBufferCount = m_clsBuffers.Count
End Property

Public Property Get BufferLengthMs() As Long
    BufferLengthMs = m_lngDSBufMs
End Property

' returns bytes filled
Private Function FillDSBuffer(ByVal bytes As Long) As Long
    Dim btData()    As Byte
    Dim lngRead     As Long

    lngRead = GetData(bytes, btData)

    m_clsBufSec.WriteBuffer m_lngWriteCursor, bytes, btData(0), DSBLOCK_DEFAULT

    m_lngWriteCursor = (m_lngWriteCursor + lngRead) Mod m_lngDSBufBytes
    
    FillDSBuffer = lngRead
End Function

' returns bytes read
Private Function GetData(ByVal bytes As Long, ByRef btData() As Byte) As Long
    ReDim btData(bytes - 1) As Byte
    Dim lngRead             As Long
    
    If FormatBitsPerSample = 8 Then
        ' fill complete buffer with silence
        FillMem btData(0), bytes, &H80
    End If
    
    ' check if there are audio buffers to read data from
    If m_clsCurBuf Is Nothing Then
        If m_clsBuffers.Count > 0 Then
            Set m_clsCurBuf = m_clsBuffers.Dequeue()
            If m_clsCurBuf Is Nothing Then
                GetData = 0
                Exit Function
            Else
                m_lngAudioQueueLengthMs = m_lngAudioQueueLengthMs - BytesToMs(m_udtFormat, m_clsCurBuf.DataSize)
            End If
        Else
            m_lngAudioQueueLengthMs = 0
            GetData = 0
            Exit Function
        End If
    End If

    Do While lngRead < bytes
        If m_clsCurBuf.EndOfBuffer Then
            m_clsCurBuf.Free
            
            RaiseEvent BufferDone
            
            ' current audio buffer was completly read, get the next one
            If m_clsBuffers.Count > 0 Then
                Set m_clsCurBuf = m_clsBuffers.Dequeue
                If m_clsCurBuf Is Nothing Then
                    Exit Do
                Else
                    m_lngAudioQueueLengthMs = m_lngAudioQueueLengthMs - BytesToMs(m_udtFormat, m_clsCurBuf.DataSize)
                End If
            Else
                m_lngAudioQueueLengthMs = 0
                Exit Do
            End If
        End If
        
        If lngRead + m_clsCurBuf.DataLeft > bytes Then
            ' more data in the current buffer then we actually need
            CpyMem btData(lngRead), ByVal m_clsCurBuf.DataPointer + m_clsCurBuf.DataPosition, bytes - lngRead
            m_clsCurBuf.DataPosition = m_clsCurBuf.DataPosition + (bytes - lngRead)
            lngRead = bytes
        Else
            ' whole audio buffer will fit in btData
            CpyMem btData(lngRead), ByVal m_clsCurBuf.DataPointer + m_clsCurBuf.DataPosition, m_clsCurBuf.DataLeft
            lngRead = lngRead + m_clsCurBuf.DataLeft
            m_clsCurBuf.DataPosition = m_clsCurBuf.DataSize
        End If
    Loop
    
    GetData = lngRead
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Audio Format of Secondary Buffer

Public Property Get FormatSamplerate() As Long
    FormatSamplerate = m_udtFormat.lSamplesPerSec
End Property

Public Property Get FormatChannels() As Integer
    FormatChannels = m_udtFormat.nChannels
End Property

Public Property Get FormatBitsPerSample() As Integer
    FormatBitsPerSample = m_udtFormat.nBitsPerSample
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Helpers

Private Function RingDistance(ByVal FromByte As Long, ByVal ToByte As Long) As Long
    If ToByte < FromByte Then
        RingDistance = ToByte + m_lngDSBufBytes - FromByte
    Else
        RingDistance = ToByte - FromByte
    End If
End Function

Private Sub SetPlaybackStatus(ByVal stat As PlaybackStatus, Optional ByVal silent As Boolean = False)
    m_udeStatus = stat
    If Not silent Then RaiseEvent StatusChanged(m_udeStatus)
End Sub

Private Function MsToBytes(fmt As WAVEFORMATEX, ByVal ms As Long) As Long
    MsToBytes = ms / 1000 * fmt.lAvgBytesPerSec
End Function

Private Function BytesToMs(fmt As WAVEFORMATEX, ByVal bytes As Long) As Long
    BytesToMs = bytes / fmt.lAvgBytesPerSec * 1000
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Constructor(s)

Private Sub Class_Initialize()
    Set m_clsBuffers = New DataQueue
    Set m_clsTimer = New WinTimer

    m_clsTimer.Interval = TimerIntervalMs
    m_udeStatus = PlaybackStopped
    m_lngWriteCursor = 0
End Sub

Private Sub Class_Terminate()
    PlaybackStop
    
    Set m_clsCurBuf = Nothing
    Set m_clsBuffers = Nothing
    Set m_clsBufSec = Nothing
    Set m_clsTimer = Nothing
End Sub

' Called by the class "DirectSound"
' Has to be the last member of this interface!
Private Sub CustomConstructor(ByVal BufferSize As Long, dsound As DirectSoundSecondaryBuffer8)
    Set m_clsBufSec = dsound
    
    If Not m_clsBufSec Is Nothing Then
        m_clsBufSec.GetFormat m_udtFormat
        
        m_lngDSBufBytes = BufferSize
        m_lngDSBufMs = BytesToMs(m_udtFormat, BufferSize)
    End If
End Sub
