VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateWindowEx Lib "user32" _
Alias "CreateWindowExA" ( _
    ByVal dwExStyle As Long, _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String, _
    ByVal dwStyle As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hWndParent As Long, _
    ByVal hMenu As Long, _
    ByVal hInstance As Long, _
    ByVal lpParam As Long _
) As Long

Private Declare Function DestroyWindow Lib "user32" ( _
    ByVal hWnd As Long _
) As Long

Private m_clsPrimBuf    As DirectSoundPrimaryBuffer8
Private m_clsDSound     As DirectSound8
Private m_blnReady      As Boolean
Private m_hDSWnd        As Long

Public Property Get DirectXAvailable() As Boolean
    DirectXAvailable = m_blnReady
End Property

Public Property Get DeviceCount() As Long
    DeviceCount = DirectX.GetDSEnum.GetCount
End Property

Public Property Get DeviceName(ByVal index As Long) As String
    DeviceName = DirectX.GetDSEnum.GetName(index)
End Property

Public Property Get DeviceDescription(ByVal index As Long) As String
    DeviceDescription = DirectX.GetDSEnum.GetDescription(index)
End Property

Public Property Get DeviceGuid(ByVal index As Long) As String
    DeviceGuid = DirectX.GetDSEnum.GetGuid(index)
End Property

Public Function CreateStream( _
    ByVal Samplerate As Long, _
    ByVal Channels As Integer, _
    ByVal BitsPerSample As Integer, _
    ByVal BufferLengthMs As Long, _
    ByRef stream As DirectSoundStream _
) As Boolean

    Dim clsStream As DirectSoundStream
    Dim clsSecBuf As DirectSoundSecondaryBuffer8
    Dim udtBufDsc As DSBUFFERDESC
    
    If DirectXAvailable Then
        If BufferLengthMs < 200 Then BufferLengthMs = 200
        If BufferLengthMs < 1000 Then
            If (BufferLengthMs \ 100) Mod 2 = 1 Then
                BufferLengthMs = (BufferLengthMs \ 100 + 1) * 100
            End If
        End If
    
        With udtBufDsc
            With .fxFormat
                .lSamplesPerSec = Samplerate
                .nChannels = Channels
                .nBitsPerSample = BitsPerSample
                .nBlockAlign = .nChannels * (.nBitsPerSample \ 8)
                .lAvgBytesPerSec = .nBlockAlign * .lSamplesPerSec
                .nFormatTag = WAVE_FORMAT_PCM
            End With
            
            .lFlags = DSBCAPS_CTRLPAN Or _
                      DSBCAPS_CTRLVOLUME Or _
                      DSBCAPS_GETCURRENTPOSITION2 Or _
                      DSBCAPS_GLOBALFOCUS Or _
                      DSBCAPS_STATIC
                      
            .lBufferBytes = (BufferLengthMs / 1000) * .fxFormat.lAvgBytesPerSec
        End With
        
        On Error GoTo ErrorHandler
            Set clsSecBuf = m_clsDSound.CreateSoundBuffer(udtBufDsc)
        On Error GoTo 0

        Set clsStream = New DirectSoundStream

        ' initialize the new stream with the created secondary buffer
        CallCustomConstructor clsStream, udtBufDsc.lBufferBytes, VarPtr(clsSecBuf)
        
        Set stream = clsStream
        
        CreateStream = True
    Else
        CreateStream = False
    End If
    
    Exit Function
    
ErrorHandler:
    CreateStream = False
End Function

Public Function Deinitialize() As Boolean
    Set m_clsPrimBuf = Nothing
    Set m_clsDSound = Nothing
    
    Deinitialize = True
End Function

Public Function Initialize( _
    ByVal DeviceIndex As Long, _
    ByVal Samplerate As Long, _
    ByVal Channels As Integer, _
    ByVal BitsPerSample As Integer _
) As Boolean

    Dim udtBufDesc  As DSBUFFERDESC

    On Error GoTo ErrorHandler
        Set m_clsDSound = DirectX.DirectSoundCreate(DeviceGuid(DeviceIndex))
        m_clsDSound.SetCooperativeLevel m_hDSWnd, DSSCL_PRIORITY
    On Error GoTo 0
    
    With udtBufDesc
        With .fxFormat
            .lSamplesPerSec = Samplerate
            .nChannels = Channels
            .nBitsPerSample = BitsPerSample
            .nBlockAlign = .nChannels * (.nBitsPerSample \ 8)
            .lAvgBytesPerSec = .nBlockAlign * .lSamplesPerSec
            .nFormatTag = WAVE_FORMAT_PCM
        End With
        
        .lFlags = DSBCAPS_PRIMARYBUFFER
    End With
    
    On Error GoTo ErrorHandler
        Set m_clsPrimBuf = m_clsDSound.CreatePrimarySoundBuffer(udtBufDesc)
    On Error GoTo 0
    
    Initialize = True
    Exit Function
    
ErrorHandler:
    Set m_clsDSound = Nothing
    Initialize = False
End Function

Private Sub Class_Initialize()
    m_blnReady = InitializeDirectX()
    m_hDSWnd = CreateWindowEx(0, "static", "VB DS Stream", 0, 0, 0, 0, 0, 0, 0, 0, 0)
End Sub

Private Sub Class_Terminate()
    Set m_clsPrimBuf = Nothing
    Set m_clsDSound = Nothing
    
    If m_hDSWnd <> 0 Then
        DestroyWindow m_hDSWnd
    End If

    DeinitializeDirectX
End Sub
