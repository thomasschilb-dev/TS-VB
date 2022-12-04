VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "DirectSound Streaming"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame frmTags 
      Caption         =   "Tags"
      Height          =   1665
      Left            =   300
      TabIndex        =   9
      Top             =   3375
      Width           =   3390
      Begin VB.ListBox lstTags 
         Height          =   1320
         IntegralHeight  =   0   'False
         Left            =   75
         TabIndex        =   10
         Top             =   225
         Width           =   3240
      End
   End
   Begin VB.Timer tmrPos 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2775
      Top             =   825
   End
   Begin MSComctlLib.Slider sldPos 
      Height          =   280
      Left            =   300
      TabIndex        =   7
      Top             =   1950
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   503
      _Version        =   393216
      TickStyle       =   3
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   3210
      Top             =   825
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.wav|*.wav"
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1425
      TabIndex        =   6
      Top             =   825
      Width           =   1065
   End
   Begin VB.ComboBox cboDevice 
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown-Liste
      TabIndex        =   5
      Top             =   180
      Width           =   2790
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   390
      Left            =   2625
      TabIndex        =   3
      Top             =   1350
      Width           =   1065
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   390
      Left            =   1425
      TabIndex        =   2
      Top             =   1350
      Width           =   1065
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   390
      Left            =   225
      TabIndex        =   1
      Top             =   1350
      Width           =   1065
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   390
      Left            =   225
      TabIndex        =   0
      Top             =   825
      Width           =   1065
   End
   Begin MSComctlLib.Slider sldVolume 
      Height          =   285
      Left            =   1050
      TabIndex        =   12
      Top             =   2550
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   503
      _Version        =   393216
      Min             =   -10000
      Max             =   0
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider sldBalance 
      Height          =   285
      Left            =   1050
      TabIndex        =   14
      Top             =   2925
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   503
      _Version        =   393216
      Min             =   -10000
      Max             =   10000
      TickStyle       =   3
   End
   Begin VB.Label lblBalance 
      AutoSize        =   -1  'True
      Caption         =   "Balance:"
      Height          =   195
      Left            =   315
      TabIndex        =   13
      Top             =   2925
      Width           =   615
   End
   Begin VB.Label lblVolume 
      AutoSize        =   -1  'True
      Caption         =   "Volume:"
      Height          =   195
      Left            =   375
      TabIndex        =   11
      Top             =   2550
      Width           =   570
   End
   Begin VB.Label lblPos 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "0:00/0:00"
      Height          =   195
      Left            =   2910
      TabIndex        =   8
      Top             =   2250
      Width           =   720
   End
   Begin VB.Label lblDevice 
      AutoSize        =   -1  'True
      Caption         =   "Device:"
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   225
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Buffer Length < 1000 can only be even hundreds
' (200, 400, 600, 800). I can't explain why,
' but for odd hundreds smaller streams can be
' played back wrong (missing or swapped buffers).
' The bigger the buffer, the better ;)
Private Const StreamBufferLength    As Long = 2000

Private WithEvents m_clsStream      As DirectSoundStream
Attribute m_clsStream.VB_VarHelpID = -1
Private m_clsDirectSound            As DirectSound
Private m_clsWaveFile               As ISoundStream

Private m_blnDoNotMove              As Boolean

Private Sub ShowDevices()
    Dim i   As Long
    
    cboDevice.Clear
    
    With m_clsDirectSound
        For i = 1 To .DeviceCount
            cboDevice.AddItem .DeviceDescription(i)
        Next
    End With
    
    If cboDevice.ListCount > 0 Then
        cboDevice.ListIndex = 0
    End If
End Sub

Private Sub cmdClose_Click()
    If Not m_clsStream Is Nothing Then
        m_clsStream.PlaybackStop
        Set m_clsStream = Nothing
    End If
    
    If Not m_clsWaveFile Is Nothing Then
        m_clsWaveFile.StreamClose
    End If
    
    If Not m_clsDirectSound Is Nothing Then
        m_clsDirectSound.Deinitialize
    End If
    
    cmdPlay.Enabled = False
    cmdPause.Enabled = False
    cmdStop.Enabled = False
    cmdClose.Enabled = False
    cmdOpen.Enabled = True
    cboDevice.Enabled = True
End Sub

Private Sub cmdOpen_Click()
    Dim clsStream   As DirectSoundStream
    Dim i           As Long

    On Error GoTo ErrorHandler
        dlgOpen.Filter = "wav, mp3, wma|*.wav;*.mp3;*.wma"
        dlgOpen.FileName = vbNullString
        dlgOpen.ShowOpen
        If dlgOpen.FileName = vbNullString Then Exit Sub
    On Error GoTo 0

    cmdClose_Click
    
    Select Case LCase$(Right$(dlgOpen.FileName, 3))
        Case "wav": Set m_clsWaveFile = New StreamWAV
        Case Else:  Set m_clsWaveFile = New StreamWMA
    End Select

    If Not m_clsWaveFile.StreamOpen(dlgOpen.FileName) = SND_ERR_SUCCESS Then
        MsgBox "Couldn't open the file!", vbExclamation
        Exit Sub
    End If

    With m_clsWaveFile.StreamInfo
        ' create the DirectSound Primary Buffer with the output format
        ' of the Wave Stream so there is no unnecessary resampling
        ' which makes playback faster
        If Not m_clsDirectSound.Initialize(cboDevice.ListIndex + 1, .SamplesPerSecond, .Channels, .BitsPerSample) Then
            MsgBox "Couldn't initialize DirectSound!", vbExclamation
            m_clsWaveFile.StreamClose
            Exit Sub
        End If

        ' create a DirectSound Secondary Buffer for playback
        If Not m_clsDirectSound.CreateStream(.SamplesPerSecond, .Channels, .BitsPerSample, StreamBufferLength, clsStream) Then
            MsgBox "Couldn't create DirectSound stream!", vbExclamation
            
            m_clsWaveFile.StreamClose
            m_clsDirectSound.Deinitialize
            
            Exit Sub
        Else
            Set m_clsStream = clsStream
            
            If m_clsWaveFile.StreamInfo.Duration = 0 Then
                sldPos.max = 1
            Else
                sldPos.max = m_clsWaveFile.StreamInfo.Duration
            End If
            
            sldPos.Min = 0
            sldPos.value = 0
            
            lstTags.Clear
            For i = 1 To m_clsWaveFile.StreamInfo.Tags.TagCount
                With m_clsWaveFile.StreamInfo.Tags.TagItem(i)
                    lstTags.AddItem .TagName & ": " & .TagValue
                End With
            Next
            
            m_clsStream.Volume = sldVolume.value
            m_clsStream.Balance = sldBalance.value
                        
            cmdClose.Enabled = True
            cmdOpen.Enabled = False
            cboDevice.Enabled = False
            cmdPlay.Enabled = True
            cmdPause.Enabled = False
            cmdStop.Enabled = False
        End If
    End With
    
ErrorHandler:
End Sub

Private Sub cmdPause_Click()
    If Not m_clsStream.PlaybackPause() Then
        MsgBox "Could not pause!", vbExclamation
    End If
End Sub

Private Sub cmdPlay_Click()
    Dim i           As Long
    Dim intData()   As Integer
    Dim lngDataSize As Long
    Dim lngRead     As Long
    
    If m_clsStream.PlaybackStatus = PlaybackStopped Then
        ' buffer 2 seconds of audio data
        lngDataSize = m_clsStream.BytesFromMs(200)
        ReDim intData(lngDataSize \ 2 - 1) As Integer
        
        For i = 1 To 10
            m_clsWaveFile.StreamRead VarPtr(intData(0)), lngDataSize, lngRead
            DoEvents
            
            If lngRead > 0 Then
                m_clsStream.AudioBufferAdd VarPtr(intData(0)), lngRead
            Else
                Exit For
            End If
        Next
    End If
    
    If Not m_clsStream.PlaybackStart() Then
        MsgBox "Could not start playback!", vbExclamation
    End If
End Sub

Private Sub cmdStop_Click()
    If Not m_clsStream.PlaybackStop() Then
        MsgBox "Could not stop playback!", vbExclamation
    End If
    
    m_clsWaveFile.StreamSeek 0, SND_SEEK_PERCENT
End Sub

Private Sub Form_Load()
    Set m_clsDirectSound = New DirectSound
    
    ProcessPrioritySet Priority:=ppHigh
    
    ShowDevices
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not m_clsStream Is Nothing Then
        m_clsStream.PlaybackStop
        Set m_clsStream = Nothing
    End If
    
    If Not m_clsWaveFile Is Nothing Then
        m_clsWaveFile.StreamClose
    End If
    
    If Not m_clsDirectSound Is Nothing Then
        m_clsDirectSound.Deinitialize
        Set m_clsDirectSound = Nothing
    End If
End Sub

Private Sub m_clsStream_BufferDone()
    Dim intData()   As Integer
    Dim lngDataSize As Long
    Dim lngRead     As Long
    
    If Not m_clsWaveFile.EndOfStream Then
        lngDataSize = m_clsStream.BytesFromMs(200)
        ReDim intData(lngDataSize \ 2 - 1) As Integer
    
        m_clsWaveFile.StreamRead VarPtr(intData(0)), lngDataSize, lngRead
        
        If lngRead > 0 Then
            m_clsStream.AudioBufferAdd VarPtr(intData(0)), lngRead
        End If
    End If
End Sub

Private Sub m_clsStream_NoDataLeft()
    Dim intData()   As Integer
    Dim lngDataSize As Long
    Dim lngRead     As Long
    Dim i           As Long
    
    If m_clsWaveFile.EndOfStream Then
        Debug.Print "End Of Stream!"
        m_clsStream.PlaybackStop
        m_clsWaveFile.StreamSeek 0, SND_SEEK_PERCENT
    Else
        ' buffer underrun, buffer 2 seconds of audio data
        lngDataSize = m_clsStream.BytesFromMs(200)
        ReDim intData(lngDataSize \ 2 - 1) As Integer
        
        For i = 1 To 10
            m_clsWaveFile.StreamRead VarPtr(intData(0)), lngDataSize, lngRead
            DoEvents
            
            If lngRead > 0 Then
                m_clsStream.AudioBufferAdd VarPtr(intData(0)), lngRead
            Else
                Exit For
            End If
        Next
    End If
End Sub

Private Sub m_clsStream_StatusChanged(ByVal status As PlaybackStatus)
    Select Case True
        Case status = PlaybackPlaying
            cmdPlay.Enabled = False
            cmdPause.Enabled = True
            cmdStop.Enabled = True
            tmrPos.Enabled = True
        Case status = PlaybackPausing
            cmdPlay.Enabled = True
            cmdPause.Enabled = False
            cmdStop.Enabled = True
            tmrPos.Enabled = False
        Case status = PlaybackStopped
            cmdPlay.Enabled = True
            cmdPause.Enabled = False
            cmdStop.Enabled = False
            tmrPos.Enabled = False
    End Select
End Sub

Private Sub sldBalance_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        sldBalance.value = 0
        
        If Not m_clsStream Is Nothing Then
            m_clsStream.Balance = 0
        End If
    End If
End Sub

Private Sub sldBalance_Scroll()
    If Not m_clsStream Is Nothing Then
        m_clsStream.Balance = sldBalance.value
    End If
End Sub

Private Sub sldPos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_blnDoNotMove = True
End Sub

Private Sub sldPos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngStreamPosition As Long

    m_blnDoNotMove = False
    
    m_clsWaveFile.StreamSeek sldPos.value \ 1000, SND_SEEK_SECONDS
    lngStreamPosition = m_clsWaveFile.StreamInfo.Position
    
    m_clsStream.PlaybackStop
    cmdPlay_Click
    m_clsStream.Elapsed = lngStreamPosition
End Sub

Private Sub sldVolume_Scroll()
    If Not m_clsStream Is Nothing Then
        m_clsStream.Volume = sldVolume.value
    End If
End Sub

Private Sub tmrPos_Timer()
    lblPos.Caption = MsToString(m_clsStream.Elapsed) & "/" & _
                     MsToString(m_clsWaveFile.StreamInfo.Duration)

    If Not m_blnDoNotMove Then
        sldPos.value = m_clsStream.Elapsed
    End If
End Sub

Private Function MsToString(ByVal ms As Long) As String
    Dim secs    As Long
    Dim mins    As Long
    
    secs = ms \ 1000
    mins = secs \ 60
    secs = secs Mod 60
    
    MsToString = mins & ":" & Format(secs, "00")
End Function
