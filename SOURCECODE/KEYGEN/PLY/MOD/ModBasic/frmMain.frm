VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "ModPlay v 0.1.4beta"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   FontTransparent =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmrTitle 
      Interval        =   1
      Left            =   3225
      Tag             =   "Scrolls the Title"
      Top             =   585
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C00000&
      Height          =   465
      Left            =   90
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   179
      TabIndex        =   6
      Top             =   45
      Width           =   2745
      Begin VB.Label lblMODtitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   270
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   555
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Dbg"
      Height          =   375
      Left            =   4965
      TabIndex        =   5
      Top             =   75
      Width           =   1050
   End
   Begin MSComctlLib.Slider sPos 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1110
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   661
      _Version        =   393216
      SelectRange     =   -1  'True
      TickFrequency   =   3
      TextPosition    =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   375
      Left            =   3915
      TabIndex        =   3
      Top             =   75
      Width           =   1050
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info"
      Height          =   375
      Left            =   2865
      TabIndex        =   1
      Top             =   75
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6015
      TabIndex        =   0
      Top             =   75
      Width           =   1050
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   1410
      Picture         =   "frmMain.frx":0442
      Stretch         =   -1  'True
      Top             =   585
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   810
      Picture         =   "frmMain.frx":061D
      Stretch         =   -1  'True
      Top             =   585
      Width           =   465
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   390
      Left            =   210
      Picture         =   "frmMain.frx":07D2
      Stretch         =   -1  'True
      Top             =   585
      Width           =   465
   End
   Begin VB.Label lblPlay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   165
      TabIndex        =   2
      Top             =   1830
      Width           =   45
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###############################################
'ModPlayer für VisualBasic
'Comment: The first ModPlayer in VB Code!
'Dev.: (c) Denis Wiegand
'Date 2002
'Sorry for my stupid English it`s translated from German :(

'Uses DirectX7 for VB.
'Sorry many of code is not correctly explained

'Note
'This modplayer is a beta development project of me.
'There a no Effects and many mod files wont work correctly.
'All samples are particularly in soundbuffers >31<, on
'some place while playing, an looped sample dont stop over
'the song because i have no "correct" channels, only
'31 buffers. I code very much to make the player good.

'This player play only 125Bpm 4 Channels mod correctly

'v 0.1beta

'STATUS
'SAMPLES    100%
'NOTES      (dont know i think ~ 95%)
'EFFECTS   (YES) 1%

'Variables
'Please give an explicit info, if you find errors. Thanks

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const GWL_EXSTYLE = (-20)



Dim dx As New DirectX7                                  'DirectX7 Objekt
Dim ds As DirectSound                                   'DirectSound (für playing)
Dim sound_buffer(31) As DirectSoundBuffer               'Sample Buffer
Dim WF As WAVEFORMATEX                                  'For all samples
Dim sound_desc(31) As DSBUFFERDESC                      'Needed: We are going to change "volume, pan, freq and position of sample"


Dim MODULE_NAME As String * 20                          'Module Title/Name
Dim Channels As Long                                    'Number of Channels (mostly mod`s use four)(this player suports only four)
Dim MOD_ID As String * 4                                'Module check id.

Dim SAMPLE_NAME(31) As String * 22                      'Sample name

Dim FILE_SAMPLE_LENGTH(31) As word                      'Sample length
Dim SAMPLE_LENGTH(31) As Long                           'converted from word to VB-long (see .bas files)

Dim FILE_FINE_TUNE(31) As Byte                          'FineTune???
Dim FINE_TUNE(31) As Long                               'converted from word to VB-long (see .bas files)

Dim VOLUME(31) As Byte                                  'Sample-Volume

Dim FILE_LOOP_START(31) As word                         'LoopPoint - start
Dim LOOP_START(31) As Long                              'converted from word to VB-long (see .bas files)

Dim FILE_LOOP_LENGTH(31) As word                        'LoopEnd = (LoopStart + LoopEnd)
Dim LOOP_LENGTH(31) As Long                             'converted from word to VB-long (see .bas files)

Dim SONG_LENGTH As Byte                                 'Song Length (indicated by in Patterns)
Dim NUMBER_PATTERNS As Long                             'Number of physical Patterns
Dim ORDERS(128) As Byte                                 'Pattern ordering through song.
Dim tempByte As Byte                                    'After SONG_LENGTH (In time, that was the loop point of module) (not used here)

'4 Bytes for the Pattern Data

Dim sn As Byte                                          'Sample Number
Dim PF As Byte                                          'Note
Dim EN As Byte                                          'Effect
Dim EP As Byte                                          'Effect Param.

'Read them all into arrays (c below in the Loader Sub)

Dim SAMPLE_NUMBER() As Byte                             ' ''
Dim PERIOD_FREQ() As Long                               ' ''
Dim EFFECT_NUMBER() As Long                             ' ''
Dim EFFECT_PARAMETER() As Long                          ' ''

Dim lt As Long

Const PAL = 7093789.2                                   'Play PAL
Const NTSC = 7159090.5                                  'Play NTSC (used here)
Dim tmp() As Byte                                       'One Sample buffer.

Dim row As Long                     'Pattern pos also named "row" max. 64
Dim patt As Long                    'Nummer of playing Pattern
'Dim Channel(5) As Long             'ignore not used here
Dim bPause As Boolean
Dim bStop As Boolean
Dim bPlay As Boolean

Private Sub Command1_Click()
    ExitProgramm
End Sub

Sub ExitProgramm()
    'Clean memory
    Set ds = Nothing
    Erase FILE_SAMPLE_LENGTH
    Erase SAMPLE_LENGTH
    Erase FILE_FINE_TUNE
    Erase FINE_TUNE
    Erase VOLUME
    Erase FILE_LOOP_START
    Erase LOOP_START
    Erase FILE_LOOP_LENGTH
    Erase LOOP_LENGTH
    Erase SAMPLE_NUMBER
    Erase PERIOD_FREQ
    Erase EFFECT_NUMBER
    Erase EFFECT_PARAMETER
    Erase tmp
    Erase sound_buffer
    End 'and byes
End Sub

Private Sub Command2_Click()
    'Display sample names in MsgBox
On Error Resume Next
    For i = 1 To 31
        Short$ = Short$ & i & vbTab & Mid(SAMPLE_NAME(i), 1, InStr(1, SAMPLE_NAME(i), Chr(0), vbTextCompare) - 1) & vbCrLf
    Next i
    MsgBox Short$
End Sub

Private Sub Command3_Click()
    'About box
    MsgBox "Programm by © Denis Wiegand (da breaker) @ 2002", vbInformation, "InfOO"
End Sub

Private Sub Command4_Click()
    frmDebug.Show
End Sub

Private Sub Form_Load()
    Me.Show         'Show the Form
    'Gradinet       'Blue smooth background
    Init            'Init DirectSound
    ChDir App.Path  'Change to out App. Path
    'If LoadMod("ChannelTest.mod") = True Then Play    'Load a mod and play them.
    If LoadMod("GroundZero.mod") = True Then Image1.Enabled = True Else Image1.Enabled = False
    'If LoadMod("GroundZero.mod") = true Then Play
End Sub

Sub Init()
    'Create a new Instance of DirectSound
    Set ds = dx.DirectSoundCreate("")               'Normal Soundadapter
    ds.SetCooperativeLevel Me.hwnd, DSSCL_NORMAL    'CooperativeLevel
    ds.SetSpeakerConfig DSSPEAKER_STEREO            'Stereo mode
    
    'Now set WAVEFORMATEX values
    '8360 = middle C (PAL)
    WF.nFormatTag = 1                               'PCM
    WF.nChannels = 1                                'Mono
    WF.lSamplesPerSec = 48100 '8360
    WF.nBitsPerSample = 8
    WF.nBlockAlign = 1
    WF.lAvgBytesPerSec = 48100 '8360
    WF.nSize = Len(WF)
End Sub

Sub DSLoadSamples(sn)
    'On Error Resume Next
    'Here become all samples an own
    'sound buffer ready to play :)
    
    If SAMPLE_LENGTH(sn) <> 0 Then
        'Sample len.
        sound_desc(sn).lBufferBytes = SAMPLE_LENGTH(sn)
        sound_desc(sn).lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN
        'Fresh sound buffer
        Set sound_buffer(sn) = ds.CreateSoundBuffer(sound_desc(sn), WF)
        'Signed Samples in Usigned umwandeln
        For i = 0 To UBound(tmp)
            tmp(i) = (tmp(i) + 128) And 255 'Thanks Olaf :)
        Next i
        
        'Write from the one sample tmp() to the sound buffer.
        sound_buffer(sn).WriteBuffer 0, UBound(tmp), tmp(0), DSBLOCK_DEFAULT
        sound_buffer(sn).SetVolume 0
    End If
End Sub

Function ScannMod(file As String) As Long
    'Every mod has its own id with contained the channles.
    
    Open file For Binary As #1
    Seek #1, 1081            'Jump to file pos 1081 (there are id)
    Get #1, , MOD_ID         'Read
    Close #1
    
    'Now check them out
    
    Select Case MOD_ID
        
        Case "M.K."
            ScannMod = 4 'Four channel mod
        Case "6CHN"
            ' ScannMod = 6 'Six channel mod  > not supported
        Case "8CHN"
            'ScannMod = 8 'Eight channel mod > not supported
            
    End Select
End Function

Function LoadMod(fname As String) As Boolean
    Channels = ScannMod(fname)      'Read ID
    If Channels = 0 Then            'if 0 > no valid mod
        LoadMod = False             'False
        MsgBox "Invalid mod file", vbCritical, "Error"  'Display that error.
        Exit Function 'Exit Loader function
    End If
    Open fname For Binary As #1
    Get #1, , MODULE_NAME 'Read title
    'Read all sample values
    For i = 1 To 31 'Samples
        Get #1, , SAMPLE_NAME(i)
        Get #1, , FILE_SAMPLE_LENGTH(i)
        Get #1, , FILE_FINE_TUNE(i)
        Get #1, , VOLUME(i)
        Get #1, , FILE_LOOP_START(i)
        Get #1, , FILE_LOOP_LENGTH(i)
    Next i
    'Ok now work on word(s)
    'Convert to VB longs (CinVB.bas)
    For i = 1 To 31
        SAMPLE_LENGTH(i) = MakeModWord(FILE_SAMPLE_LENGTH(i))
        LOOP_START(i) = MakeModWord(FILE_LOOP_START(i))
        LOOP_LENGTH(i) = MakeModWord(FILE_LOOP_LENGTH(i))
        If FILE_FINE_TUNE(i) > 7 Then
            FINE_TUNE(i) = FILE_FINE_TUNE(i) - 16
        Else
            FINE_TUNE(i) = FILE_FINE_TUNE(i)
        End If
    Next i
    'Now the song len.
    Get #1, , SONG_LENGTH
    Get #1, , tempByte 'Read that single own (used in latter versions)
    'Get the highest ORDER VALUE equal number of patterns in mod.
    For i = 1 To 128
        Get #1, , tempByte
        If tempByte > NUMBER_PATTERNS Then NUMBER_PATTERNS = tempByte
        ORDERS(i) = tempByte
    Next i
    'ReDim all pattern info arrays
    ReDim SAMPLE_NUMBER(NUMBER_PATTERNS, 64, Channels)
    ReDim PERIOD_FREQ(NUMBER_PATTERNS, 64, Channels)
    ReDim EFFECT_NUMBER(NUMBER_PATTERNS, 64, Channels)
    ReDim EFFECT_PARAMETER(NUMBER_PATTERNS, 64, Channels)
    'Now read all Patterns from Mod file
    Get #1, , MOD_ID 'Read id again because we are on pos. 1081
    For i = 0 To NUMBER_PATTERNS
        For p = 0 To (64 * Channels) - 1
            Get #1, , sn
            Get #1, , PF
            Get #1, , EN
            Get #1, , EP
            pp = pp + 4
            'The byte converting explanation is done in ModDoc.rtf
            SAMPLE_NUMBER(i, (p \ 4) + 1, (p Mod 4) + 1) = (sn And &HF0) + SHR(EN, 4)
            PERIOD_FREQ(i, (p \ 4) + 1, (p Mod 4) + 1) = (SHL((sn And &HF), 8) + PF)
            EFFECT_NUMBER(i, (p \ 4) + 1, (p Mod 4) + 1) = (EN And &HF)
            EFFECT_PARAMETER(i, (p \ 4) + 1, (p Mod 4) + 1) = (EP)
        Next p
    Next i
    
    'Load samples into DS.
    For i = 0 To 31
        If SAMPLE_LENGTH(i) <> 0 Then
            ReDim tmp(SAMPLE_LENGTH(i))         'Sample len
            Get #1, , tmp()                     'Read from file.
            DSLoadSamples i                     'Load to DS
        End If
    Next i
    Close #1
    
    'Display general information of mod
    ' lblFile.Caption = "File Name: " & fname
    lblMODtitle.Caption = MODULE_NAME
    lblMODtitle.Visible = True
    'lblChn.Caption = "Channels: " & Channels
    'lblLeng.Caption = "Länge: " & NUMBER_PATTERNS
    LoadMod = True 'Konnte ohne Fehler geladen werden
End Function

Sub Play()
    'The very simplest Play funktion of any ModPlayers :)
    'No Effects
    'No Correct Loops
    'This is a fresh error les sucess! :)
    'All posible errors removed!
    'revised version
    
    
On Error Resume Next 'For (not suported mods :( )
    
    Dim lastn(4) As Long                'Store note (we can not divide through zero)
    Dim lasts(4) As Long                ' ''
    Dim cu As DSCURSORS                 'DirectSoundBuffer pos
    Dim channel(4) As Long
    patt = 1                            'Play first Pattern
    sPos.Max = SONG_LENGTH        'For the NotWorking beta pos. changer :(
    bPlay = True
    
    Do While bPlay
        DoEvents    'Do anything while play!
        'Check for Pause
        While bPause
            DoEvents
            For i = 0 To 31
                If SAMPLE_LENGTH(i) > 0 Then
                    sound_buffer(i).Stop
                End If
            Next i
        Wend
        'Display row and pattern
        lblPlay.Caption = "ROW: " & row & " " & "PATTERN: " & ORDERS(patt)
        'lt = LastTick
        
        If lt <= dx.TickCount Then
            For i = 1 To Channels
                
                If PERIOD_FREQ(ORDERS(patt), row, i) > 0 Then lastn(i) = PERIOD_FREQ(ORDERS(patt), row, i)
                If SAMPLE_NUMBER(ORDERS(patt), row, i) > 0 Then lasts(i) = SAMPLE_NUMBER(ORDERS(patt), row, i)
                channel(i) = lasts(i)
                If lastn(i) > 0 Then sound_buffer(lasts(i)).SetFrequency (NTSC / (lastn(i) * 2))
                'Set Panning Left and Right
                
                If SAMPLE_LENGTH(lasts(i)) > 0 Then
                    If i = 1 Then sound_buffer(lasts(i)).SetPan -1000
                    If i = 2 Then sound_buffer(lasts(i)).SetPan -500
                    If i = 3 Then sound_buffer(lasts(i)).SetPan 500
                    If i = 4 Then sound_buffer(lasts(i)).SetPan 1000
                End If
                
                ef = EFFECT_NUMBER(ORDERS(patt), row, i)
                If ef = &HD Then
                    patt = patt + 1
                    row = EFFECT_PARAMETER(ORDERS(patt), row, i)
                    ef = 0
                End If
                
                'Debugger window
                frmDebug.note(i - 1).Text = lastn(i)
                frmDebug.sm(i - 1).Text = SAMPLE_NUMBER(ORDERS(patt), row, i)
                frmDebug.fxv(i - 1).Text = "0x" & Hex(ef)
                frmDebug.fxp(i - 1).Text = EFFECT_PARAMETER(ORDERS(patt), row, i)
                If lastn(i) > 0 Then frmDebug.notef(i - 1).Text = (NTSC / (lastn(i) * 2)) Else frmDebug.notef(i - 1).Text = ""
            Next i
            
            For i = 1 To Channels
                If SAMPLE_LENGTH(SAMPLE_NUMBER(ORDERS(patt), row, i)) > 0 Then
                    If LOOP_START(lasts(i)) > 0 Then
                        If PERIOD_FREQ(ORDERS(patt), row, i) > 0 Then sound_buffer(SAMPLE_NUMBER(ORDERS(patt), row, i)).SetCurrentPosition 0
                        If PERIOD_FREQ(ORDERS(patt), row, i) > 0 Then sound_buffer(SAMPLE_NUMBER(ORDERS(patt), row, i)).Play DSBPLAY_LOOPING
                    Else
                        If PERIOD_FREQ(ORDERS(patt), row, i) > 0 Then sound_buffer(SAMPLE_NUMBER(ORDERS(patt), row, i)).SetCurrentPosition 0
                        If PERIOD_FREQ(ORDERS(patt), row, i) > 0 Then sound_buffer(SAMPLE_NUMBER(ORDERS(patt), row, i)).Play DSBPLAY_DEFAULT
                    End If
                End If
            Next i
            
            
            row = row + 1   'Row um eins eröhen
            
            If patt > SONG_LENGTH Then
                patt = 0
            End If
            
            If row = 64 Then
                patt = patt + 1
                row = 0
                
            End If
            sPos.Value = patt
            lt = dx.TickCount + 110
        End If
        
        'If the player beginns, he reads the sample nr. and the note from the
        'pattern dat. he plays only the soundbuffer(sample) selected by the actual
        'row. If one sample is repeated (loop), it`s never stops until the next
        'command from the pattern plays this sample with an other note.
        'but it not stops!!! But when the actual sample is at present not used in
        'our four channels??
        'Here is my solution:
        'I have done this litle counter. This recognizes (not played samples)
        'only four channels.
        For s = 0 To 31
            c1 = channel(1)
            c2 = channel(2)
            c3 = channel(3)
            c4 = channel(4)
            'Stop the unused Samples
            'exclude playing used samples:
            If c1 <> s And c2 <> s And c3 <> s And c4 <> s Then
                If SAMPLE_LENGTH(s) > 0 Then sound_buffer(s).Stop 'exclude this and compare them
                'with this funktion and not :)
            End If
            's = sample compared with the aktual four plaing samples in channel(i)
            'We check every sample at once, to stop "not a played sample"
            'hope you have understand that
        Next s
        
        
        'Try to make loops correct (works mach times)
        
        For i = 0 To 31
            If LOOP_START(i) <> 0 Then
                sound_buffer(i).GetCurrentPosition cu
                If cu.lPlay > (LOOP_START(i) + LOOP_LENGTH(i)) Then
                    sound_buffer(i).SetCurrentPosition LOOP_START(i)
                    sound_buffer(i).Play DSBPLAY_LOOPING
                End If
            End If
        Next
        
    Loop
    'Stop all that samples
    For i = 0 To 31
        If SAMPLE_LENGTH(i) > 0 Then
            sound_buffer(i).Stop
        End If
    Next i
End Sub


Sub Gradinet()
    For i = 0 To frmMain.ScaleHeight
        col = RGB((i), (i), (i))
        Me.Line (0, i)-(Me.ScaleWidth, i), col
    Next i
    Me.Refresh
End Sub

Static Function Log10(X As Double) As Double
Log10 = Log(X) / Log(10)
End Function

Private Sub mnuDebug_Click()
    frmDebug.Show
End Sub

Private Sub Form_Terminate()
    ExitProgramm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ExitProgramm
End Sub

Private Sub Image1_Click()
    row = 0
    patt = 1
    sPos.Value = 0
    Play
End Sub

Private Sub Image2_Click()
    bPause = Not bPause
End Sub

Private Sub Image3_Click()
    bPlay = False
    sPos.Value = 0
    bPause = False
End Sub

Private Sub Picture1_Click()
    tmrTitle.Enabled = Not tmrTitle.Enabled
End Sub

Private Sub sPos_Click()
    'Works :)
    patt = sPos.Value
End Sub

Private Sub tmrTitle_Timer()
    Picture1.Cls
    lblMODtitle.Left = lblMODtitle.Left - 1
    If lblMODtitle.Left + lblMODtitle.Width < -10 Then
        lblMODtitle.Left = Picture1.Width + 10
    End If
End Sub
