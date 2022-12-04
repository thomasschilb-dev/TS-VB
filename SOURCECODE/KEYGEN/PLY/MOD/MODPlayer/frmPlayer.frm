VERSION 5.00
Begin VB.Form frmPlayer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MOD, IT, XM, and S3M Player"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4680
   Icon            =   "frmPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   4320
      Top             =   0
   End
   Begin VB.CommandButton cmdUnPause 
      Caption         =   "&UnPause"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "P&ause"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame frmStats 
      Caption         =   "Stats"
      Height          =   1935
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   4215
      Begin VB.Label lblGlobalVolume 
         AutoSize        =   -1  'True
         Caption         =   "Global Volume: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label lblPattern 
         AutoSize        =   -1  'True
         Caption         =   "Pattern: 0/0"
         Height          =   195
         Left            =   2160
         TabIndex        =   13
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "File Type: None"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1125
      End
      Begin VB.Label lblRow 
         AutoSize        =   -1  'True
         Caption         =   "Row: 0/0"
         Height          =   195
         Left            =   2160
         TabIndex        =   14
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         Caption         =   "Order: 0/0"
         Height          =   195
         Left            =   2160
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblElapsed 
         AutoSize        =   -1  'True
         Caption         =   "Elapsed Time: 0:00"
         Height          =   195
         Left            =   2160
         TabIndex        =   11
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Speed: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label lblBPM 
         AutoSize        =   -1  'True
         Caption         =   "Tempo(BPM): 0"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1110
      End
      Begin VB.Label lblMusicname 
         AutoSize        =   -1  'True
         Caption         =   "(No Music Loaded)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Label lblFilename 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(None)"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuRepeat 
         Caption         =   "&Repeat"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SongPointer As Long, PrevOrder As Byte
Public Function PhraseTime(TimeSec As Long) As String
    Dim Min As Byte, Sec As String
    Min = Fix(TimeSec / 60)
    Sec = CStr(TimeSec - (Min * 60))
    If Len(Sec) = 1 Then
        Sec = "0" & Sec
    End If
    PhraseTime = Min & ":" & Sec
End Function

Public Sub ShowCurrentStats()
    lblBPM.Caption = "Tempo(BPM): " & FMUSIC_GetBPM(SongPointer)
    lblSpeed.Caption = "Speed: " & FMUSIC_GetSpeed(SongPointer)
    lblElapsed.Caption = "Elapsed Time: " & PhraseTime(Int(FMUSIC_GetTime(SongPointer) / 1000))
    lblOrder.Caption = "Order: " & FMUSIC_GetOrder(SongPointer) + 1 & "/" & FMUSIC_GetNumOrders(SongPointer)
    lblPattern.Caption = "Pattern: " & FMUSIC_GetPattern(SongPointer) & "/" & FMUSIC_GetNumPatterns(SongPointer) - 1
    lblRow.Caption = "Row: " & FMUSIC_GetRow(SongPointer) & "/" & FMUSIC_GetPatternLength(SongPointer, FMUSIC_GetOrder(SongPointer)) - 1
End Sub

Private Sub cmdPause_Click()
    FMUSIC_SetPaused SongPointer, True
    cmdPause.Enabled = False
    cmdUnPause.Enabled = True
End Sub

Private Sub cmdPlay_Click()
    PrevOrder = 0
    FMUSIC_PlaySong SongPointer
    If FMUSIC_GetMasterVolume(SongPointer) < 72 Then
        FMUSIC_SetMasterVolume SongPointer, 140
    End If
End Sub

Private Sub cmdStop_Click()
    FMUSIC_StopSong SongPointer
    PrevOrder = 0
End Sub

Private Sub cmdUnPause_Click()
    FMUSIC_SetPaused SongPointer, False
    cmdPause.Enabled = True
    cmdUnPause.Enabled = False
End Sub

Private Sub Form_Load()
    FSOUND_Init 44001, 128, 0
End Sub

Private Sub Form_Terminate()
    FMUSIC_StopSong SongPointer
    FMUSIC_FreeSong SongPointer
    FSOUND_Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    FMUSIC_StopSong SongPointer
    FMUSIC_FreeSong SongPointer
    FSOUND_Close
    End
End Sub


Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuExit_Click()
    StopLoop = True
    FMUSIC_StopSong SongPointer
    FMUSIC_FreeSong SongPointer
    FSOUND_Close
    End
End Sub

Private Sub mnuOpen_Click()
    Dim OpenFile As OPENFILENAME, FileOpened As Boolean, FileName As String, TempPointer As Long, FileType As FMUSIC_TYPES
    OpenFile.lpstrFilter = "Module Files (*.MOD;*.IT;*.XM;*.S3M)" & Chr(0) & "*.MOD;*.IT;*,XM;*.S3M" & Chr(0) & "ProTracker Module (*.MOD)" & Chr(0) & "*.MOD" & Chr(0) & "Impulse Tracker Module (*.IT)" & Chr(0) & "*.IT" & Chr(0) & "eXtended Module (*.XM)" & Chr(0) & "*.XM" & Chr(0) & "ScreamTracker 3 Module (*.S3M)" & Chr(0) & "*.S3M" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
    OpenFile.lpstrTitle = "Open"
    OpenFile.lpstrInitialDir = App.Path
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = Me.hWnd
    OpenFile.hInstance = App.hInstance
    FileOpened = GetOpenFileName(OpenFile)
    If FileOpened = False Then
        Exit Sub
    End If
    FileName = Trim(OpenFile.lpstrFile)
    TempPointer = FMUSIC_LoadSong(FileName)
    If TempPointer = 0 Then
        MsgBox "Loading Failed"
        Exit Sub
    End If
    FMUSIC_StopSong SongPointer
    FMUSIC_FreeSong SongPointer
    lblFilename.Caption = FileName
    SongPointer = TempPointer
    lblMusicname.Caption = GetStringFromPointer(FMUSIC_GetName(SongPointer))
    lblGlobalVolume.Caption = "Global Volume: " & FMUSIC_GetGlobalVolume(SongPointer)
    FileType = FMUSIC_GetType(SongPointer)
    Select Case FileType
        Case Is = FMUSIC_TYPE_MOD
            lblType.Caption = "File Type: MOD"
        Case Is = FMUSIC_TYPE_IT
            lblType.Caption = "File Type: IT"
        Case Is = FMUSIC_TYPE_XM
            lblType.Caption = "File Type: XM"
        Case Is = FMUSIC_TYPE_S3M
            lblType.Caption = "File Type: S3M"
    End Select
    ShowCurrentStats
End Sub

Private Sub mnuRepeat_Click()
    If mnuRepeat.Checked = True Then
        mnuRepeat.Checked = False
    Else
        mnuRepeat.Checked = True
    End If
End Sub


Private Sub tmrUpdate_Timer()
    On Error Resume Next
    If FMUSIC_GetOrder(SongPointer) + 1 < PrevOrder Then
        If mnuRepeat.Checked = False Then
            PrevOrder = 0
            FMUSIC_StopSong SongPointer
        End If
    End If
    If FMUSIC_GetOrder(SongPointer) > PrevOrder Then
        PrevOrder = FMUSIC_GetOrder(SongPointer)
    End If
    ShowCurrentStats
End Sub


