VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8520
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   2295
      Left            =   2760
      Max             =   20000
      Min             =   20
      TabIndex        =   12
      Top             =   0
      Value           =   15000
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Play2"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox DaLoop 
      Caption         =   "Use the LAME loop system"
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   2280
      Width           =   2835
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show commentary"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show Samples"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "stop"
      Height          =   255
      Left            =   1380
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton play 
      Caption         =   "Play"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.HScrollBar length 
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   2580
      Width           =   8415
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      ItemData        =   "Form2.frx":0000
      Left            =   3060
      List            =   "Form2.frx":0002
      TabIndex        =   2
      Top             =   60
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "show instruments"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton XmLoad 
      Caption         =   "Load XM"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2655
   End
   Begin VB.Label Label2 
      Height          =   195
      Left            =   5700
      TabIndex        =   10
      Top             =   2340
      Width           =   2055
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   60
      TabIndex        =   8
      Top             =   1140
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'######################################################################################################################################################
'# this xm player is still very early in it's development, lots of things has not yet been implanted
'# it does not support 16-samples yet (I removed them since they all seemed to be half in length ????
'# I know there's alot of spelling error and grammar misstakes (I would be supprised if not)
'# and that some things just isn't described very well. so please bear with me. hopefully they will be corrected someday :)
'#
'# if you have made an improvment, a bug-fix or just has some ideas don't hesitate to drop me a mail (no snail mails)
'# at lars_akesson_1@hotmail.com . And please state a describing subject otherwise it will be removed considerd as spam.
'######################################################################################################################################################

Implements DirectXEvent8
Dim timerbuf As DirectSoundSecondaryBuffer8
Dim notify As Long
Dim msize As Long
Dim emptybuf() As Byte

Private Sub Command5_Click()
        'On Error Resume Next
        msize = 100
        Me.Caption = (xm.BPM * 2 / 5)
        ReDim emptybuf(msize)
        Set timerbuf = Nothing
       ' Set Buf2(0) = Nothing
        
        dsbd.lFlags = DSBCAPS_CTRLFREQUENCY + DSBCAPS_CTRLPAN + DSBCAPS_CTRLVOLUME + DSBCAPS_STATIC + DSBCAPS_GLOBALFOCUS + DSBCAPS_CTRLPOSITIONNOTIFY
        
        With dsbd.fxFormat
            .nFormatTag = WAVE_FORMAT_PCM
            .nChannels = 1
            .lSamplesPerSec = 44100
            .nBitsPerSample = 8
            .nBlockAlign = .nBitsPerSample / 8 * .nChannels
            .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
            .nSize = 0
        End With
        
        dsbd.lBufferBytes = msize
        
        Set timerbuf = ds.CreateSoundBuffer(dsbd)
        
        timerbuf.WriteBuffer 0, msize, emptybuf(0), DSBLOCK_DEFAULT
        
        ChangeEvent msize
        timerbuf.play DSBPLAY_LOOPING
End Sub

Sub ChangeEvent(value As Long)
      '  On Error GoTo err
        If notify <> 0 Then
           dx.DestroyEvent notify
        End If
        
        notify = dx.CreateEvent(Me)
        
        Dim psa(1) As DSBPOSITIONNOTIFY

        psa(0).hEventNotify = notify
        psa(0).lOffset = value - 1
        
        timerbuf.SetNotificationPositions 1, psa()
      '  Exit Sub
'err:
        'Debug.Print debugs.DisplayError(err.Number, DirectSound)
End Sub

'######################################################################################################################################################
'# This is test play routine, just to se if i can get the darn song to play properly, (and not so CPU intensive)
'# right speed and all (the right speed has not yet been implanted. atleast i don't think so :)  )
'######################################################################################################################################################
Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)
On Error Resume Next
        
        If eventid = notify Then
           playXM.play2
        End If
End Sub
Private Sub Command1_Click()
        On Error Resume Next
        
        List1.Clear
        
        For i = 1 To UBound(ih)
            List1.AddItem i & vbTab & ih(i).name
        Next
End Sub

Private Sub Command2_Click()
        StopSong = True
        timerbuf.Stop
        On Error Resume Next
        For i = 0 To xm.channels - 1
        channelBuf(i).Stop
        Next
End Sub

Private Sub Command3_Click()
        List1.Clear
        Dim tmpstr As String
        Dim tmpstr2 As String
        
        For i = 1 To Len(commentary)
            tmpstr = tmpstr & Mid(commentary, i, 1)
            
            If Mid(commentary, i, 1) = vbCr Then
            tmpstr2 = Mid$(tmpstr, 1, Len(tmpstr) - 1)
            List1.AddItem tmpstr2
            tmpstr = ""
            End If
        Next
End Sub

Private Sub Command4_Click()
        On Error Resume Next
        
        List1.Clear
        
        For i = 1 To UBound(sh)
            List1.AddItem i & vbTab & sh(i).name
        Next
End Sub

Private Sub Form_Load()
        MsgBox "This software is only a beta " & Chr(13) & "be cautious when listening to xm files containing 16-bit samples since they're not supported yet", vbOKOnly, "this is just a test"
        DXMod.InitDX "{00000000-0000-0000-0000-000000000000}" 'change this into an enumerated GUID.
        debugMod.openlog
End Sub

Private Sub Form_Unload(Cancel As Integer)
        StopSong = True
        DXMod.UnloadDx
        debugMod.savelog
        End
End Sub

Private Sub length_Change()
        playXM.SetPosition length.value
        Label2.Caption = "playing pattern nr: " & xm.Order(length.value)
End Sub

Private Sub play_Click()
        playXM.play
End Sub

Private Sub VScroll1_Change()
On Error Resume Next
        timerbuf.SetFrequency VScroll1.value * 5
End Sub

Private Sub VScroll1_Scroll()
        VScroll1_Change
End Sub

Private Sub XmLoad_Click()
        Dim FileName As String
        Const Title As String = "Load a XM Music file"
        Const filt As String = "Music *.xm|*.xm"

        cd.ShowOpen FileName, Title, filt
        
        loadxm.load_file FileName
        
        Me.Caption = xm.name
        


        
        length.Max = xm.patterns
        If Len(FileName) <> 0 Then
        If warning16 = True Then
           Label1.Caption = "Warning this song contains 16-bit samples and so far this program DOES NOT support 16-bit samples. This will result in the creation of a loud noise. Keep your volume down to prevent your speakers from being damaged."
           Label1.BackColor = RGB(255, 100, 100)
        Else
           Label1.Caption = "This song contains 8-bit samples only."
           Label1.BackColor = RGB(100, 255, 100)
        End If
        End If
End Sub
