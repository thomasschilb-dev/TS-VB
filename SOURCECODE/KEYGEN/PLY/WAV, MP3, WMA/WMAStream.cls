VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamWMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Windows Media Audio/MP3 Decoder
' (WMF SDK v9 Runtime oder WMP 9/10)

Implements ISoundStream
Implements ISoundInfo


Private Const Extensions As String = "WMA;MP3"
Private Const Description As String = "Windows Media Audio"


Private Declare Function CallWindowProc Lib "user32" _
Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

' WMF SDK Header Übersetzung (nicht vollständig)

Private Enum HRESULT
    S_OK = 0
End Enum

Private Enum WMT_VERSION
    WMT_VER_4_0 = &H40000
    WMT_VER_7_0 = &H70000
    WMT_VER_8_0 = &H80000
    WMT_VER_9_0 = &H90000
End Enum

Private Enum WMT_ATTR_DATATYPE
    WMT_TYPE_DWORD = 0
    WMT_TYPE_STRING = 1
    WMT_TYPE_BINARY = 2
    WMT_TYPE_BOOL = 3
    WMT_TYPE_QWORD = 4
    WMT_TYPE_WORD = 5
    WMT_TYPE_GUID = 6
End Enum

Private Enum WMT_ERRORS
    NS_E_NO_MORE_SAMPLES = &HC00D0BCF
End Enum

Private Enum WMT_RIGHTS
    WMT_RIGHT_PLAYBACK = &H1
    WMT_RIGHT_COPY_TO_NON_SDMI_DEVICE = &H2
    WMT_RIGHT_COPY_TO_CD = &H8
    WMT_RIGHT_COPY_TO_SDMI_DEVICE = &H10
    WMT_RIGHT_ONE_TIME = &H20
    WMT_RIGHT_SAVE_STREAM_PROTECTED = &H40
    WMT_RIGHT_COPY = &H80
    WMT_RIGHT_COLLABORATIVE_PLAY = &H100
    WMT_RIGHT_SDMI_TRIGGER = &H10000
    WMT_RIGHT_SDMI_NOMORECOPIES = &H20000
End Enum

Private Type MMWAVEFORMATEX
    wFormatTag                  As Integer
    nChannels                   As Integer
    nSamplesPerSec              As Long
    nAvgBytesPerSec             As Long
    nBlockAlign                 As Integer
    wBitsPerSample              As Integer
    cbSize                      As Integer
End Type

Private Type QWORD
    lo                          As Long
    hi                          As Long
End Type

Private Type GUID
    Data1                       As Long
    Data2                       As Integer
    Data3                       As Integer
    Data4(7)                    As Byte
End Type

Private Type IUnknown
    QueryInterface              As Long
    AddRef                      As Long
    Release                     As Long
End Type

Private Type IWMWriter
    IUnk                        As IUnknown
    SetProfileByID              As Long
    SetProfile                  As Long
    SetOutputFilename           As Long
    GetInputCount               As Long
    GetInputProps               As Long
    SetInputProps               As Long
    GetInputFormatCount         As Long
    GetInputFormat              As Long
    BeginWriting                As Long
    EndWriting                  As Long
    AllocateSample              As Long
    WriteSample                 As Long
    Flush                       As Long
End Type

Private Type INSSBuffer
    IUnk                        As IUnknown
    GetLength                   As Long
    SetLength                   As Long
    GetMaxLength                As Long
    GetBuffer                   As Long
    GetBufferAndLength          As Long
End Type

Private Type IWMHeaderInfo
    IUnk                        As IUnknown
    GetAttributeCount           As Long
    GetAttributeByIndex         As Long
    GetAttributeByName          As Long
    SetAttribute                As Long
    GetMarkerCount              As Long
    GetMarker                   As Long
    AddMarker                   As Long
    RemoveMarker                As Long
    GetScriptCount              As Long
    GetScript                   As Long
    AddScript                   As Long
    RemoveScript                As Long
End Type

Private Type IWMHeaderInfo2 ' : IWMHeaderInfo
    IUnk                        As IUnknown
    GetAttributeCount           As Long
    GetAttributeByIndex         As Long
    GetAttributeByName          As Long
    SetAttribute                As Long
    GetMarkerCount              As Long
    GetMarker                   As Long
    AddMarker                   As Long
    RemoveMarker                As Long
    GetScriptCount              As Long
    GetScript                   As Long
    AddScript                   As Long
    RemoveScript                As Long
    GetCodecInfoCount           As Long
    GetCodecInfo                As Long
End Type

Private Type IWMHeaderInfo3 ' : IWMHeaderInfo2
    IUnk                        As IUnknown
    GetAttributeCount           As Long
    GetAttributeByIndex         As Long
    GetAttributeByName          As Long
    SetAttribute                As Long
    GetMarkerCount              As Long
    GetMarker                   As Long
    AddMarker                   As Long
    RemoveMarker                As Long
    GetScriptCount              As Long
    GetScript                   As Long
    AddScript                   As Long
    RemoveScript                As Long
    GetCodecInfoCount           As Long
    GetCodecInfo                As Long
    GetAttributeCountEx         As Long
    GetAttributeIndices         As Long
    GetAttributeByIndexEx       As Long
    ModifyAttribute             As Long
    AddAttribute                As Long
    DeleteAttribute             As Long
    AddCodecInfo                As Long
End Type

Private Type IWMMediaProps
    IUnk                        As IUnknown
    GetType                     As Long
    GetMediaType                As Long
    SetMediaType                As Long
End Type

Private Type IWMInputMediaProps ' : IWMMediaProps
    IUnk                        As IUnknown
    GetType                     As Long
    GetMediaType                As Long
    SetMediaType                As Long
    GetConnectionName           As Long
    GetGroupName                As Long
End Type

Private Type IWMProfileManager
    IUnk                        As IUnknown
    CreateEmtpyProfile          As Long
    LoadProfileByID             As Long
    LoadProfileByData           As Long
    SaveProfile                 As Long
    GetSystemProfileCount       As Long
    LoadSystemProfile           As Long
End Type

Private Type IWMProfile
    IUnk                        As IUnknown
    GetVersion                  As Long
    GetName                     As Long
    SetName                     As Long
    GetDescription              As Long
    SetDescription              As Long
    GetStreamCount              As Long
    GetStream                   As Long
    GetStreamByNumber           As Long
    RemoveStream                As Long
    RemoveStreamByNumber        As Long
    AddStream                   As Long
    ReconfigStream              As Long
    CreateNewStream             As Long
    GetMutualExclusionCount     As Long
    AddMutualExclusion          As Long
    CreateNewMutualExclusion    As Long
End Type

Private Type IWMProfile2 ' : IWMProfile
    IUnk                        As IUnknown
    GetVersion                  As Long
    GetName                     As Long
    SetName                     As Long
    GetDescription              As Long
    SetDescription              As Long
    GetStreamCount              As Long
    GetStream                   As Long
    GetStreamByNumber           As Long
    RemoveStream                As Long
    RemoveStreamByNumber        As Long
    AddStream                   As Long
    ReconfigStream              As Long
    CreateNewStream             As Long
    GetMutualExclusionCount     As Long
    AddMutualExclusion          As Long
    CreateNewMutualExclusion    As Long
    GetProfileID                As Long
End Type

Private Type IWMProfile3
    IUnk                        As IUnknown
    GetVersion                  As Long
    GetName                     As Long
    SetName                     As Long
    GetDescription              As Long
    SetDescription              As Long
    GetStreamCount              As Long
    GetStream                   As Long
    GetStreamByNumber           As Long
    RemoveStream                As Long
    RemoveStreamByNumber        As Long
    AddStream                   As Long
    ReconfigStream              As Long
    CreateNewStream             As Long
    GetMutualExclusionCount     As Long
    AddMutualExclusion          As Long
    CreateNewMutualExclusion    As Long
    GetProfileID                As Long
    GetStorageFormat            As Long
    SetStorageFormat            As Long
    GetBandwidthSharingCount    As Long
    GetBandwidthSharing         As Long
    RemoveBandwidthSharing      As Long
    AddBandwidthSharing         As Long
    CreateNewBandwidthSharing   As Long
    GetStreamPrioritization     As Long
    SetStreamPrioritization     As Long
    RemoveStreamPrioritization  As Long
    CreateNewStreamPrioritization As Long
    GetExpectedPacketCount      As Long
End Type

Private Type IWMCodecInfo
    IUnk                        As IUnknown
    GetCodecInfoCount           As Long
    GetCodecFormatCount         As Long
    GetCodecFormat              As Long
End Type

Private Type IWMCodecInfo2
    IUnk                        As IUnknown
    GetCodecInfoCount           As Long
    GetCodecFormatCount         As Long
    GetCodecFormat              As Long
    GetCodecName                As Long
    GetCodecFormatDesc          As Long
End Type

Private Type IWMStreamConfig
    IUnk                        As IUnknown
    GetStreamType               As Long
    GetStreamNumber             As Long
    SetStreamNumber             As Long
    GetStreamName               As Long
    SetStreamName               As Long
    GetConnectionName           As Long
    SetConnectionName           As Long
    GetBitrate                  As Long
    SetBitrate                  As Long
    GetBufferWindow             As Long
    SetBufferWindow             As Long
End Type

Private Type IWMOutputMediaProps
    IUnk                        As IUnknown
    GetType                     As Long
    GetMediaType                As Long
    SetMediaType                As Long
    GetStreamGroupName          As Long
    GetConnectionName           As Long
End Type

Private Type IWMSyncReader
    IUnk                        As IUnknown
    Open                        As Long
    Close                       As Long
    SetRange                    As Long
    SetRangeByFrame             As Long
    GetNextSample               As Long
    SetStreamsSelected          As Long
    GetStreamSelected           As Long
    SetReadStreamSamples        As Long
    GetReadStreamSamples        As Long
    GetOutputSetting            As Long
    SetOutputSetting            As Long
    GetOutputCount              As Long
    GetOutputProps              As Long
    SetOutputProps              As Long
    GetOutputFormatCount        As Long
    GetOutputFormat             As Long
    GetOutputNumberForStream    As Long
    GetStreamNumberForOutput    As Long
    GetMaxOutputSampleSize      As Long
    GetMaxStreamSampleSize      As Long
    OpenStream                  As Long
End Type

Private Type WM_MEDIA_TYPE
    majortype                   As GUID
    subtype                     As GUID
    bFixedSizeSamples           As Long
    bTemporalCompression        As Long
    lSampleSize                 As Long
    formattype                  As GUID
    pUnk                        As Long
    cbFormat                    As Long
    pbFormat                    As Long
End Type


Private Const WMMEDIATYPE_Audio      As String _
    = "{73647561-0000-0010-8000-00AA00389B71}"

Private Const WMMEDIASUBTYPE_PCM     As String _
    = "{00000001-0000-0010-8000-00AA00389B71}"


Private Const WMFORMAT_WaveFormatEx  As String _
    = "{05589f81-c356-11ce-bf01-00aa0055595a}"


Private Const IID_IWMHeaderInfo      As String _
    = "{96406BDA-2B2B-11d3-B36B-00C04F6108FF}"

Private Const IID_IWMHeaderInfo2     As String _
    = "{15CF9781-454E-482e-B393-85FAE487A810}"

Private Const IID_IWMHeaderInfo3     As String _
    = "{15CC68E3-27CC-4ecd-B222-3F5D02D80BD5}"


Private Const IID_IWMProfileManager  As String _
    = "{d16679f2-6ca0-472d-8d31-2f5d55aee155}"

Private Const IID_IWMProfileManager2 As String _
    = "{7A924E51-73C1-494d-8019-23D37ED9B89A}"


Private Const IID_IWMCodecInfo       As String _
    = "{A970F41E-34DE-4a98-B3BA-E4B3CA7528F0}"

Private Const IID_IWMCodecInfo2     As String _
    = "{AA65E273-B686-4056-91EC-DD768D4DF710}"


Private Const IID_IWMMediaProps      As String _
    = "{96406BCE-2B2B-11d3-B36B-00C04F6108FF}"


Private Const WMMEDIASUBTYPE_DRM     As String _
    = "{00000009-0000-0010-8000-00AA00389B71}"



Private Const IID_IUnknown           As String _
    = "{00000000-0000-0000-C000-000000000046}"

Private Const IID_IWMOutputMediaProps As String _
    = "{96406BD7-2B2B-11d3-B36B-00C04F6108FF}"

Private Const IID_IWMSyncReader      As String _
    = "{9397F121-7705-4dc9-B049-98B698188414}"

Private Const IID_INSSBuffer         As String _
    = "{E1CD3524-03D7-11d2-9EED-006097D2D7CF}"


Private Const attr_WMDuration    As String = "Duration"
Private Const attr_WMBitrate     As String = "Bitrate"
Private Const attr_WMSeekable    As String = "Seekable"
Private Const attr_WMStridable   As String = "Stridable"
Private Const attr_WMBroadcast   As String = "Broadcast"
Private Const attr_WMProtected   As String = "Is_Protected"
Private Const attr_WMTrusted     As String = "Is_Trusted"
Private Const attr_WMSigName     As String = "Signature_Name"
Private Const attr_WMHasAudio    As String = "HasAudio"
Private Const attr_WMHasImage    As String = "HasImage"
Private Const attr_WMHasScript   As String = "HasScript"
Private Const attr_WMHasVideo    As String = "HasVideo"
Private Const attr_WMCurBitrate  As String = "CurrentBitrate"
Private Const attr_WMOptBitrate  As String = "OptimalBitrate"
Private Const attr_WMSkipBackw   As String = "Can_Skip_Backward"
Private Const attr_WMSkipForw    As String = "Can_Skip_Forward"
Private Const attr_WMNumFrames   As String = "NumberOfFrames"
Private Const attr_WMFileSize    As String = "FileSize"

Private Const attr_WMTitle       As String = "Title"
Private Const attr_WMAuthor      As String = "Author"
Private Const attr_WMDescript    As String = "Description"
Private Const attr_WMRating      As String = "Rating"
Private Const attr_WMCopyright   As String = "Copyright"

Private Const attr_WMAlbumTitle  As String = "WM/AlbumTitle"
Private Const attr_WMAlbumArtist As String = "WM/AlbumArtist"
Private Const attr_WMTrack       As String = "WM/Track"
Private Const attr_WMGenre       As String = "WM/Genre"
Private Const attr_WMYear        As String = "WM/Year"
Private Const attr_WMGenreID     As String = "WM/GenreID"
Private Const attr_WMMCDI        As String = "WM/MCDI"
Private Const attr_WMComposer    As String = "WM/Composer"
Private Const attr_WMLyrics      As String = "WM/Lyrics"
Private Const attr_WMTrackNumber As String = "WM/TrackNumber"
Private Const attr_WMIsVBR       As String = "IsVBR"

' Synchronous WMF Reader
Private IReader                 As IWMSyncReader
Private oReader                 As Long
Private pReaderVTbl             As Long

' Media Properties
Private IProps                  As IWMOutputMediaProps
Private oProps                  As Long
Private pPropsVTbl              As Long

' Media Header
Private IHeader                 As IWMHeaderInfo3
Private oHeader                 As Long
Private pHeaderVTbl             As Long

' waveOut WAVEFORMATEX
Private wfxout                  As MMWAVEFORMATEX

' WMA Attribute
Private curDuration             As Currency
Private curBitrate              As Currency
Private curPosition             As Currency
Private blnSeekable             As Boolean
Private blnEOS                  As Boolean

' PCM Buffer
Private btBuffer()              As Byte
Private lngBufferData           As Long
Private lngPosInBuffer          As Long

' WM Core DLL
Private clsWMCore               As DynamicLibrary

Private blnFileOpened           As Boolean

Private clsTags                 As StreamTags

''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''

Private Function FillBuffer( _
) As Boolean

    Dim lngRet          As Long
    Dim ISample         As INSSBuffer
    Dim oSample         As Long
    Dim pSampleVTbl     As Long
    Dim curSampleDur    As Currency
    Dim pdwFlags        As Long
    Dim pdwOutputNum    As Long
    Dim ppdwBuffer      As Long
    Dim pdwLength       As Long

    If oReader = 0 Then Exit Function

    ' ein Stück der Datei decodieren
    lngRet = CallPointer(IReader.GetNextSample, _
                         oReader, _
                         0, _
                         VarPtr(oSample), _
                         VarPtr(curPosition), _
                         VarPtr(curSampleDur), _
                         VarPtr(pdwFlags), _
                         VarPtr(pdwOutputNum), _
                         0)

    If lngRet <> 0 Then
        blnEOS = lngRet = NS_E_NO_MORE_SAMPLES
        Exit Function
    End If

    ' INSSBuffer's VTable
    CpyMem pSampleVTbl, ByVal oSample, 4
    CpyMem ISample, ByVal pSampleVTbl, Len(ISample)

    ' Pointer zu den Samples und deren Länge holen
    lngRet = CallPointer(ISample.GetBufferAndLength, _
                         oSample, _
                         VarPtr(ppdwBuffer), _
                         VarPtr(pdwLength))

    If lngRet <> 0 Then
        lngRet = CallPointer(ISample.IUnk.Release, oSample)
        Exit Function
    End If

    ' PCM Buffer füllen
    ReDim btBuffer(pdwLength - 1) As Byte
    CpyMem btBuffer(0), ByVal ppdwBuffer, pdwLength

    ' INSSBuffer zerstören
    lngRet = CallPointer(ISample.IUnk.Release, oSample)

    ' Buffer Descriptors zurücksetzen
    lngPosInBuffer = 0
    lngBufferData = pdwLength

    FillBuffer = True
End Function

Private Function GUID2Str( _
    g As GUID _
) As String

    Dim nTemp   As String

    nTemp = "{"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data1), 8) & "-"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data2), 4) & "-"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data3), 4) & "-"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(0)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(1)), 2) & "-"
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(2)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(3)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(4)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(5)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(6)), 2)
    nTemp = nTemp & FmtStrLen(Hex$(g.Data4(7)), 2) & "}"

    GUID2Str = nTemp
End Function

Private Function GUIDEqual( _
    g1 As GUID, _
    g2 As GUID _
) As Boolean

    GUIDEqual = GUID2Str(g1) = GUID2Str(g2)
End Function

Private Function Str2GUID( _
    strGUID As String _
) As GUID

    Dim g           As GUID
    Dim strParts()  As String

    strParts = Split(Mid$(strGUID, 2, Len(strGUID) - 2), "-")

    g.Data1 = Val("&H" & strParts(0))
    g.Data2 = Val("&H" & strParts(1))
    g.Data3 = Val("&H" & strParts(2))
    g.Data4(0) = Val("&H" & Mid$(strParts(3), 1, 2))
    g.Data4(1) = Val("&H" & Mid$(strParts(3), 3, 2))
    g.Data4(2) = Val("&H" & Mid$(strParts(4), 1, 2))
    g.Data4(3) = Val("&H" & Mid$(strParts(4), 3, 2))
    g.Data4(4) = Val("&H" & Mid$(strParts(4), 5, 2))
    g.Data4(5) = Val("&H" & Mid$(strParts(4), 7, 2))
    g.Data4(6) = Val("&H" & Mid$(strParts(4), 9, 2))
    g.Data4(7) = Val("&H" & Mid$(strParts(4), 11, 2))

    Str2GUID = g
End Function

Private Function RemNullChars( _
    ByVal strString As String _
) As String

    If InStr(strString, Chr$(0)) > 0 Then
        RemNullChars = Left$(strString, InStr(strString, Chr$(0)) - 1)
    Else
        RemNullChars = strString
    End If
End Function

Private Sub CpyAttribute( _
    pDst As Long, _
    pSrc As Long, _
    dt As WMT_ATTR_DATATYPE _
)

    Dim nSize   As Long

    Select Case dt
        Case WMT_TYPE_BOOL: nSize = 2
        Case WMT_TYPE_DWORD: nSize = 4
        Case WMT_TYPE_GUID: nSize = 16
        Case WMT_TYPE_QWORD: nSize = 8
        Case WMT_TYPE_WORD: nSize = 2
    End Select

    CpyMem ByVal pDst, ByVal pSrc, nSize
End Sub

Private Function FmtStrLen( _
    ByVal B As String, _
    length As Integer _
) As String

    B = String(length - Len(B), "0") & B
    FmtStrLen = B
End Function

' http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=49987&lngWId=1
Private Function DecToHex( _
    ByVal nSource As Double _
) As String

    Const BASECHAR  As String = "0123456789ABCDEF"
    Dim n           As Double
    Dim nSrc        As Double

    On Error GoTo ErrorLine

    If (nSource = 0) Then
        DecToHex = "00"
        Exit Function
    End If

    If (nSource < 2147483648#) Then
        DecToHex = Hex(nSource)
    Else
        nSrc = nSource

        Do
            n = CDec(nSrc - (16 * Int(nSrc / 16)))
            DecToHex = Mid$(BASECHAR, n + 1, 1) & DecToHex
            nSrc = CDec(Int(nSrc / 16))
        Loop While (nSrc > 0)

    End If

    If (Len(DecToHex) Mod 2) Then
        DecToHex = "0" & DecToHex
    End If
    Exit Function

ErrorLine:
    DecToHex = ""
End Function

Private Function CallPointer( _
    ByVal fnc As Long, _
    ParamArray Params() _
) As Long

    Dim btASM(&HEC00& - 1)  As Byte
    Dim pASM                As Long
    Dim i                   As Integer

    If fnc = 0 Then Exit Function

    pASM = VarPtr(btASM(0))

    AddByte pASM, &H58                  ' POP EAX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H50                  ' PUSH EAX

    If UBound(Params) = 0 Then
        If IsArray(Params(0)) Then
            For i = UBound(Params(0)) To 0 Step -1
                AddPush pASM, CLng(Params(0)(i))    ' PUSH dword
            Next
        Else
            For i = UBound(Params) To 0 Step -1
                AddPush pASM, CLng(Params(i))       ' PUSH dword
            Next
        End If
    Else
        For i = UBound(Params) To 0 Step -1
            AddPush pASM, CLng(Params(i))           ' PUSH dword
        Next
    End If

    AddCall pASM, fnc                   ' CALL rel addr
    AddByte pASM, &HC3                  ' RET

    CallPointer = CallWindowProc(VarPtr(btASM(0)), _
                                 0, 0, 0, 0)
End Function

Private Sub AddPush( _
    pASM As Long, _
    lng As Long _
)

    AddByte pASM, &H68
    AddLong pASM, lng
End Sub

Private Sub AddCall( _
    pASM As Long, _
    addr As Long _
)

    AddByte pASM, &HE8
    AddLong pASM, addr - pASM - 4
End Sub

Private Sub AddLong( _
    pASM As Long, _
    lng As Long _
)

    CpyMem ByVal pASM, lng, 4
    pASM = pASM + 4
End Sub

Private Sub AddByte( _
    pASM As Long, _
    Bt As Byte _
)

    CpyMem ByVal pASM, Bt, 1
    pASM = pASM + 1
End Sub

Private Sub Class_Initialize()
    Dim lngRet  As Long

    Set clsWMCore = New DynamicLibrary
    Set clsTags = New StreamTags

    If Not clsWMCore.LoadDLL("wmvcore.dll") Then
        Exit Sub
    End If

    ' IWMSyncReader Instanz erstellen
    lngRet = clsWMCore.CallFunc("WMCreateSyncReader", 0, WMT_RIGHT_PLAYBACK, VarPtr(oReader))
    If lngRet <> 0 Then
        oReader = 0
        Exit Sub
    End If

    ' IWMSyncReader`s VTable
    CpyMem pReaderVTbl, ByVal oReader, 4
    CpyMem IReader, ByVal pReaderVTbl, Len(IReader)
End Sub

Private Sub Class_Terminate()
    Dim lngRet  As Long

    ISoundStream_StreamClose

    ' IWMSyncReader zerstören
    ' (Set IWMSyncReader = Nothing)
    If oReader <> 0 Then
        CallPointer IReader.IUnk.Release, oReader
    End If

    clsWMCore.UnloadDLL
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''

Private Property Get ISoundInfo_BitsPerSample( _
) As Integer

    ISoundInfo_BitsPerSample = wfxout.wBitsPerSample
End Property

Private Property Get ISoundInfo_BitsPerSecond( _
) As Long

    ISoundInfo_BitsPerSecond = curBitrate * 10000
End Property

Private Property Get ISoundInfo_Channels( _
) As Integer

    ISoundInfo_Channels = wfxout.nChannels
End Property

Private Property Get ISoundInfo_Duration( _
) As Long

    ISoundInfo_Duration = curDuration
End Property

Private Property Get ISoundInfo_Position( _
) As Long

    ISoundInfo_Position = curPosition
End Property

Private Property Get ISoundInfo_SamplesPerSecond( _
) As Long

    ISoundInfo_SamplesPerSecond = wfxout.nSamplesPerSec
End Property

Private Property Get ISoundInfo_Tags( _
) As StreamTags

    Set ISoundInfo_Tags = clsTags
End Property

''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''

Private Property Get ISoundStream_Description( _
) As String

    ISoundStream_Description = Description
End Property

Private Property Get ISoundStream_EndOfStream( _
) As Boolean

    ISoundStream_EndOfStream = blnEOS
End Property

Private Function ISoundStream_Extensions( _
) As String()

    ISoundStream_Extensions = Split(Extensions, ";")
End Function

Private Function ISoundStream_StreamClose( _
) As SND_RESULT

    Dim lngRet  As Long

    If oReader = 0 Then
        ISoundStream_StreamClose = SND_ERR_INTERNAL
        Exit Function
    End If

    If Not blnFileOpened Then
        ISoundStream_StreamClose = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    lngRet = CallPointer(IReader.Close, oReader)
    blnFileOpened = False

    blnEOS = False

    clsTags.ClearTags

    ISoundStream_StreamClose = SND_ERR_SUCCESS
End Function

Private Property Get ISoundStream_StreamInfo( _
) As ISoundInfo

    Set ISoundStream_StreamInfo = Me
End Property

Private Function ISoundStream_StreamOpen( _
    ByVal SOURCE As String _
) As SND_RESULT

    Dim lngRet  As Long

    ' haben wir eine IWMSyncReader Instanz?
    If oReader = 0 Then
        ISoundStream_StreamOpen = SND_ERR_INTERNAL
        Exit Function
    End If

    If Not IsValidFile(SOURCE) Then
        ISoundStream_StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    ' Datei öffnen
    lngRet = CallPointer(IReader.Open, _
                         oReader, _
                         StrPtr(SOURCE))

    If lngRet <> 0 Then
        ISoundStream_StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    If Not GetOutputFormat() Then
        CallPointer IReader.Close, oReader
        ISoundStream_StreamOpen = SND_ERR_UNKNOWN
        Exit Function
    End If

    ' zu decodierendes Gebiet (0 bis EOF)
    lngRet = CallPointer(IReader.SetRangeByFrame, _
                         oReader, _
                         1, 0, 0, _
                         0, 0)

    If lngRet <> 0 Then
        lngRet = CallPointer(IReader.Close, oReader)
        ISoundStream_StreamOpen = SND_ERR_UNKNOWN
        Exit Function
    End If

    If Not GetAttributes Then
        lngRet = CallPointer(IReader.Close, oReader)
        ISoundStream_StreamOpen = SND_ERR_UNKNOWN
        Exit Function
    End If

    blnFileOpened = True

    ISoundStream_StreamOpen = SND_ERR_SUCCESS
End Function

Private Function GetAttributes( _
) As Boolean

    Dim lngRet      As Long
    Dim cAttrs      As Long
    Dim attIndex    As Long
    Dim cchName     As Long
    Dim cbValue     As Long

    Dim pwszName    As String
    Dim strValue    As String

    Dim pbValue()   As Byte

    Dim attType     As WMT_ATTR_DATATYPE

    Dim iid_hdr     As GUID

    clsTags.ClearTags

    ' IWMHeaderInfo3 aus IWMSyncReader quetschen
    iid_hdr = Str2GUID(IID_IWMHeaderInfo3)
    lngRet = CallPointer(IReader.IUnk.QueryInterface, _
                         oReader, _
                         VarPtr(iid_hdr), _
                         VarPtr(oHeader))

    ' IWMHeaderInfo3`s VTable
    CpyMem pHeaderVTbl, ByVal oHeader, 4
    CpyMem IHeader, ByVal pHeaderVTbl, Len(IHeader)

    ' Anzahl an Attributen
    lngRet = CallPointer(IHeader.GetAttributeCountEx, _
                         oHeader, _
                         &HFFFF&, _
                         VarPtr(cAttrs))

    If lngRet <> 0 Then
        lngRet = CallPointer(IHeader.IUnk.Release, oHeader)
        oHeader = 0
        Exit Function
    End If

    ' Durch alle Attribute loopen
    For attIndex = 0 To cAttrs - 1

        ' Größe des Attributnamens und des Inhalts
        lngRet = CallPointer(IHeader.GetAttributeByIndexEx, _
                             oHeader, _
                             &HFFFF&, _
                             attIndex, _
                             0, _
                             VarPtr(cchName), _
                             0, 0, 0, _
                             VarPtr(cbValue))

        If lngRet <> 0 Then GoTo NextAttribute

        ' Buffer für den Namen und den Inhalt erstellen
        pwszName = Space(cchName)
        ReDim pbValue(cbValue - 1) As Byte

        ' Attribut vollständig auslesen
        lngRet = CallPointer(IHeader.GetAttributeByIndexEx, _
                             oHeader, _
                             &HFFFF&, _
                             attIndex, _
                             StrPtr(pwszName), _
                             VarPtr(cchName), _
                             VarPtr(attType), _
                             0, _
                             VarPtr(pbValue(0)), _
                             VarPtr(cbValue))

        If lngRet <> 0 Then GoTo NextAttribute

        pwszName = RemNullChars(pwszName)

        ' alle Attribute in die Tag Collection stopfen
        Select Case attType
            Case WMT_TYPE_WORD                  ' integer
                Dim intVal  As Integer
                CpyMem intVal, pbValue(0), 2
                strValue = intVal
            Case WMT_TYPE_DWORD                 ' long
                Dim lngVal  As Long
                CpyMem lngVal, pbValue(0), 4
                strValue = lngVal
            Case WMT_TYPE_QWORD                 ' currency?
                Dim curVal  As Currency
                CpyMem curVal, pbValue(0), 8
                strValue = curVal
            Case WMT_TYPE_STRING                ' Unicode String
                strValue = RemNullChars(pbValue)
            Case WMT_TYPE_BINARY                ' binary...
                ' nicht lesbar...
                ' Debug.Print "<binary>"
            Case WMT_TYPE_BOOL                  ' boolean
                CpyMem lngVal, pbValue(0), cbValue
                strValue = CBool(lngVal)
            Case WMT_TYPE_GUID                  ' global unique identifier
                Dim g   As GUID
                CpyMem g, pbValue(0), Len(g)
                strValue = GUID2Str(g)
        End Select

        ' Attribute, die wir speichern wollen
        Select Case pwszName

            Case attr_WMDuration
                CpyAttribute VarPtr(curDuration), _
                             VarPtr(pbValue(0)), _
                             attType

            Case attr_WMBitrate
                CpyAttribute VarPtr(curBitrate), _
                             VarPtr(pbValue(0)), _
                             attType

            Case attr_WMSeekable
                CpyAttribute VarPtr(blnSeekable), _
                             VarPtr(pbValue(0)), _
                             attType

            Case attr_WMAuthor
                clsTags.AddTag "artist", strValue

            Case attr_WMAlbumArtist
                clsTags.AddTag "artist", strValue

            Case attr_WMAlbumTitle
                clsTags.AddTag "album", strValue

            Case attr_WMGenre
                clsTags.AddTag "genre", strValue

            Case attr_WMTitle
                clsTags.AddTag "title", strValue

            ' Track Nummer wird oft als DWORD
            ' gespeichert (laut MS), obwohl
            ' WM/Track(Number) ein String Attribut
            ' ist.
            Case attr_WMTrackNumber
                If lngVal < 100 And lngVal > -1 Then
                    clsTags.AddTag "tracknumber", CStr(lngVal)
                Else
                    clsTags.AddTag "tracknumber", strValue
                End If

            Case attr_WMTrackNumber
                If lngVal < 100 And lngVal > -1 Then
                    clsTags.AddTag "tracknumber", CStr(lngVal)
                Else
                    clsTags.AddTag "tracknumber", strValue
                End If

            Case attr_WMDescript
                clsTags.AddTag "comment", strValue

            Case attr_WMYear
                clsTags.AddTag "year", strValue

        End Select

NextAttribute:
    Next

    ' IWMHeaderInfo3 zerstören
    lngRet = CallPointer(IHeader.IUnk.Release, oHeader)
    oHeader = 0

    GetAttributes = True
End Function

Private Function GetOutputFormat( _
) As Boolean

    Dim lngRet      As Long
    Dim pcOutputs   As Long
    Dim pcFormats   As Long
    Dim pcbType     As Long
    Dim i           As Long
    Dim j           As Long
    Dim MediaType   As WM_MEDIA_TYPE

    Dim btMT()      As Byte

    ' Anzahl an Outputs
    lngRet = CallPointer(IReader.GetOutputCount, _
                         oReader, _
                         VarPtr(pcOutputs))

    If lngRet <> 0 Then
        lngRet = CallPointer(IReader.Close, oReader)
        Exit Function
    End If

    ' Formate jedes Outputs
    For i = 0 To pcOutputs - 1

        ' Anzahl an Formaten für den aktuellen Output
        lngRet = CallPointer(IReader.GetOutputFormatCount, _
                             oReader, _
                             i, _
                             VarPtr(pcFormats))

        If lngRet <> 0 Then GoTo NextOutput

        ' Details für jedes Format
        For j = 0 To pcFormats

            ' IWMOutputMediaProps Instanz erstellen
            lngRet = CallPointer(IReader.GetOutputFormat, _
                                 oReader, _
                                 j, i, _
                                 VarPtr(oProps))

            If lngRet <> 0 Then GoTo NextFormat

            ' IWMOutputMediaProps`s VTable
            CpyMem pPropsVTbl, ByVal oProps, 4
            CpyMem IProps, ByVal pPropsVTbl, Len(IProps)

            ' Es ist nicht erlaubt,
            ' einfach direkt WM_MEDIA_TYPE in
            ' eine Struktur anzufordern.
            ' Man muss erst die Größe von WM_MEDIA_TYPE
            ' abfragen, einen ausreichend großen Buffer
            ' erstellen, und die Daten dann in
            ' eine WM_MEDIA_TYPE Struktur kopieren.

            ' Größe von WM_MEDIA_TYPE
            lngRet = CallPointer(IProps.GetMediaType, _
                                 oProps, _
                                 0, _
                                 VarPtr(pcbType))

            If lngRet <> 0 Then
                lngRet = CallPointer(IProps.IUnk.Release, oProps)
                oProps = 0
                GoTo NextFormat
            Else
                ReDim btMT(pcbType - 1) As Byte
            End If

            ' Daten in eine WM_MEDIA_TYPE Struktur kopieren
            lngRet = CallPointer(IProps.GetMediaType, _
                                 oProps, _
                                 VarPtr(btMT(0)), _
                                 VarPtr(pcbType))

            CpyMem MediaType, btMT(0), Len(MediaType)

            ' Audio?
            If GUIDEqual(MediaType.majortype, Str2GUID(WMMEDIATYPE_Audio)) Then
                ' PCM?
                If GUIDEqual(MediaType.subtype, Str2GUID(WMMEDIASUBTYPE_PCM)) Then
                    ' 16 Bit?
                    CpyMem wfxout, ByVal MediaType.pbFormat, Len(wfxout)
                    If wfxout.wBitsPerSample = 16 Then
                        lngRet = CallPointer(IProps.SetMediaType, oProps, VarPtr(btMT(0)))
                        GoTo ExitLoops
                    End If
                End If
            End If

            ' IWMOutputMediaProps zerstören
            lngRet = CallPointer(IProps.IUnk.Release, oProps)
            oProps = 0

NextFormat:         ' Nicht schön,
        Next j

NextOutput:         ' aber macht verschachtelte
    Next i
ExitLoops:          ' Loops zum Kinderspiel.

    ' Kein Format gefunden?
    If i = pcOutputs Then
        Exit Function
    End If

    ' neues Output Format setzen
    lngRet = CallPointer(IReader.SetOutputProps, oReader, i, oProps)
    lngRet = CallPointer(IProps.IUnk.Release, oProps)
    oProps = 0

    GetOutputFormat = True
End Function

Private Function ISoundStream_StreamRead( _
    ByVal buffer_ptr As Long, _
    ByVal buffer_len As Long, _
    buffer_read As Long _
) As SND_RESULT

    ISoundStream_StreamRead = SND_ERR_SUCCESS

    buffer_read = 0

    ' decodieren, solange nicht alle
    ' angeforderten Bytes übergeben wurden
    ' oder Ende des Streams erreicht worden ist
    Do While buffer_read < buffer_len

        ' PCM Buffer ist leer
        If lngBufferData = 0 Then
            If Not FillBuffer Then
                ISoundStream_StreamRead = SND_ERR_END_OF_STREAM
                Exit Function
            End If

        ' Daten im PCM Buffer reichen nicht,
        ' um verbleibende Bytes zu füllen
        ElseIf (lngBufferData - lngPosInBuffer) < (buffer_len - buffer_read) Then
            If 0 < (lngBufferData - lngPosInBuffer) Then

                If 0 = IsBadReadPtr(ByVal VarPtr(btBuffer(0)) + lngPosInBuffer, lngBufferData - lngPosInBuffer) Then
                    If 0 = IsBadWritePtr(ByVal buffer_ptr + buffer_read, lngBufferData - lngPosInBuffer) Then
                        CpyMem ByVal buffer_ptr + buffer_read, ByVal VarPtr(btBuffer(0)) + lngPosInBuffer, lngBufferData - lngPosInBuffer
                    End If
                End If

                buffer_read = buffer_read + (lngBufferData - lngPosInBuffer)
            End If

            If Not FillBuffer Then
                ISoundStream_StreamRead = SND_ERR_END_OF_STREAM
                Exit Function
            End If

        ' genügend Daten im PCM Buffer,
        ' um verbleibende Bytes zu übergeben
        Else
            If 0 = IsBadReadPtr(ByVal VarPtr(btBuffer(0)) + lngPosInBuffer, buffer_len - buffer_read) Then
                If 0 = IsBadWritePtr(ByVal buffer_ptr + buffer_read, buffer_len - buffer_read) Then
                    CpyMem ByVal buffer_ptr + buffer_read, ByVal VarPtr(btBuffer(0)) + lngPosInBuffer, buffer_len - buffer_read
                End If
            End If

            lngPosInBuffer = lngPosInBuffer + (buffer_len - buffer_read)
            buffer_read = buffer_read + (buffer_len - buffer_read)

        End If

    Loop
End Function

Private Function ISoundStream_StreamSeek( _
    ByVal value As Long, _
    ByVal seek_mode As SND_SEEK_MODE _
) As SND_RESULT

    Dim curTime     As QWORD, curTime0  As QWORD
    Dim dblTime     As Double, strHex   As String
    Dim lngRet      As Long

    ' IWMSyncReader::SetRange will QWORDS
    ' (64-bit integers).
    ' Da VB die nicht unterstützt,
    ' (naja, Currency, aber das scheint
    ' hier nicht so recht hinzuhauen),
    ' rechnen wir den Wert einfach über
    ' einen Double ins hexadezimale um,
    ' und splitten das dann in 2 DWORDS.
    ' Danke an Cristian Calugar für seine
    ' Dbl2Hex Funktion:
    ' http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=49987&lngWId=1

    Select Case seek_mode

        Case SND_SEEK_PERCENT
            If value < 0 Or value > 99 Then
                ISoundStream_StreamSeek = SND_ERR_OUT_OF_RANGE
                Exit Function
            End If

            ' Position in 100-Nanosekunden berechnen
            dblTime = Fix(curDuration / CDbl(100) * CDbl(value) * 10000#)

        Case SND_SEEK_SECONDS
            If value < 0 Or value > (ISoundInfo_Duration / 1000) Then
                ISoundStream_StreamSeek = SND_ERR_OUT_OF_RANGE
                Exit Function
            End If

            ' Position in 100-Nanosekunden berechnen
            dblTime = Fix(curDuration / CDbl(100) * CDbl(value / (ISoundInfo_Duration / 1000) * 100) * 10000#)

    End Select

    ' von Base 10 nach Base 16
    strHex = FmtStrLen(DecToHex(dblTime), 16)

    ' lo DWORD
    curTime.lo = Val("&H" & Right$(strHex, 8) & "&")
    ' hi DWORD
    curTime.hi = Val("&H" & Left$(strHex, 8) & "&")

    ' neue Position setzen
    lngRet = CallPointer(IReader.SetRange, oReader, _
                         curTime.lo, curTime.hi, _
                         curTime0.lo, curTime0.hi)

    If lngRet = 0 Then
        FillBuffer
        ISoundStream_StreamSeek = SND_ERR_SUCCESS
    Else
        ISoundStream_StreamSeek = SND_ERR_OUT_OF_RANGE
    End If
End Function
