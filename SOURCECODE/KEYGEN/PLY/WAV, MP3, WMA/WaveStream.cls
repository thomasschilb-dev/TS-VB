VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamWAV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' WAV Stream mit ACM Unterstützung
' für komprimierte WAV Dateien (z.B. ADPCM, GSM)

Implements ISoundStream
Implements ISoundInfo


Private Const Extensions    As String = "WAV"
Private Const Description   As String = "Wave Audio"


Private Declare Function acmStreamPrepareHeader Lib "msacm32" ( _
    ByVal has As Long, _
    pash As ACMSTREAMHEADER, _
    ByVal fdwPrepare As Long _
) As Long

Private Declare Function acmStreamUnprepareHeader Lib "msacm32" ( _
    ByVal has As Long, _
    pash As ACMSTREAMHEADER, _
    ByVal fdwUnprepare As Long _
) As Long

Private Declare Function acmStreamOpen Lib "msacm32" ( _
    phas As Long, _
    ByVal had As Long, _
    pwfxSrc As Any, _
    pwfxDst As Any, _
    ByVal pwfltr As Long, _
    ByVal dwCallback As Long, _
    ByVal dwInstance As Long, _
    ByVal fdwOpen As Long _
) As Long

Private Declare Function acmStreamSize Lib "msacm32" ( _
    ByVal has As Long, _
    ByVal cbInput As Long, _
    pdwOutputBytes As Long, _
    ByVal fdwSize As Long _
) As Long

Private Declare Function acmStreamConvert Lib "msacm32" ( _
    ByVal has As Long, _
    pash As ACMSTREAMHEADER, _
    ByVal fdwConvert As Long _
) As Long

Private Declare Function acmStreamReset Lib "msacm32" ( _
    ByVal has As Long, _
    ByVal fdwReset As Long _
) As Long

Private Declare Function acmStreamClose Lib "msacm32" ( _
    ByVal has As Long, _
    ByVal fdwClose As Long _
) As Long

Private Declare Function mmioClose Lib "winmm" ( _
    ByVal hmmio As Long, _
    ByVal uFlags As Long _
) As Long

Private Declare Function mmioDescend Lib "winmm" ( _
    ByVal hmmio As Long, _
    lpck As MMCKINFO, _
    lpckParent As MMCKINFO, _
    ByVal uFlags As Long _
) As Long

Private Declare Function mmioDescendParent Lib "winmm" _
Alias "mmioDescend" ( _
    ByVal hmmio As Long, _
    lpck As MMCKINFO, _
    ByVal x As Long, _
    ByVal uFlags As Long _
) As Long

Private Declare Function mmioOpen Lib "winmm" _
Alias "mmioOpenA" ( _
    ByVal szFileName As String, _
    lpmmioinfo As MMIOINFO, _
    ByVal dwOpenFlags As Long _
) As Long

Private Declare Function mmioSeek Lib "winmm" ( _
    ByVal hmmio As Long, _
    ByVal lOffset As Long, _
    ByVal iOrigin As Long _
) As Long

Private Declare Function mmioStringToFOURCC Lib "winmm" _
Alias "mmioStringToFOURCCA" ( _
    ByVal sz As String, _
    ByVal uFlags As Long _
) As Long

Private Type ACMSTREAMHEADER
    cbStruct                    As Long
    fdwStatus                   As Long
    dwUser                      As Long
    pbSrc                       As Long
    cbSrcLength                 As Long
    cbSrcLengthUsed             As Long
    dwSrcUser                   As Long
    pbDst                       As Long
    cbDstLength                 As Long
    cbDstLengthUsed             As Long
    dwDstUser                   As Long
    dwReservedDriver(9)         As Long
End Type

Private Type MMIOINFO
   dwFlags                      As Long
   fccIOProc                    As Long
   pIOProc                      As Long
   wErrorRet                    As Long
   htask                        As Long
   cchBuffer                    As Long
   pchBuffer                    As String
   pchNext                      As String
   pchEndRead                   As String
   pchEndWrite                  As String
   lBufOffset                   As Long
   lDiskOffset                  As Long
   adwInfo(4)                   As Long
   dwReserved1                  As Long
   dwReserved2                  As Long
   hmmio                        As Long
End Type

Private Type WAVE_FORMAT
    wFormatTag                  As Integer
    wChannels                   As Integer
    dwSampleRate                As Long
    dwBytesPerSec               As Long
    wBlockAlign                 As Integer
    wBitsPerSample              As Integer
End Type

Private Type MMCKINFO
   ckid                         As Long
   ckSize                       As Long
   fccType                      As Long
   dwDataOffset                 As Long
   dwFlags                      As Long
End Type

Private Type CHUNKINFO
    Start                       As Long
    Length                      As Long
End Type

Private Type WAVEFORMATEX
    wFormatTag                  As Integer
    nChannels                   As Integer
    nSamplesPerSec              As Long
    nAvgBytesPerSec             As Long
    nBlockAlign                 As Integer
    wBitsPerSample              As Integer
    cbSize                      As Integer
End Type

Private Enum HACMSTREAM
    INVALID_STREAM_HANDLE = 0
End Enum

Private Enum ACM_STREAMSIZEF
    ACM_STREAMSIZEF_DESTINATION = &H1
    ACM_STREAMSIZEF_SOURCE = &H0
    ACM_STREAMSIZEF_QUERYMASK = &HF
End Enum

Private Enum ACM_STREAMCONVERTF
    ACM_STREAMCONVERTF_BLOCKALIGN = &H4
    ACM_STREAMCONVERTF_START = &H10
    ACM_STREAMCONVERTF_END = &H20
End Enum

Private Const MMIO_READ         As Long = &H0
Private Const MMIO_FINDCHUNK    As Long = &H10
Private Const MMIO_FINDRIFF     As Long = &H20

Private Const SEEK_CUR          As Long = 1

Private Const WAVE_FORMAT_PCM   As Long = 1


' Stream

Private hFWave                  As hFile
Private cnkData                 As CHUNKINFO
Private cnkInfo                 As CHUNKINFO
Private udtWFXIn                As WAVEFORMATEX
Private udtWFXOut               As WAVEFORMATEX
Private btWFX()                 As Byte

' ACM

Private Const OUTPUT_BUFFER_MS  As Long = 500

Private hStream                 As HACMSTREAM

Private btInput()               As Byte
Private intOutput()             As Integer

Private lngInputLen             As Long
Private lngOutputLen            As Long

Private lngPosInBuffer          As Long
Private lngBufferData           As Long

Private blnEndOfStream          As Boolean
Private blnFirst                As Boolean

Private lngKeepInBuffer         As Long

Private Sub Class_Initialize()
    hStream = INVALID_STREAM_HANDLE
    hFWave.handle = INVALID_HANDLE
End Sub

Private Sub Class_Terminate()
    ISoundStream_StreamClose
End Sub

Private Property Get ISoundStream_Description( _
) As String

    ISoundStream_Description = Description
End Property

Private Property Get ISoundStream_EndOfStream( _
) As Boolean

    If blnEndOfStream Then
        If lngPosInBuffer = lngBufferData Then
            ISoundStream_EndOfStream = True
        End If
    End If
End Property

Private Function ISoundStream_Extensions( _
) As String()

    ISoundStream_Extensions = Split(Extensions, ";")
End Function

Private Property Get ISoundStream_StreamInfo( _
) As ISoundInfo

    Set ISoundStream_StreamInfo = Me
End Property

Private Function ISoundStream_StreamOpen( _
    ByVal SOURCE As String _
) As SND_RESULT

    If Not IsValidFile(SOURCE) Then
        ISoundStream_StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    ISoundStream_StreamClose

    ' WAV Chunks "data" und "fmt " finden
    cnkData = GetWavChunkPos(SOURCE, "data")
    cnkInfo = GetWavChunkPos(SOURCE, "fmt ")

    ' valide Chunks?
    If cnkData.Start = 0 Then
        ISoundStream_StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    If cnkInfo.Start = 0 Then
        ISoundStream_StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    If cnkInfo.Length < 16 Then
        ISoundStream_StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    hFWave = FileOpen(SOURCE, _
                      GENERIC_READ, _
                      FILE_SHARE_READ)

    ' Data Chunks mit illegaler Länge
    ' auf Dateigröße beschränken
    If FileLength(hFWave) < (cnkData.Start + cnkData.Length) Then
        cnkData.Length = FileLength(hFWave) - cnkData.Start
    End If

    If hFWave.handle = INVALID_HANDLE Then
        ISoundStream_StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    ' Info Chunk auslesen
    ReDim btWFX(cnkInfo.Length - 1) As Byte
    FileSeek hFWave, cnkInfo.Start, FILE_BEGIN
    FileRead hFWave, VarPtr(btWFX(0)), cnkInfo.Length

    CpyMem udtWFXIn, btWFX(0), Len(udtWFXIn)

    ' zum Start der eigentlichen Audio Daten seeken
    FileSeek hFWave, cnkData.Start, FILE_BEGIN

    ' ACM Konvertierung starten
    If Not InitConversion(True) Then
        ISoundStream_StreamOpen = SND_ERR_INTERNAL
        ISoundStream_StreamClose
        Exit Function
    End If

    ISoundStream_StreamOpen = SND_ERR_SUCCESS
End Function

Private Function ISoundStream_StreamClose( _
) As SND_RESULT

    If hFWave.handle = INVALID_HANDLE Then
        ISoundStream_StreamClose = SND_ERR_INVALID_SOURCE
    Else
        CloseConverter
        FileClose hFWave
        ISoundStream_StreamClose = SND_ERR_SUCCESS
    End If
End Function

Private Function ISoundStream_StreamRead( _
    ByVal buffer_ptr As Long, _
    ByVal buffer_len As Long, _
    ByRef buffer_read As Long _
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

                If 0 = IsBadReadPtr(ByVal VarPtr(intOutput(0)) + lngPosInBuffer, lngBufferData - lngPosInBuffer) Then
                    If 0 = IsBadWritePtr(ByVal buffer_ptr + buffer_read, lngBufferData - lngPosInBuffer) Then
                        CpyMem ByVal buffer_ptr + buffer_read, ByVal VarPtr(intOutput(0)) + lngPosInBuffer, lngBufferData - lngPosInBuffer
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
            If 0 = IsBadReadPtr(ByVal VarPtr(intOutput(0)) + lngPosInBuffer, buffer_len - buffer_read) Then
                If 0 = IsBadWritePtr(ByVal buffer_ptr + buffer_read, buffer_len - buffer_read) Then
                    CpyMem ByVal buffer_ptr + buffer_read, ByVal VarPtr(intOutput(0)) + lngPosInBuffer, buffer_len - buffer_read
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

    Dim lngBytes    As Long

    If hFWave.handle = INVALID_HANDLE Then
        ISoundStream_StreamSeek = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    Select Case seek_mode

        Case SND_SEEK_PERCENT
            If value < 0 Or value > 99 Then
                ISoundStream_StreamSeek = SND_ERR_OUT_OF_RANGE
                Exit Function
            End If

            lngBytes = value / 100 * cnkData.Length

        Case SND_SEEK_SECONDS
            If value < 0 Or value > (ISoundInfo_Duration / 1000) Then
                ISoundStream_StreamSeek = SND_ERR_OUT_OF_RANGE
                Exit Function
            End If

            lngBytes = udtWFXIn.nAvgBytesPerSec * value

    End Select

    If value = 0 Then
        FileSeek hFWave, cnkData.Start, FILE_BEGIN
    Else
        lngBytes = AlignBytes(lngBytes) + cnkData.Start
        FileSeek hFWave, lngBytes, FILE_BEGIN
    End If

    ' ACM Stream zurücksetzen, um
    ' Codec Cache zu leeren
    ResetConverter

    ISoundStream_StreamSeek = SND_ERR_SUCCESS
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''

Private Property Get ISoundInfo_Tags( _
) As StreamTags

    Set ISoundInfo_Tags = New StreamTags
End Property

Private Property Get ISoundInfo_Position( _
) As Long

    ISoundInfo_Position = (FilePosition(hFWave) - cnkData.Start) / udtWFXIn.nAvgBytesPerSec * 1000
End Property

Private Property Get ISoundInfo_Duration( _
) As Long

    ISoundInfo_Duration = (cnkData.Length) / udtWFXIn.nAvgBytesPerSec * 1000
End Property

Private Property Get ISoundInfo_BitsPerSample( _
) As Integer

    ISoundInfo_BitsPerSample = 16
End Property

Private Property Get ISoundInfo_BitsPerSecond( _
) As Long

    ISoundInfo_BitsPerSecond = udtWFXIn.nAvgBytesPerSec * 8
End Property

Private Property Get ISoundInfo_Channels( _
) As Integer

    ISoundInfo_Channels = udtWFXIn.nChannels
End Property

Private Property Get ISoundInfo_SamplesPerSecond( _
) As Long

    ISoundInfo_SamplesPerSecond = udtWFXIn.nSamplesPerSec
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''

Private Function InitConversion( _
    Optional Force16Bit As Boolean = False _
) As Boolean

    Dim mmr     As Long

    If hStream <> INVALID_STREAM_HANDLE Then
        CloseConverter
    End If

    udtWFXOut = udtWFXIn

    ' udtWFXOut hat jetzt noch alle Daten von
    ' udtWFXIn, bei komprimierten Formaten
    ' können die BPS kleiner als 8 sein,
    ' für die Ausgabe ist nur 8 oder 16 Bit
    ' erlaubt.
    ' Ebenso kann es sein, dass das Input Format
    ' 24 oder 32 (float) Bit Samples hat.
    If udtWFXOut.wBitsPerSample < 8 Then
        udtWFXOut.wBitsPerSample = 8
    ElseIf udtWFXOut.wBitsPerSample > 8 Then
        udtWFXOut.wBitsPerSample = 16
    End If

    If Force16Bit Then
        udtWFXOut.wBitsPerSample = 16
    End If

    With udtWFXOut
        udtWFXOut = CreateWFX(.nSamplesPerSec, _
                              .nChannels, _
                              .wBitsPerSample)
    End With

    mmr = acmStreamOpen(hStream, 0, _
                        btWFX(0), udtWFXOut, _
                        0, 0, 0, 0)

    If mmr <> 0 Then
        ' ACM Stream konnte nicht erstellt werden,
        ' vielleicht unterstützt der jeweilige
        ' Codec die BPS nicht.
        ' Versuchen, zu wechseln.
        If Force16Bit Then Exit Function

        If udtWFXOut.wBitsPerSample = 16 Then
            udtWFXOut.wBitsPerSample = 8
        Else
            udtWFXOut.wBitsPerSample = 16
        End If

        mmr = acmStreamOpen(hStream, 0, _
                            btWFX(0), udtWFXOut, _
                            0, 0, 0, 0)

        If mmr <> 0 Then Exit Function
    End If

    ' Größe des Output Buffers festlegen
    lngOutputLen = OUTPUT_BUFFER_MS / 1000 * udtWFXOut.nAvgBytesPerSec

    ' Benötigte Größe des Input Buffers für
    ' die Größe des Output Buffers über
    ' letztere feststellen
    mmr = acmStreamSize(hStream, _
                        lngOutputLen, _
                        lngInputLen, _
                        ACM_STREAMSIZEF_DESTINATION)

    If mmr <> 0 Then
        acmStreamClose hStream, 0
        hStream = 0
        Exit Function
    End If

    ReDim intOutput(lngOutputLen / 2 - 1) As Integer
    ReDim btInput(lngInputLen - 1) As Byte

    blnEndOfStream = False
    blnFirst = True

    lngKeepInBuffer = 0

    InitConversion = True
End Function

Private Function Convert( _
    ByVal lngInLen As Long, _
    ByVal lngOutLen As Long, _
    lngInUsed As Long, _
    lngOutUsed As Long, _
    Optional ByVal LastConversion As Boolean = False _
) As Boolean

    Dim lngFlags    As Long
    Dim udtHdr      As ACMSTREAMHEADER

    lngFlags = ACM_STREAMCONVERTF_BLOCKALIGN

    If blnFirst Then _
        lngFlags = lngFlags Or ACM_STREAMCONVERTF_START
    If LastConversion Then
        lngFlags = lngFlags Or ACM_STREAMCONVERTF_END
    End If

    ' Stream Header für Konvertierung erstellen
    With udtHdr
        .cbStruct = Len(udtHdr)
        .cbSrcLength = lngInLen
        .cbDstLength = lngOutLen
        .pbDst = VarPtr(intOutput(0))
        .pbSrc = VarPtr(btInput(0))
    End With

    acmStreamPrepareHeader hStream, udtHdr, 0

    blnFirst = False

    If 0 = acmStreamConvert(hStream, udtHdr, lngFlags) Then
        With udtHdr
            lngInUsed = .cbSrcLengthUsed
            lngOutUsed = .cbDstLengthUsed

            lngKeepInBuffer = .cbSrcLength - .cbSrcLengthUsed
        End With

        If lngKeepInBuffer > 0 Then
            ' Codec hat nicht alle Bytes angenommen,
            ' verbleibende Bytes im Buffer nach vorne
            ' ziehen, um sie bei der nächsten Konvertierung
            ' zu verwenden.
            CpyMem btInput(0), btInput(lngInLen - lngKeepInBuffer), lngKeepInBuffer
        End If

        Convert = True
    End If

    acmStreamUnprepareHeader hStream, udtHdr, 0
End Function

Private Function FillBuffer( _
) As Boolean

    Dim lngRead     As Long
    Dim lngWritten  As Long
    Dim udeRet      As SND_RESULT

    If blnEndOfStream Then
        'If lngPosInBuffer >= lngBufferData Then
            lngBufferData = 0
            lngPosInBuffer = 0
            ZeroMem intOutput(0), lngOutputLen
            ZeroMem btInput(0), lngInputLen
            Exit Function
        'Else
        '    FillBuffer = True
        '    Exit Function
        'End If
    End If

    ' Daten aus WAV Stream lesen
    udeRet = ReadWAVData(VarPtr(btInput(lngKeepInBuffer)), _
                         lngInputLen - lngKeepInBuffer, _
                         lngRead)

    If udeRet <> SND_ERR_SUCCESS Then
        ' Lesefehler oder Ende des Streams
        blnEndOfStream = True
    End If

    Convert lngRead + lngKeepInBuffer, lngOutputLen, _
            lngRead, lngWritten, _
            blnEndOfStream

    lngPosInBuffer = 0
    lngBufferData = lngWritten

    FillBuffer = True
End Function

' dreckige Variante, aber besonders bei
' Codecs, die tonnenweise cachen, effektiv
' beim Seeken
Private Sub ResetConverter()
    If hStream = INVALID_STREAM_HANDLE Then
        Exit Sub
    End If

    CloseConverter

    ' ACM Stream erstellen
    acmStreamOpen hStream, 0, _
                  btWFX(0), udtWFXOut, _
                  0, 0, 0, 0

    ' benötigte Bytes für OUT_BUFFER_MS
    ' Millisekunden PCM Samples
    lngOutputLen = OUTPUT_BUFFER_MS / 1000 * udtWFXOut.nAvgBytesPerSec
    acmStreamSize hStream, lngOutputLen, lngInputLen, ACM_STREAMSIZEF_DESTINATION

    ReDim intOutput(lngOutputLen / 2 - 1) As Integer
    ReDim btInput(lngInputLen - 1) As Byte

    blnEndOfStream = False
    blnFirst = True

    lngKeepInBuffer = 0
End Sub

Private Function CloseConverter( _
) As Boolean

    On Error Resume Next

    acmStreamClose hStream, 0
    hStream = INVALID_STREAM_HANDLE

    ZeroMem btInput(0), lngInputLen
    ZeroMem intOutput(0), lngOutputLen

    blnEndOfStream = False
    lngBufferData = 0
    lngPosInBuffer = 0
    lngInputLen = 0
    lngOutputLen = 0
    lngKeepInBuffer = 0

    CloseConverter = True
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''

Private Function ReadWAVData( _
    ByVal data_ptr As Long, _
    ByVal data_len As Long, _
    data_read As Long _
) As SND_RESULT

    ReadWAVData = SND_ERR_SUCCESS

    If hFWave.handle = INVALID_HANDLE Then
        ReadWAVData = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    If FilePosition(hFWave) > (cnkData.Start + cnkData.Length) Then
        ' Ende der Datei erreicht
        ReadWAVData = SND_ERR_END_OF_STREAM
        data_read = 0
        Exit Function
    End If

    If FilePosition(hFWave) + data_len > (cnkData.Start + cnkData.Length) Then
        ' noch nicht Ende der Datei, mit dem
        ' angeforderten Datenblock wird
        ' es aber überschritten
        data_len = (cnkData.Start + cnkData.Length) - FilePosition(hFWave)
        ReadWAVData = SND_ERR_END_OF_STREAM
    End If

    data_read = FileRead(hFWave, data_ptr, data_len)
End Function

' Position und Länge eines Chunks in einem
' WAV Container finden
Private Function GetWavChunkPos( _
    ByVal strFile As String, _
    ByVal strChunk As String _
) As CHUNKINFO

    Dim hMmioIn             As Long
    Dim lngRet              As Long
    Dim mmckinfoParentIn    As MMCKINFO
    Dim mmckinfoSubchunkIn  As MMCKINFO
    Dim mmioinf             As MMIOINFO

    ' WAV für Lesezugriff öffnen
    hMmioIn = mmioOpen(strFile, mmioinf, MMIO_READ)
    If hMmioIn = 0 Then
        Exit Function
    End If

    ' Auf Formatgültigkeit überprüfen
    mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
    lngRet = mmioDescendParent(hMmioIn, mmckinfoParentIn, 0, MMIO_FINDRIFF)
    If Not (lngRet = 0) Then
        mmioClose hMmioIn, 0
        Exit Function
    End If

    ' Chunk suchen
    mmckinfoSubchunkIn.ckid = mmioStringToFOURCC(strChunk, 0)
    lngRet = mmioDescend(hMmioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
    If Not (lngRet = 0) Then
        mmioClose hMmioIn, 0
        Exit Function
    End If

    ' Startposition und Länge des Chunks zurückgeben
    GetWavChunkPos.Start = mmioSeek(hMmioIn, 0, SEEK_CUR)
    GetWavChunkPos.Length = mmckinfoSubchunkIn.ckSize

    mmioClose hMmioIn, 0
End Function

' Beim Seeken in WAVs muss das Block Align
' beachtet werden, sonst können die Channels
' vertauscht werden oder man erhält nur noch
' Rauschen.
Private Function AlignBytes( _
    ByVal bytes As Long _
) As Long

    AlignBytes = bytes - (bytes Mod udtWFXIn.nBlockAlign)
End Function

' PCM WAVEFORMATEX aus Samplerate, Channels
' und BitsPerSample erstellen
Private Function CreateWFX( _
    sr As Long, _
    chs As Integer, _
    bps As Integer _
) As WAVEFORMATEX

    With CreateWFX
        .wFormatTag = WAVE_FORMAT_PCM
        .nChannels = chs
        .nSamplesPerSec = sr
        .wBitsPerSample = bps
        .nBlockAlign = chs * (bps / 8)
        .nAvgBytesPerSec = sr * .nBlockAlign
    End With
End Function
