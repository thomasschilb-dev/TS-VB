Attribute VB_Name = "DXMod"
Public dx                   As New DirectX8
Public ds                   As DirectSound8
Public buf()                As DirectSoundSecondaryBuffer8
Public channelBuf()         As DirectSoundSecondaryBuffer8

Public dsbd                 As DSBUFFERDESC
Public dsbd2                As DSBUFFERDESC

Public WaveFormat           As WAVEFORMATEX
Public dscaps               As DSBCAPS
Public dscur                As DSCURSORS

Function InitDX(DSID As String) As Boolean
'On Error GoTo FailedInit
        InitDX = True
        
        Set ds = dx.DirectSoundCreate(DSID)
        ds.SetCooperativeLevel Form2.hWnd, DSSCL_PRIORITY
                
        Exit Function

FailedInit:
Debug.Print "failed to init DX"
        InitDX = False
        Exit Function
End Function

Sub UnloadDx()
    Set ds = Nothing
    Set dx = Nothing
End Sub


