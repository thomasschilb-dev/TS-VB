Attribute VB_Name = "tsMIDIMod"
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal _
    lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength _
    As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal _
    fdwError As Long, ByVal lpszErrorText As String, ByVal cchErrorText As Long) As Long

Public Sub OpenMID(MIDFile As String, MIDAlias As String)
Dim ErrCode As Long
ErrCode = mciSendString("open " & MIDFile & " alias " & MIDAlias, "", 0, 0)
If ErrCode <> 0 Then DisplayError ErrCode
End Sub

Public Sub PlayMID(MIDAlias As String)
Dim ErrCode As Long

ErrCode = mciSendString("play " & MIDAlias, "", 0, 0)
If ErrCode <> 0 Then DisplayError ErrCode
End Sub

Public Sub StopMID(MIDAlias As String)
Dim ErrCode As Long
    
ErrCode = mciSendString("stop " & MIDAlias, "", 0, 0)
If ErrCode <> 0 Then DisplayError ErrCode
End Sub

Public Sub CloseMID(MIDAlias As String)
Dim ErrCode As Long
ErrCode = mciSendString("close " & MIDAlias, "", 0, 0)
End Sub

Private Sub DisplayError(ByVal ErrCode As Long)
Dim errstr As String
Dim retval As Long

errstr = Space(128)
retval = mciGetErrorString(ErrCode, errstr, Len(errstr))
errstr = Left(errstr, InStr(errstr, vbNullChar) - 1)

retval = MsgBox(errstr, vbOKOnly Or vbCritical, "MCI")
End Sub
