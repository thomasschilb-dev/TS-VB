Option Strict Off
Option Explicit On
Module tsMIDIMod
	Public Declare Function mciSendString Lib "winmm.dll"  Alias "mciSendStringA"(ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Integer, ByVal hwndCallback As Integer) As Integer
	Public Declare Function mciGetErrorString Lib "winmm.dll"  Alias "mciGetErrorStringA"(ByVal fdwError As Integer, ByVal lpszErrorText As String, ByVal cchErrorText As Integer) As Integer
	
	Public Sub OpenMID(ByRef MIDFile As String, ByRef MIDAlias As String)
		Dim ErrCode As Integer
		ErrCode = mciSendString("open " & MIDFile & " alias " & MIDAlias, "", 0, 0)
		If ErrCode <> 0 Then DisplayError(ErrCode)
	End Sub
	
	Public Sub PlayMID(ByRef MIDAlias As String)
		Dim ErrCode As Integer
		
		ErrCode = mciSendString("play " & MIDAlias, "", 0, 0)
		If ErrCode <> 0 Then DisplayError(ErrCode)
	End Sub
	
	Public Sub StopMID(ByRef MIDAlias As String)
		Dim ErrCode As Integer
		
		ErrCode = mciSendString("stop " & MIDAlias, "", 0, 0)
		If ErrCode <> 0 Then DisplayError(ErrCode)
	End Sub
	
	Public Sub CloseMID(ByRef MIDAlias As String)
		Dim ErrCode As Integer
		ErrCode = mciSendString("close " & MIDAlias, "", 0, 0)
	End Sub
	
	Private Sub DisplayError(ByVal ErrCode As Integer)
		Dim errstr As String
		Dim retval As Integer
		
		errstr = Space(128)
		retval = mciGetErrorString(ErrCode, errstr, Len(errstr))
		errstr = Left(errstr, InStr(errstr, vbNullChar) - 1)
		
		retval = MsgBox(errstr, MsgBoxStyle.OKOnly Or MsgBoxStyle.Critical, "MCI")
	End Sub
End Module