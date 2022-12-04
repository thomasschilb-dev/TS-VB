Attribute VB_Name = "modMain"
'INI Functions
Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
    
'Variable for Channel
Public ChannelArray() As New frmChannel
Public ChannelArrayFree() As Boolean
'Variable for Private
Public PrivateArray() As New frmPrivate
Public PrivateArrayFree() As Boolean

'Variables store servers
Public CRLF As String
Public strIrcServer As String

Public Enum hWndType
    Channel_hWnd = 1
    Private_hWnd = 2
End Enum
'Temporary Variables
Public strEmail As String
Public strTheName As String
    


Public Sub Main()
    'Initialize form arrays
    ReDim Preserve PrivateArray(0)
    ReDim Preserve PrivateArrayFree(0)
    
    ReDim Preserve ChannelArray(0)
    ReDim Preserve ChannelArrayFree(0)
    
    'frmChannel.Show
    CRLF = Chr(13) & Chr(10)
    Load frmMain
    frmMain.Show
    frmStatus.Show
    Load frmSocket
End Sub

Public Function ReadINI(Header As String, Key As String, Default As String, Location As String) As String
    Dim strData As String
    Dim lngLength As Long
    
    strData = Space(255)
    lngLength = GetPrivateProfileString(Header, Key, Default, strData, 255, Location)
    ReadINI = Left(strData, lngLength)
End Function

Public Sub WriteINI(Header As String, Key As String, Value As String, Location As String)
    Call WritePrivateProfileString(Header, Key, Value, Location)
End Sub

Public Sub GetServers()
    Dim strLoc As String
    Dim intCount As Integer
    Dim strResult As String
    strLoc = "servers.ini"
    intCount = 0
    Do
        strResult = ReadINI("servers", "n" & intCount, "", strLoc)
        If Len(strResult) = 0 Or Trim(strResult) = "" Then Exit Do
        'Austnet: Random AU serverSERVER:au.austnet.org:6667GROUP:Austnet
        
        frmOption.cmbServers.AddItem strResult
        intCount = intCount + 1
    Loop
End Sub

'pause few seconds before next action
Public Sub Timeout(intDuration As Integer)
Dim lngStart As Long
Dim intNothing As Integer
    
    lngStart = Timer

    Do While Timer - lngStart < intDuration
        intNothing = DoEvents()
    Loop

End Sub
