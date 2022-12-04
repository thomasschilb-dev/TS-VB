Attribute VB_Name = "modIRC"
Option Explicit

Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(32) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(32) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

Public QueueMsg As New Collection
Public RealName As String
Public Email As String
Dim strMyNick As String


Public Sub Connect(strIP As String, strPort As String)
    With frmSocket.Socket
        .Close
        
        While .State
            DoEvents
        Wend
        
        .Connect strIP, strPort
    End With

End Sub

Public Sub Disconnect()
    If frmSocket.Socket.State = sckConnected Then
        SendData "QUIT"
    End If
    Timeout 0.5
    frmSocket.Socket.Close
    LogText frmStatus.rtfStatus, "*** Disconnected", udtColor.colorQuit
End Sub

Public Property Get MyNick() As String
    MyNick = strMyNick
End Property

Public Property Let MyNick(ByVal strTemp As String)
    strMyNick = strTemp
End Property

Public Sub ParseMsg(strHost As String, strTrigger As String, strMiddle As String, strMsg As String)

    Dim strChannel As String
    Dim strTemp As String
    Dim strMode As String
    Static strTopic As String
    Dim intTemp As Integer
    Dim strText As String
    Dim strAction As String
    Dim strTime As String
    Dim strNick As String
    
    Dim strDummy As String
    Dim strDummy1 As String
    
    Dim lngCounter As Long
    Dim i As Integer
    Dim Email


    
    If Left(strMsg, 1) = ":" Then
        strMsg = Mid(strMsg, 2)
    End If
    
    If Left(strHost, 1) = ":" Then
                                       
        strHost = Mid(strHost, 2)
                                      
    End If
    
    If Left(strMiddle, 1) = ":" Then
        strMiddle = Mid(strMiddle, 2)
    End If
                
    For i = 1 To Len(MyNick)
        If Mid(MyNick, i, 1) = "!" Then
            Email = Mid(MyNick, i + 1)
            MyNick = Mid(MyNick, 1, i - 1)
            If Left(MyNick, 1) = ":" Then
                MyNick = Mid(MyNick, 2)
            End If
        End If
    Next i
                
    Select Case UCase(strTrigger)
        
        Case "NOTICE"
            'store the hostname when first connect only certain server have this msg
            If UCase(strMiddle) = "AUTH" Then strIrcServer = strHost
            
            'Get the nick from the host string
            If InStr(strHost, "!") Then
                strNick = Trim(Left(strHost, InStr(strHost, "!") - 1))
            Else
                strNick = strHost
            End If
            
            LogToAll strNick, strMsg, 1
            LogText frmStatus.rtfStatus, "-" & strNick & "- " & strMsg, udtColor.colorNotice
        
        Case "PING"
            'Ping to server to keep connection alive
            SendData "PONG :" & strHost
            LogText frmStatus.rtfStatus, "PING? PONG!", udtColor.colorOther
        Case "MODE"

            If InStr(strMsg, " ") Then
                'check to see if mode set by server or users
                strNick = Trim(Left(strHost, InStr(strHost, "!") - 1))
                strChannel = LCase(strMiddle)
                strMode = Trim(Left(strMsg, InStr(strMsg, " ") - 1))
                strText = Trim(Right(strMsg, Len(strMsg) - InStr(strMsg, " ")))
                
                SetNickMode LCase(strChannel), strText, strMode
                LogTextToHwnd LCase(strChannel), "*** " & strNick & " sets mode: " & strMode & " " & strText, Channel_hWnd, udtColor.colorMode, False
            Else
                'channel mode has been set
                If LCase(strMiddle) <> LCase(MyNick) Then
                    strChannel = strMiddle
                    strNick = Trim(Left(strHost, InStr(strHost, "!") - 1))
                    strMode = strMsg
                    SaveMode LCase(strChannel), strMode 'save the modes to the mode list
                    
                    'store the topic of the channel
                    strText = GetChannel(GetCaption(LCase(strChannel)))
                    If Len(strText) <> 0 Then
                        ChangeCaption LCase(strChannel), strChannel & GetMode(LCase(strChannel)) & ": " & GetChannel(GetCaption(LCase(strChannel)))
                    Else
                        If Len(GetMode(LCase(strChannel))) = 0 Then
                            ChangeCaption LCase(strChannel), strChannel
                        Else
                            ChangeCaption LCase(strChannel), strChannel & GetMode(LCase(strChannel))
                        End If
                    End If
                    
                    LogTextToHwnd LCase(strChannel), strNick & " sets mode: " & strMsg, Channel_hWnd, udtColor.colorMode, False
                Else
                    'mode set by server
                    frmStatus.Caption = "Status: " & MyNick & "[" & strMsg & "] on " & strIrcServer
                    LogText frmStatus.rtfStatus, "*** " & strHost & " sets mode: " & strMsg, udtColor.colorMode
                End If
            End If
        Case "JOIN"
            
            If InStr(strHost, "!") Then
                'get nickname
                strNick = Trim(Left(strHost, InStr(strHost, "!") - 1))
                'You join the channel
                If LCase(strNick) = LCase(MyNick) Then
                    'You join the channel
                    strChannel = Mid(strMiddle, 1)
                    'create a channel's hWnd
                    CreateHwnd LCase(strChannel), Channel_hWnd
                    'change caption to channel
                    ChangeCaption LCase(strChannel), strChannel
                    'request for channel's mode
                    SendData "MODE " & strChannel
                    LogTextToHwnd LCase(strChannel), "*** Now talking in " & strChannel, Channel_hWnd, udtColor.colorJoin, False
                Else
                    'Other people join the channel
                    strChannel = Mid(strMiddle, 1)
                    LogTextToHwnd LCase(strChannel), "*** " & strNick & " has joined " & strChannel, Channel_hWnd, udtColor.colorJoin, False
                    AddToList strChannel, strNick
                End If
            End If
        Case "PART"
            'get nick name
            strNick = Trim(Left(strHost, InStr(strHost, "!") - 1))
            strChannel = strMiddle
            
            If LCase(strNick) <> LCase(MyNick) Then
                LogTextToHwnd LCase(strChannel), "*** " & strNick & " has left " & strMiddle, Channel_hWnd, udtColor.colorPart, False
                RemoveName LCase(strChannel), strNick
            End If
            
        Case "NICK"
            strNick = Trim(Left(strHost, InStr(strHost, "!") - 1))
            LogToAll strNick, strMiddle, 2
        Case "KICK"

            strChannel = strMiddle
            strNick = Trim(Left(strHost, InStr(strHost, "!") - 1))
            strTemp = Trim(Left(strMsg, InStr(strMsg, ":") - 1))
            strText = Trim(Right(strMsg, Len(strMsg) - InStr(strMsg, ":")))
            
            If IsOp(LCase(strChannel), strTemp) Then
                strDummy = "@" & strTemp
                strDummy1 = "@" & strTemp
            ElseIf IsVoice(LCase(strChannel), strTemp) Then
                strDummy = "+" & strTemp
                strDummy1 = "+" & strTemp
            Else
                strDummy = strTemp
                strDummy1 = strTemp
            End If
            
            If LCase(strTemp) = LCase(MyNick) Then
                CloseWindow LCase(strChannel)
                LogText frmStatus.rtfStatus, "You were kicked by " & strNick & " (" & strMsg & ")", udtColor.colorKick
            Else
                RemoveName LCase(strChannel), strNick
                LogTextToHwnd LCase(strChannel), strTemp & " was kicked by " & strNick & " (" & strText & ")", Channel_hWnd, udtColor.colorKick, False
            End If
            
        Case "QUIT"
            'get nick name
            strNick = Trim(Left(strHost, InStr(strHost, "!") - 1))

            'remove name and show error message
            If Len(strMsg) = 0 Then
                strMiddle = strMiddle
            Else
                strMiddle = strMiddle & " "
            End If
            LogToAll strNick, strMiddle & strMsg, 3

        Case "PRIVMSG"
            'get nickname
            strNick = Trim(Left(strHost, InStr(strHost, "!") - 1))
            
            If Left(strMiddle, 1) = "#" Then
            'channel's private message
                strChannel = strMiddle
                strText = Trim(Right(strMsg, Len(strMsg)))
                If InStr(strText, Chr(1) & "ACTION") Then
                    intTemp = InStr(1, strText, " ")
                    strAction = Mid(strText, intTemp, Len(strText) - intTemp)
                    LogTextToHwnd LCase(strChannel), "*" & strNick & strAction, Channel_hWnd, udtColor.colorAction, False
                Else
                    LogTextToHwnd LCase(strChannel), "<" & strNick & "> " & strText, Channel_hWnd, udtColor.colorChat, True
                End If
            Else
            'private message from other nick
            'the tag will be the nick
                strNick = Trim(Left(strHost, InStr(strHost, "!") - 1))
                'CreateHwnd LCase(strNick), Private_hWnd
                LogTextToHwnd LCase(strNick), "<" & strNick & "> " & strMsg, Private_hWnd, udtColor.colorChat, True
                ChangeCaption LCase(strNick), strNick
                'LogTextToHwnd LCase(strNick), "<" & strNick & "> " & strMsg, Private_hWnd, udtColor.colorChat, True
            End If
                        
        Case "001"
            'Get the hostname when first connect
            strIrcServer = strHost
            If Len(strMsg) <> 0 Then
                LogText frmStatus.rtfStatus, strMsg, udtColor.colorOther
            End If
            
        Case "324"
        'Channel's mode will be store in a listbox
        '#adx +stnk anime
            strChannel = Trim(Left(strMsg, InStr(strMsg, " ") - 1))
            intTemp = InStr(1, strMsg, " ")
            strMode = Trim(Mid(strMsg, intTemp + 1, InStr(intTemp, strMsg, " ")))
            'Only save the mode if there are some modes
            If Len(strMode) > 1 Then
                SaveMode LCase(strChannel), strMode
            Else
                strMode = ""
            End If
            
            'Check to see if there is a topic
            If Len(strTopic) <> 0 Or Trim(strTopic) <> "" Then
                ChangeCaption LCase(strChannel), strChannel & "[" & strMode & "]: " & strTopic
            Else
                If Len(strMode) <> 0 Then
                    ChangeCaption LCase(strChannel), strChannel & strMode & ": "
                Else
                    ChangeCaption LCase(strChannel), strChannel
                End If
            End If
            strTopic = ""
            
        Case "332"
        'TOPIC
            strChannel = Trim(Left(strMsg, InStr(strMsg, " ") - 1))
            strTopic = Trim(Right(strMsg, Len(strMsg) - InStr(strMsg, ":")))
            LogTextToHwnd LCase(strChannel), "*** Topic is '" & strTopic & "'", Channel_hWnd, udtColor.colorTopic, False
        Case "333"
        'SET BY
            'Parse the Channel, Nick and Time
            intTemp = InStr(1, strMsg, " ")
            strChannel = LCase(Trim(Left(strMsg, InStr(strMsg, " ") - 1)))
            strNick = Trim(Mid(strMsg, intTemp + 1, InStr(intTemp + 1, strMsg, " ") - (intTemp + 1)))
            strTime = Trim(Right(strMsg, Len(strMsg) - InStr(intTemp + 1, strMsg, " ")))
            LogTextToHwnd strChannel, "*** Set by " & strNick & " on " & AscTime(strTime), Channel_hWnd, udtColor.colorMode, False
            
        Case "329"
        'CREATED AT
            strChannel = LCase(Trim(Left(strMsg, InStr(strMsg, " ") - 1)))
            strTime = Trim(Right(strMsg, Len(strMsg) - InStr(strMsg, " ")))
            LogTextToHwnd strChannel, "*** " & strChannel & " created on " & AscTime(strTime), Channel_hWnd, udtColor.colorMode, False
        Case "353"
        'PEOPLE
            'get the channel's name
            strChannel = Trim(Mid(strMsg, 3, InStr(strMsg, ":") - 4))
            strTemp = Right(strMsg, Len(strMsg) - InStr(strMsg, ":"))
            GetList LCase(strChannel), strTemp
            LogText frmStatus.rtfStatus, strChannel & " " & strTemp, udtColor.colorOther
        Case Else
        'Other messages need to parse, just temporary display it
            If Len(strMsg) <> 0 Then
                LogText frmStatus.rtfStatus, strMsg, udtColor.colorAction
            End If
    End Select
End Sub
'Check to see if the data is one complete data or broken data
Public Sub CheckForLine(strLine As String)
    Dim intCounter As Integer
    Dim strOneLine As String
    Static strRestLine As String
    
    strLine = strRestLine & strLine
    For intCounter = 1 To Len(strLine)
        If Mid(strLine, intCounter, 1) = Chr(13) Or Mid(strLine, intCounter, 1) = Chr(10) Then
            strOneLine = Mid(strLine, 1, intCounter - 1)
            'if the data is one complete line, then add to the queue
            QueueMsg.Add strOneLine
            strLine = Mid(strLine, intCounter + 1)
            intCounter = 1
        End If
    Next
    strRestLine = strLine
End Sub
'Send data w/ winsock control in frmSocket
Public Sub SendData(strData As String)
    With frmSocket.Socket
        If .State = sckConnected Then
            .SendData strData & vbCrLf
        End If
    End With
End Sub
'Return the time calculate from given time to January 1 1970
'This is UNIX time system
Public Function AscTime(TheTime As String) As String
    Dim lpTime As TIME_ZONE_INFORMATION
    Dim TheDate As Date
    Dim Msg
    
    On Error GoTo err_handle
    
    TheDate = "January 1 1970 00:00:00"
    GetTimeZoneInformation lpTime
    
    If IsNumeric(TheTime) Then
        'given timer
        AscTime = Format(DateAdd("s", Val(TheTime) - (lpTime.Bias * 60), TheDate), "ddd mmm dd yyyy hh:mm:ss")
    Else
        'given date
        AscTime = (DateDiff("s", TheDate, TheTime) - (lpTime.Bias * 60))
    End If
    Exit Function
err_handle:
    AscTime = "Invalid date/time format"
End Function
'Remove certain character from a string
Public Function TrimChar(strMsg As String, strChar As String)
    Dim lngCounter As Long
    Dim strTemp As String
    
    For lngCounter = 1 To Len(strMsg)
        If Mid(strMsg, lngCounter, 1) <> strChar Then
            strTemp = strTemp & Mid(strMsg, lngCounter, 1)
        Else
            strTemp = strTemp
        End If
    Next
    TrimChar = strTemp
End Function

