Attribute VB_Name = "modVar"
Type Admin
  Name As String
  Email As String
  Location As String
End Type

Type Channel
  Title As String
  Topic As String
  TopicTime As Long
  TopicUser As String
  Modes As String
  Bans(0 To 1000) As String
  UMode(0 To 1000) As String
  Created As Long
  Key As String
  ULimit As Integer
  IndexDB As Integer
End Type

Type Client
  Nick As String
  User As String
  Host As String
  Modes As String
  State As Integer
  LastPing As Long
  GotPong As Long
  Chans(1 To 1000) As Integer
  Invite(1 To 1000) As Integer
  RealName As String
  ConnectTime As Long
  LastAction As Long
  Pass As String
  NickTimer As Long
  NickMatch As Integer
  IsAway As Byte
  AwayMsg As String
End Type

Type Oper
  Name As String
  Host As String
  Password As String
  Modes As String
End Type

Type SvcServer
  Host As String
  Name As String
  Password As String
  Desc As String
End Type

Type RemoteType
  Host As String
  User As String
  Password As String
End Type

Public SQct As Integer
Public Services As SvcServer
Public SvcSock As Integer
Public Creation As String
Public WorkDir As String
Public Sysop As Admin
Public ServerHost As String
Public ServerName As String
Public ServType As Integer
Public Remote As RemoteType
Public Clients(0 To 1000) As Client
Public Chans(0 To 1000) As Channel
Public Opers(1 To 1000) As Oper
Public Buffer(0 To 1000) As String
Public InBuffer(0 To 1000) As String
Public Reserved(1 To 100) As String
Public lPort As Integer
Public MOTD As String
Public KillFlag(0 To 1000) As Byte
Public Temp(0 To 1000) As Byte
Public OperCt As Integer
Public Hash As New MD5
Public NickLen As Byte
Public ChanLen As Byte
Public TopicLen As Integer
Public PingInterval As Integer
Public PingTimeout As Integer

Public Const NickChars As String = "0123456789qwertyuiop[]|\asdfghjkl;'zxcvbnm./^-_"
Public Const ChanChars As String = "#0123456789qwertyuiop[]|\asdfghjkl;'zxcvbnm./^-_"
Public Const Build As String = "tsIRCd 1.1"

Public Sub InitServer()
Creation = Date$ + " " + Time$
WorkDir = ""
NickLen = 32
ChanLen = 32
TopicLen = 400
PingInterval = 120
PingTimeout = 60
NS.AutoProtect = 1
NS.GuestPrefix = "Guest"

RehashMOTD
RehashConfig
For n = frmMain.Sock.lbound To frmMain.Sock.UBound
  frmMain.Sock(n).Close
Next n
curunix& = UnixTimeEnc&(Date$, Timer)
For n = frmMain.sckListen.lbound To frmMain.sckListen.UBound
  frmMain.sckListen(n).Close
  frmMain.sckListen(n).LocalPort = lPort
  'On Error Resume Next
  frmMain.sckListen(n).Listen
  Clients(n).LastPing = curunix&
Next n

InitServices
frmMain.PingTimer.Interval = 1000
frmMain.PingTimer.Enabled = True
End Sub

Public Sub RehashConfig()
Dim args(10) As String
nf = FreeFile
Open WorkDir + "ircd.conf" For Input As nf
For n = 1 To 1000
  Opers(n).Host = ""
  Opers(n).Modes = ""
  Opers(n).Name = ""
  Opers(n).Password = ""
Next n
OperCt = 1
Do Until EOF(nf)
  Line Input #nf, tmp$
  tmp$ = LTrim$(tmp$)
  If Left$(tmp$, 1) <> "#" Then
    carg = 0
    Do Until Len(tmp$) = 0
      xp = InStr(1, tmp$, ":")
      If xp < 1 Then xp = Len(tmp$) + 1
      args(carg) = Left$(tmp$, xp - 1)
      tmp$ = Mid$(tmp$, xp + 1)
      carg = carg + 1
    Loop
      Select Case LCase$(args(0))
        Case "d"
          ServerName = args(1)
        Case "m"
          ServerHost = args(1)
        Case "a"
          Sysop.Name = args(1)
          Sysop.Email = args(2)
        Case "r"
          Remote.Host = args(1)
          Remote.Password = args(2)
          Remote.User = args(3)
        Case "o"
          If Len(args(1)) > 0 Then
            Opers(OperCt).Host = args(1)
            Opers(OperCt).Password = args(2)
            Opers(OperCt).Name = args(3)
            Opers(OperCt).Modes = args(5)
            OperCt = OperCt + 1
          End If
        Case "p"
          lPort = Val(args(4))
        Case "s"
          Select Case LCase$(args(1))
            Case "none"
              ServType = 0
            Case "internal"
              ServType = 1
            Case Else
              ServType = 2
              Services.Host = args(1)
              Services.Password = args(2)
              Services.Name = args(3)
          End Select
        Case "ns"
          Select Case LCase$(args(1))
            Case "autoprotect"
              NS.AutoProtect = Val(args(2))
            Case "guestprefix"
              NS.GuestPrefix = args(2)
          End Select
      End Select
      'carg = carg + 1
    'Loop
  End If
Loop
Close nf
frmMain.MaintTimer.Enabled = True
End Sub

Public Sub SM(Recipient As Integer, Message As String)
'If Recipient = SvcSock And frmMain.chkDebug.value = vbChecked Then frmMain.txtDebug = frmMain.txtDebug + "---> " + Message + vbCrLf: frmMain.txtDebug.SelStart = Len(frmMain.txtDebug) + 1

If Len(LTrim$(RTrim$(Message))) = 0 Then Exit Sub

origsize = Len(Buffer(Recipient))
Buffer(Recipient) = Buffer(Recipient) + Message + Chr$(13)
xp = InStr(1, Buffer(Recipient), Chr$(13))
If xp < 1 Then xp = Len(Buffer(Recipient)) + 1
a$ = Left$(Buffer(Recipient), xp - 1)
Buffer(Recipient) = Mid$(Buffer(Recipient), xp + 1)
If Recipient > frmMain.Sock.UBound Then Exit Sub
If origsize = 0 Then 'And frmMain.Sock(Recipient).State = sckConnected Then
  On Error Resume Next: frmMain.Sock(Recipient).SendData a$ + vbCrLf: On Error GoTo 0
  If frmMain.Visible = True Then
    frmMain.Text1 = frmMain.Text1 + "--->" + a$ + vbCrLf
    frmMain.Text1.SelStart = Len(frmMain.Text1) + 1
  End If
End If
End Sub

Public Sub ClearClientSlot(SockVal As Integer)
frmMain.Sock(SockVal).Close
Clients(SockVal).Host = ""
Clients(SockVal).Modes = ""
Clients(SockVal).Nick = ""
Clients(SockVal).State = 0
Clients(SockVal).User = ""
Clients(SockVal).LastPing = 0
Clients(SockVal).GotPong = 0
Clients(SockVal).RealName = ""
Clients(SockVal).LastAction = 0
Clients(SockVal).ConnectTime = 0
Clients(SockVal).Pass = ""
Clients(SockVal).NickTimer = 0
Clients(SockVal).IsAway = 0
Clients(SockVal).AwayMsg = ""
For n = 1 To 1000
  If Clients(SockVal).Chans(n) = 1 Then Chans(n).UMode(SockVal) = ""
  Clients(SockVal).Chans(n) = 0
  Clients(SockVal).Invite(n) = 0
Next n
Buffer(SockVal) = ""
KillFlag(SockVal) = 0
End Sub

Public Function ReplaceAll$(NewString As String, LookFor As String, ReplaceWith As String)
Dim NS As String
NS = NewString

Do
  xp = InStr(1, NS, LookFor)
  If xp < 1 Then Exit Do
  NS = Left$(NS, xp - 1) + ReplaceWith + Mid$(NS, xp + Len(ReplaceWith))
Loop Until Not InStr(1, NS, LookFor)
ReplaceAll$ = NS
End Function

Public Sub ProcessInput(FromSock As Integer, InData As String)
If Len(InData) = 0 Then Exit Sub

Dim newchan As Integer
Dim args(20) As String

tmpd$ = LTrim$(InData)
Select Case UCase$(Left$(tmpd$, 3))
  Case "NS "
    tmpd$ = "PRIVMSG NickServ " + Mid$(tmpd$, 4)
  Case "CS "
    tmpd$ = "PRIVMSG ChanServ " + Mid$(tmpd$, 4)
  Case "OS "
    tmpd$ = "PRIVMSG OperServ " + Mid$(tmpd$, 4)
  Case "HS "
    tmpd$ = "PRIVMSG HostServ " + Mid$(tmpd$, 4)
  Case "MS "
    tmpd$ = "PRIVMSG MemoServ " + Mid$(tmpd$, 4)
End Select

Select Case UCase$(Left$(tmpd$, 9))
  Case "NICKSERV "
    tmpd$ = "PRIVMSG NickServ " + Mid$(tmpd$, 10)
  Case "CHANSERV "
    tmpd$ = "PRIVMSG ChanServ " + Mid$(tmpd$, 10)
  Case "OPERSERV "
    tmpd$ = "PRIVMSG OperServ " + Mid$(tmpd$, 10)
  Case "HOSTSERV "
    tmpd$ = "PRIVMSG HostServ " + Mid$(tmpd$, 10)
  Case "MEMOSERV "
    tmpd$ = "PRIVMSG MemoServ " + Mid$(tmpd$, 10)
End Select

'tmpd$ = ReplaceAll$(tmpd$, Chr$(13), "")
'tmpd$ = ReplaceAll$(tmpd$, Chr$(10), "")
fullline$ = tmpd$
If frmMain.Visible = True Then
  frmMain.Text1 = frmMain.Text1 + "<---" + fullline$ + vbCrLf
  frmMain.Text1.SelStart = Len(frmMain.Text1) + 1
End If
xp = InStr(1, tmpd$, " ")
If xp < 1 Then xp = Len(tmpd$) + 1
args(0) = UCase$(Left$(tmpd$, xp - 1))
tmpd$ = LTrim$(Mid$(tmpd$, xp + 1))
origline$ = tmpd$

carg = 1
Do
  xp = InStr(1, tmpd$, " ")
  If xp < 1 Then args(carg) = tmpd$: Exit Do
  args(carg) = Left$(tmpd$, xp - 1)
  tmpd$ = LTrim$(Mid$(tmpd$, xp + 1))
  carg = carg + 1
Loop Until carg = 21

If UCase$(args(1)) = "PONG" Then args(0) = args(1): args(1) = args(2)
If UCase$(args(0)) <> "PONG" Then Clients(FromSock).LastAction = UnixTimeEnc&(Date$, Timer)

'If Clients(FromSock).State = 2 Then
'
'End If

If Clients(FromSock).State = 0 Then
  Select Case args(0)
    Case "NICK"
      args(1) = NormalizeNick(args(1))
     If NickInUse(args(1)) Then
       SendError FromSock, 433, args(1) + " :Nickname is already in use"
     Else
       If Len(args(1)) > 0 Then
         Clients(FromSock).Nick = args(1)
         Clients(FromSock).NickMatch = 0
         If IsUserMode(FromSock, "r") Then UnSetUserMode FromSock, ServerHost, "r"
       Else
         SendError FromSock, 431, ":No nickname given"
       End If
     End If
     
    Case "USER"
      If Len(args(1)) = 0 Then
        SendError FromSock, 461, args(0) + " :Not enough parameters"
      Else
        Clients(FromSock).User = args(1)
        xp = InStr(1, origline$, ":")
        If xp > 0 Then Clients(FromSock).RealName = Mid$(origline$, xp + 1) Else Clients(FromSock).RealName = "Nobody"
      End If

    Case "PASS"
      If Len(args(1)) > 0 Then
        If Left$(args(1), 1) = ":" Then args(1) = Mid$(args(1), 2)
        Clients(FromSock).Pass = Hash.DigestStrToHexStr(args(1))
      Else
        SendError FromSock, 461, args(0) + " :Not enough parameters"
      End If
      
    Case "REMOTE"
      If LCase$(args(2)) = Hash.DigestStrToHexStr(Clients(FromSock).Pass) Then
      End If
      
    'Case "SERIVCES"
      'if lcase$(args(1)) = lcase$(services.host) then
      
    Case "PONG"
      Clients(FromSock).GotPong = 1
    Case Else
      If Len(args(0)) > 0 Then SendError FromSock, 451, ":You have not registered " + args(0) + " " + origline$
  End Select
  If Len(Clients(FromSock).Nick) > 0 And Len(Clients(FromSock).User) > 0 Then SendGreeting FromSock: Clients(FromSock).State = 1
  Exit Sub
End If

Select Case args(0)
  Case "MOTD"
    SendMOTD FromSock

  Case "USER"
    SendError FromSock, 462, ":You may not reregister"
    
  Case "WHO"
    If Len(args(1)) = 0 Then SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
    args(1) = LTrim$(RTrim$(args(1)))
    args(2) = LTrim$(RTrim$(args(2)))
    If LCase$(args(2)) = "o" Then filteroper = 1 Else filteroper = 0
    If Left$(args(1), 1) = "#" Then
      newchan = ChanExists(args(1))
      If newchan > 0 Then
        For n = 0 To frmMain.Sock.UBound
          If n > frmMain.Sock.UBound Then Exit For
          If Clients(n).Chans(newchan) = 1 Then
            If Clients(n).IsAway = 0 Then prefix$ = "H" Else prefix$ = "G"
            If IsUserMode(Int(n), "ao") Then prefix$ = prefix$ + "*"
            If IsUserChanMode(Int(n), newchan, "v") Then prefix$ = prefix$ + "+"
            If IsUserChanMode(Int(n), newchan, "o") Then prefix$ = prefix$ + "@"
            If IsUserChanMode(Int(n), newchan, "q") Then prefix$ = prefix$ + "~"
            If filteropers = 0 Then
              SM FromSock, ":" + ServerHost + " 352 " + Clients(FromSock).Nick + " " + args(1) + " " + Clients(n).User + " " + Clients(n).Host + " " + ServerHost + " " + Clients(n).Nick + " " + prefix$ + " :0 " + Clients(n).RealName
            Else
              If IsUserMode(Int(n), "ao") Then SM FromSock, ":" + ServerHost + " 352 " + Clients(FromSock).Nick + " " + args(1) + " " + Clients(n).User + " " + Clients(n).Host + " " + ServerHost + " " + Clients(n).Nick + " " + prefix$ + " :0 " + Clients(n).RealName
            End If
          End If
        Next n
      Else
        For n = 0 To frmMain.Sock.UBound
          If n > frmMain.Sock.UBound Then Exit For
          If Len(Clients(n).Nick) > 0 Then
            If Clients(n).IsAway = 0 Then prefix$ = "H" Else prefix$ = "G"
            If IsUserMode(Int(n), "ao") Then prefix$ = prefix$ + "*"
            If filteropers = 0 Then
              If IsUserMode(Int(n), "i") = 0 Then SM FromSock, ":" + ServerHost + " 352 " + Clients(FromSock).Nick + " * " + Clients(n).User + " " + Clients(n).Host + " " + ServerHost + " " + Clients(n).Nick + " " + prefix$ + " :0 " + Clients(n).RealName
            Else
              If IsUserMode(Int(n), "ao") <> 0 And IsUserMode(Int(n), "i") = 0 Then SM FromSock, ":" + ServerHost + " 352 " + Clients(FromSock).Nick + " * " + Clients(n).User + " " + Clients(n).Host + " " + ServerHost + " " + Clients(n).Nick + " " + prefix$ + " :0 " + Clients(n).RealName
            End If
          End If
        Next n
      End If
    End If
    SM FromSock, ":" + ServerHost + " 315 " + Clients(FromSock).Nick + " " + args(1) + " :End of /WHO list."
    
  Case "PART"
    If Len(args(1)) = 0 Then SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
    xp = InStr(1, origline$, " ")
    If xp > 0 Then
      partreason$ = Mid$(origline$, xp + 1)
      origline$ = Left$(origline$, xp - 1)
    Else
      partreason$ = ""
    End If
    Do Until Len(origline$) = 0
      xp = InStr(1, origline$, ",")
      If xp < 1 Then xp = Len(origline$) + 1
      cc$ = Left$(origline$, xp - 1)
      origline$ = Mid$(origline$, xp + 1)
      cc$ = NormalizeChan$(cc$)
      If Len(cc$) < 2 Then
        SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
      Else
        newchan = ChanExists(cc$)
        If newchan > 0 Then
          If Clients(FromSock).Chans(newchan) = 1 Then
            PartChan FromSock, newchan, partreason$
          Else
            SendError FromSock, 442, cc$ + " :You're not on that channel"
          End If
        Else
          SendError FromSock, 403, cc$ + " :No such channel"
        End If
      End If
    Loop

  Case "JOIN"
    If Len(args(1)) = 0 Then SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
    Do Until Len(origline$) = 0
      xp = InStr(1, origline$, ",")
      If xp < 1 Then xp = Len(origline$) + 1
      cc$ = Left$(origline$, xp - 1)
      origline$ = Mid$(origline$, xp + 1)
      xp = InStr(1, cc$, " ")
      If xp < 1 Then ckey$ = "" Else ckey$ = Mid$(cc$, xp + 1): cc$ = Left$(cc$, xp - 1)
      cc$ = NormalizeChan$(cc$)
      If Len(cc$) < 2 Then
        SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
      Else
        If Left$(cc$, 1) = "#" Then
          newchan = JoinChan(FromSock, cc$, ckey$)
        Else
          SendError FromSock, 401, cc$ + " :No such nick/channel"
        End If
      End If
    Loop
    'If Chans(newchan).IndexDB > 0 Then ChanMaint newchan

  Case "OPER"
    For n = 1 To OperCt
      If LCase$(args(1)) = LCase$(Opers(n).Name) And Hash.DigestStrToHexStr(args(2)) = Opers(n).Password And MatchMask%(Clients(FromSock).Host, Opers(n).Host) = 1 Then
        SetUserMode FromSock, ServerHost, Opers(n).Modes
        SM FromSock, ":" + ServerHost + " 381 " + Clients(FromSock).Nick + " :You are now an e-god. Play nicely with the other children."
        SM FromSock, ":" + ServerHost + " NOTICE " + Clients(FromSock).Nick + " :*** Oper priveleges are " + Opers(n).Modes
        Exit Sub
      End If
    Next n
    SM FromSock, ":" + ServerHost + " 491 " + Clients(FromSock).Nick + " :There is no O-line for your host."
    
  Case "PRIVMSG", "NOTICE"
    If IsRsvd(args(1)) = 1 Then
      HandleServiceMsg FromSock, args(1), Mid$(origline$, Len(args(1)) + 1)
      Exit Sub
    End If
    
    'SM SvcSock, ":" + Clients(FromSock).Nick + " " + args(0) + " " + args(1) + " " + Mid$(origline$, Len(args(1)) + 1): Exit Sub
    
    If Len(args(1)) = 0 Then SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
    If Len(args(2)) = 0 Then SendError FromSock, 412, ":No text to send": Exit Sub
    cansend = 1
    If Left$(args(1), 1) = "#" Then
      newchan = ChanExists(args(1))
      If newchan = 0 Then
        SendError FromSock, 401, cc$ + " :No such nick/channel"
      Else
        If IsChanMode(newchan, "n") = 1 And Clients(FromSock).Chans(newchan) = 0 Then
          SendError FromSock, 404, Chans(newchan).Title + " :Cannot send to channel (+n)"
          cansend = 0
        Else
          If IsUserChanMode(FromSock, newchan, "qov") = 0 And IsChanMode(newchan, "m") = 1 Then
        SendError FromSock, 404, Chans(newchan).Title + " :Cannot send to channel (+m)"
            cansend = 0
          End If
        End If
      End If
      For n = 0 To 1000
        If Len(Chans(newchan).Bans(n)) > 0 Then
          If MatchMask%(FullHost(FromSock), Chans(newchan).Bans(n)) = 1 And IsUserChanMode(FromSock, newchan, "qo") = 0 Then
            If IsUserMode(FromSock, "ao") Then cansend = 1: Exit For
            SendError FromSock, 404, Chans(newchan).Title + " :Cannot send to channel (+b)"
            cansend = 0
            Exit For
          End If
        End If
      Next n
      If cansend = 1 Then BroadcastChan newchan, ":" + FullHost(FromSock) + " " + args(0) + " " + Chans(newchan).Title + Mid$(origline$, Len(args(1)) + 1), FromSock
    Else
      tmpub = frmMain.Sock.UBound
      For n = 0 To tmpub
        If LCase$(Clients(n).Nick) = LCase$(args(1)) Then
          SM Int(n), ":" + FullHost(FromSock) + " " + args(0) + " " + Clients(n).Nick + Mid$(origline$, Len(args(1)) + 1)
          If Clients(n).IsAway = 1 Then SM Int(n), ":" + ServerHost + " 301 " + Clients(FromSock).Nick + " " + Clients(n).Nick + " " + Clients(n).AwayMsg
          Exit For
        End If
      Next n
      If n > tmpub Then SendError FromSock, 401, args(1) + " :No such nick/channel"
    End If
    
  Case "NICK"
    args(1) = NormalizeNick(args(1))
    If NickInUse(args(1)) Then
      SendError FromSock, 433, args(1) + " :Nickname is already in use"
    Else
      If Len(args(1)) > 0 Then
        Clients(FromSock).NickMatch = 0
        ChangeNick FromSock, args(1)
        If ServType = 1 Then
          If IsUserMode(FromSock, "r") Then UnSetUserMode FromSock, ServerHost, "r"
          NickCheck FromSock, Clients(FromSock).Nick
        End If
      Else
        SendError FromSock, 431, ":No nickname given"
      End If
    End If
    
  Case "NAMES"
    If Len(args(1)) > 0 Then
      Do Until Len(origline$) = 0
        xp = InStr(1, origline$, ",")
        If xp < 1 Then xp = Len(origline$) + 1
        cc$ = Left$(origline$, xp - 1)
        origline$ = Mid$(origline$, xp + 1)
        canview% = ChanExists(cc$)
        If canview% > 0 Then
          If Clients(FromSock).Chans(canview%) = 0 Then
            If IsChanMode(canview%, "ps") Then canview% = 0
          End If
        End If
        If canview% > 0 Then SendNames FromSock, canview%
      Loop
    Else
      'code for seeings all users on all channels goes here. but EWWW, massive CPU usage right thar.
    End If
    SM FromSock, ":" + ServerHost + " 366 " + Clients(SockVal).Nick + " :End of /NAMES list"
  
  Case "USERHOST"
    If Len(args(1)) = 0 Then SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
    whichuser% = FindClient(args(1))
    If whichuser% = -1 Then SendError FromSock, 401, args(1) + " :No such nick/channel": Exit Sub
    SM FromSock, ":" + ServerHost + " 302 " + Clients(FromSock).Nick + " :" + Clients(whichuser%).Nick + "=+" + Clients(whichuser%).User + "@" + Clients(whichuser%).Host
      
  Case "MODE"
    If Len(args(1)) = 0 Then SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
    If Left$(args(1), 1) = "#" Then
      newchan = ChanExists(args(1))
      If newchan = 0 Then
        SendError FromSock, 403, args(1) + " :No such channel"
      Else
        'MsgBox args(2)
        If Len(args(2)) < 2 Then 'requesting chan's modes
          extra$ = ""
          If IsUserMode(FromSock, "ao") = 1 Or IsUserChanMode(FromSock, newchan, "qo") = 1 Then
            For n = 1 To Len(Chans(newchan).Modes)
              Select Case Mid$(Chans(newchan).Modes, n, 1)
                Case "l"
                  extra$ = extra$ + Str$(Chans(newchan).ULimit)
              End Select
            Next n
          End If
          SM FromSock, ":" + ServerHost + " 324 " + Clients(FromSock).Nick + " " + Chans(newchan).Title + " +" + Chans(newchan).Modes + extra$
          SM FromSock, ":" + ServerHost + " 329 " + Clients(FromSock).Nick + " " + Chans(newchan).Title + Str$(Chans(newchan).Created)
        Else
          curop$ = Left$(args(2), 1)
          If curop$ <> "+" And curop$ <> "-" Then SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
          carg = 3
          For n = 1 To Len(args(2))
            'If Len(args(carg)) = 0 Then Exit For
            cc$ = Mid$(args(2), n, 1)
            If cc$ <> "b" And IsUserMode(FromSock, "ao") = 0 And IsUserChanMode(FromSock, newchan, "qo") = 0 Then SendError FromSock, 482, Chans(newchan).Title + " :You're not channel operator": Exit Sub
            Select Case cc$
              Case "b"
                If Len(args(3)) = 0 Then 'requesting chan's ban list
                  For nu = 0 To 1000
                    If Len(Chans(newchan).Bans(nu)) > 0 Then SM FromSock, ":" + ServerHost + " 367 " + Clients(FromSock).Nick + " " + Chans(newchan).Title + " " + Chans(newchan).Bans(nu)
                  Next nu
                  SM FromSock, ":" + ServerHost + " 368 " + Clients(FromSock).Nick + " " + Chans(newchan).Title + " :End of Channel Ban List"
                Else
                  'do ban code stuffs here
                  If IsUserMode(FromSock, "ao") = 0 And IsUserChanMode(FromSock, newchan, "qo") = 0 Then SendError FromSock, 482, Chans(newchan).Title + " :You're not channel operator": Exit Sub
                  BroadcastChan newchan, ":" + FullHost(FromSock) + " MODE " + Chans(newchan).Title + " " + curop$ + "b " + args(carg), -1
                  For nu = 0 To 1000
                    If curop$ = "+" Then
                      If Len(Chans(newchan).Bans(nu)) = 0 Then
                        Chans(newchan).Bans(nu) = args(carg)
                        Exit For
                      End If
                    Else
                      If Chans(newchan).Bans(nu) = args(carg) Then
                        Chans(newchan).Bans(nu) = ""
                        Exit For
                      End If
                    End If
                  Next nu
                  carg = carg + 1
                End If
              'Case "q"
              
              Case "o", "v"
                newuser% = FindClient(args(carg))
                If newuser% > -1 Then
                  If Clients(newuser%).Chans(newchan) = 0 Then
                    SendError FromSock, 441, Clients(newuser%).Nick + " " + Chans(newchan).Title + " :They aren't on that channel"
                  Else
                    If curop$ = "+" Then
                      SetUserChanMode newchan, FullHost(FromSock), newuser%, cc$
                      ChanMaint newchan
                      SendModes FromSock, newchan
                    Else
                      UnSetUserChanMode newchan, FullHost(FromSock), newuser%, cc$
                      ChanMaint newchan
                      SendModes FromSock, newchan
                    End If
                  End If
                Else
                  SendError FromSock, 401, args(carg) + " :No such nick/channel"
                End If
                carg = carg + 1
              
              Case "t", "n", "m", "p", "s", "i"
                If curop$ = "+" Then
                  SetChanMode newchan, FullHost(FromSock), cc$
                  ChanMaint newchan
                  SendModes FromSock, newchan
                Else
                  UnSetChanMode newchan, FullHost(FromSock), cc$
                  ChanMaint newchan
                  SendModes FromSock, newchan
                End If
              Case "k"
                If curop$ = "+" Then
                  If Len(args(carg)) > 0 Then
                    If InStr(1, Chans(newchan).Modes, cc$) < 1 Then Chans(newchan).Modes = Chans(newchan).Modes + cc$
                    BroadcastChan newchan, ":" + FullHost(FromSock) + " MODE " + Chans(newchan).Title + " +k " + args(carg), -1
                    Chans(newchan).Key = args(carg)
                  End If
                Else
                  UnSetChanMode newchan, FullHost(FromSock), cc$
                  ChanMaint newchan
                  SendModes FromSock, newchan
                  Chans(newchan).Key = ""
                End If
                carg = carg + 1
              
              Case "l"
                If curop$ = "+" Then
                  If Val(args(carg)) > 0 Then
                    If InStr(1, Chans(newchan).Modes, cc$) < 1 Then Chans(newchan).Modes = Chans(newchan).Modes + cc$
                    Chans(newchan).ULimit = Val(args(carg))
                    BroadcastChan newchan, ":" + FullHost(FromSock) + " MODE " + Chans(newchan).Title + " :+l" + Str$(Chans(newchan).ULimit), -1
                    ChanMaint newchan
                    SendModes FromSock, newchan
                  End If
                Else
                  UnSetChanMode newchan, FullHost(FromSock), cc$
                  ChanMaint newchan
                  SendModes FromSock, newchan
                  Chans(newchan).ULimit = 0
                End If
                carg = carg + 1
              
              Case "+", "-"
                curop$ = cc$
              Case Else
                SendError FromSock, 472, cc$ + " :is unknown mode char to me"
            End Select
          Next n
        End If
      End If
    Else
      newuser% = FindClient(args(1))
      If newuser% <> FromSock And IsUserMode(FromSock, "qo") = 0 Then SendError FromSock, 502, ":Can't change mode for other users": Exit Sub
      If Len(args(2)) = 0 Then
        SM FromSock, ":" + ServerHost + " 221 " + Clients(newuser%).Nick + " +" + Clients(newuser%).Modes
      Else
        If Len(args(1)) < 2 Then SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
        curop$ = "+"
        For n = 1 To Len(args(1))
          cc$ = Mid$(args(1), n, 1)
          Select Case cc$
            Case "+", "-"
              curop$ = cc$
            Case "i"
              If curop$ = "+" Then
                SetUserMode newuser%, FullHost(FromSock), cc$
              Else
                UnSetUserMode newuser%, FullHost(FromSock), cc$
              End If
          End Select
        Next n
      End If
    End If
  
  Case "TOPIC"
    If Len(args(1)) = 0 Then SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
    newchan = ChanExists(args(1))
    If newchan = 0 Then
      SendError FromSock, 401, args(1) + " :No such nick/channel"
    Else
      newtopic$ = Mid$(origline$, Len(args(1)) + 1)
      If Len(newtopic$) = 0 Then
        SendTopic FromSock, newchan
      Else
        If Clients(FromSock).Chans(newchan) = 0 Then SendError FromSock, 442, Chans(newchan).Title + " :You're not on that channel": Exit Sub
        If IsChanMode(newchan, "t") <> 0 And IsUserChanMode(FromSock, newchan, "qo") = 0 Then
          SendError FromSock, 482, Chans(newchan).Title + " :You're not channel operator (+t is set)"
          Exit Sub
        End If
        BroadcastChan newchan, ":" + FullHost(FromSock) + " TOPIC " + Chans(newchan).Title + newtopic$, -1
        Chans(newchan).Topic = newtopic$
        Chans(newchan).TopicTime = UnixTimeEnc&(Date$, Timer)
        Chans(newchan).TopicUser = Clients(FromSock).Nick
        IndexDB = IsChanReg(Chans(newchan).Title)
        If IndexDB > 0 Then
          ChanDB(IndexDB).Topic = newtopic$
          ChanDB(IndexDB).TopicUser = Clients(FromSock).Nick
          ChanDB(IndexDB).TopicTime = Chans(newchan).TopicTime
          ServMaint
        End If
      End If
    End If
    
  Case "WHOIS"
    If Len(args(1)) = 0 Then SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
    Do
      xp = InStr(1, args(1), ",")
      If xp < 1 Then xp = Len(args(1)) + 1
      tmpuser$ = Left$(args(1), xp - 1)
      args(1) = Mid$(args(1), xp + 1)
      newuser% = FindClient(tmpuser$)
      If newuser% = -1 Then
        SendError FromSock, 401, tmpuser$ + " :No such nick/channel"
      Else
        SM FromSock, ":" + ServerHost + " 311 " + Clients(FromSock).Nick + " " + Clients(newuser%).Nick + " " + Clients(newuser%).User + " " + Clients(newuser%).Host + " * :" + Clients(newuser%).RealName
        If IsUserMode(newuser%, "ao") Then SM FromSock, ":" + ServerHost + " 313 " + Clients(FromSock).Nick + " " + Clients(newuser%).Nick + " :is an IRC operator"
        SM FromSock, ":" + ServerHost + " 317 " + Clients(FromSock).Nick + " " + Clients(newuser%).Nick + Str$(UnixTimeEnc&(Date$, Timer) - Clients(newuser%).LastAction) + Str$(Clients(newuser%).ConnectTime) + " :seconds idle, signon time"
        SM FromSock, ":" + ServerHost + " 318 " + Clients(FromSock).Nick + " " + Clients(newuser%).Nick + " :End of /WHOIS list"
      End If
    Loop Until Len(args(1)) = 0

  Case "VERSION"
    SM FromSock, ":" + ServerHost + " 351 " + Clients(FromSock).Nick + " 1.1 " + ServerHost + " :Running tsIRCd. Server is experimental and under development."
  
  Case "KICK"
    If Len(args(2)) = 0 Then SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
    newchan = ChanExists(args(1))
    If newchan > 0 Then
      If IsUserChanMode(FromSock, newchan, "qo") = 0 Then SendError FromSock, 482, Chans(newchan).Title + " :You're not channel operator": Exit Sub
      kickreason$ = Mid$(origline$, Len(args(1)) + Len(args(2)) + 2)
      kickwho% = FindClient(args(2))
      If kickwho% = -1 Then SendError FromSock, 401, args(1) + " :No such nick/channel": Exit Sub
      If Clients(kickwho%).Chans(newchan) = 0 Then SendError FromSock, 441, Clients(kickwho%).Nick + " " + Chans(newchan).Title + " :They aren't on that channel": Exit Sub
      BroadcastChan newchan, ":" + FullHost(FromSock) + " KICK " + Chans(newchan).Title + " " + Clients(kickwho%).Nick + " " + kickreason$, -1
      Clients(kickwho%).Chans(newchan) = 0
      Chans(newchan).UMode(kickwho%) = ""
    Else
      SendError FromSock, 403, args(1) + " :No such channel"
    End If

  Case "AWAY"
    If Len(args(1)) = 0 Then
      SM FromSock, ":" + ServerHost + " 305 " + Clients(FromSock).Nick + " :You are no longer marked as being away"
      Clients(FromSock).IsAway = 0
      Clients(FromSock).AwayMsg = ""
    Else
      SM FromSock, ":" + ServerHost + " 306 " + Clients(FromSock).Nick + " :You are now marked as being away"
      Clients(FromSock).IsAway = 1
      Clients(FromSock).AwayMsg = origline$
    End If
  Case "REHASH"
    If IsUserMode(FromSock, "a") = 0 Then SendError FromSock, 481, Clients(FromSock).Nick + " :Permission Denied - You're not an IRC operator": Exit Sub
    Select Case LCase$(args(1))
      Case "motd"
        RehashMOTD
        For n = 0 To frmMain.Sock.UBound
          If n > frmMain.Sock.UBound Then Exit For
          If frmMain.Sock(n).State = sckConnected And IsUserMode(Int(n), "a") Then
            SM Int(n), ":" + ServerHost + " NOTICE " + Clients(n).Nick + " :*** NOTICE: " + Chr$(2) + FullHost(FromSock) + Chr$(2) + " is forcing re-reading of MOTD file."
          End If
        Next n
      Case Else
        RehashConfig
        For n = 0 To frmMain.Sock.UBound
          If n > frmMain.Sock.UBound Then Exit For
          If frmMain.Sock(n).State = sckConnected And IsUserMode(Int(n), "a") Then
            SM Int(n), ":" + ServerHost + " NOTICE " + Clients(n).Nick + " :*** NOTICE: " + Chr$(2) + FullHost(FromSock) + Chr$(2) + " is rehashing server configuration file."
          End If
        Next n
    End Select
  
  Case "KILL"
    If IsUserMode(FromSock, "ao") = 0 Then SendError FromSock, 481, Clients(FromSock).Nick + " :Permission Denied - You're not an IRC operator": Exit Sub
    If Len(args(1)) = 0 Then SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
    killwho% = FindClient(args(1))
    If killwho% = -1 Then SendError FromSock, 401, args(1) + " :No such nick/channel": Exit Sub
    killreason$ = Mid$(origline$, Len(args(1)) + 2)
    If Len(killreason$) = 0 Then killreason$ = "No reason given."
    QuitClient killwho%, "Local kill by " + Clients(FromSock).Nick + " [Reason" + killreason$ + "]"
  
  Case "ISON"
    If Len(origline$) = 0 Then If Len(args(1)) = 0 Then SendError FromSock, 461, args(0) + " :Not enough parameters": Exit Sub
    isonrpl$ = ""
    Do
      xp = InStr(1, origline$, " ")
      If xp < 1 Then xp = Len(origline$) + 1
      curnick$ = LTrim$(RTrim$(Left$(origline$, xp - 1)))
      origline$ = Mid$(origline$, xp + 1)
      If FindClient(curnick$) > -1 Then isonrpl$ = isonrpl$ + " " + curnick$
    Loop Until Len(origline$) = 0
    SM FromSock, ":" + ServerHost + " 303 " + Clients(FromSock).Nick + " :" + isonrpl$

  Case "DIE"
    If IsUserMode(FromSock, "a") = 0 Then SendError FromSock, 481, Clients(FromSock).Nick + " :Permission Denied - You're not an IRC administrator": Exit Sub
    Shutdown "Command /DIE given by " + Clients(FromSock).Nick
    
  Case "LIST"
    SendList FromSock, args(1)
  Case "PING"
    SM FromSock, ":" + ServerHost + " PONG " + ServerHost + " " + origline$
  Case "PONG"
    Clients(FromSock).GotPong = 1
  Case "QUIT"
    QuitClient FromSock, origline$
  Case Else
    If Len(args(0)) > 0 Then SendError FromSock, 421, args(0) + " :Unknown command"
End Select
End Sub

Public Sub SendError(ToSock As Integer, ErrVal As Integer, Message As String)
SM ToSock, ":" + ServerHost + Str$(ErrVal) + " " + Clients(ToSock).Nick + " " + Message
End Sub

Public Function NormalizeNick$(InNick As String)
NewNick$ = ""
For n = 1 To Len(InNick)
  cc$ = Mid$(InNick, n, 1)
  If InStr(1, NickChars, LCase$(cc$)) < 1 Then cc$ = ""
  NewNick$ = NewNick$ + cc$
Next n
NormalizeNick$ = Left$(NewNick$, NickLen)
End Function

Public Function NickInUse(CheckNick As String)
For n = 1 To SQct
  If LCase$(CheckNick) = LCase$(Reserved(n)) Then NickInUse = True: Exit Function
Next n

For n = 0 To 1000
  If LCase$(Clients(n).Nick) = LCase$(CheckNick) Then NickInUse = True: Exit Function
Next n
NickInUse = False
End Function

Public Function NormalizeChan$(InChan As String)
newchan$ = ""
For n = 1 To Len(InChan)
  cc$ = Mid$(InChan, n, 1)
  If InStr(1, ChanChars, LCase$(cc$)) < 1 Then cc$ = ""
  newchan$ = newchan$ + cc$
Next n
NormalizeChan$ = Left$(newchan$, ChanLen)
End Function

Public Sub SendGreeting(ToSock As Integer)
SM ToSock, ":" + ServerHost + " 001 " + Clients(ToSock).Nick + " :Welcome to the " + Chr$(2) + ServerName + Chr$(2) + " Internet Relay Chat network " + Clients(ToSock).Nick
SM ToSock, ":" + ServerHost + " 002 " + Clients(ToSock).Nick + " :Your host is " + ServerHost + "[" + frmMain.sckListen(0).LocalIP + "], running " + Build + ""
SM ToSock, ":" + ServerHost + " 003 " + Clients(ToSock).Nick + " :This server was created " + Creation
SM ToSock, ":" + ServerHost + " 004 " + Clients(ToSock).Nick + " " + ServerHost + " " + Build
SM ToSock, ":" + ServerHost + " 005 " + Clients(ToSock).Nick + " NICKLEN=" + Mid$(Str$(NickLen), 2) + " PREFIX=(qov)~@+ STATUSMSG=~@+ TOPICLEN=" + Mid$(Str$(TopicLen), 2) + " NETWORK=" + ServerName + " MAXTARGETS=1 CHANTYPES=# :are supported by this server"
SM ToSock, ":" + ServerHost + " 005 " + Clients(ToSock).Nick + " CHANLIMIT=#:50 CHANMODES=bklmt AWAYLEN=160 :are supported by this server"
SM ToSock, ":" + ServerHost + " 251 " + Clients(ToSock).Nick + " :There are" + Str$(OnlineCount) + " clients and 1 servers"
SM ToSock, ":" + ServerHost + " 254 " + Clients(ToSock).Nick + Str$(ChanCount) + " :channels formed"
SM ToSock, ":" + ServerHost + " 265 " + Clients(ToSock).Nick + " :Current local users:" + Str$(OnlineCount) + "  Max: 1000"
SendMOTD ToSock
If ServType = 1 Then NickCheck ToSock, Clients(ToSock).Nick
End Sub

Public Function OnlineCount()
tmpcount = 0
For n = 0 To frmMain.Sock.UBound
  If frmMain.Sock(n).State = sckConnected Then tmpcount = tmpcount + 1
Next n
OnlineCount = tmpcount
End Function

Public Function ChanCount()
tmpcount = 0
For n = 0 To 1000
  If Len(Chans(n).Title) > 0 Then tmpcount = tmpcount + 1
Next n
ChanCount = tmpcount
End Function

Public Sub RehashMOTD()
nf = FreeFile
Open WorkDir + "motd.txt" For Binary As #nf: Close #nf
Open WorkDir + "motd.txt" For Input As #nf
MOTD = ""
Do Until EOF(nf)
  Line Input #nf, tmpin$
  If Len(tmpin$) = 0 Then tmpin$ = " "
  MOTD = MOTD + tmpin$ + vbCrLf
Loop
Close #nf
End Sub

Public Sub SendMOTD(ToSock As Integer)
SM ToSock, ":" + ServerHost + " 375 " + Clients(ToSock).Nick + " :" + ServerHost + " Message Of The Day"
tmpmotd$ = MOTD
Do Until Len(tmpmotd$) = 0
  xp = InStr(1, tmpmotd$, vbCrLf)
  SM ToSock, ":" + ServerHost + " 372 " + Clients(ToSock).Nick + " :" + Left$(tmpmotd$, xp - 1)
  tmpmotd$ = Mid$(tmpmotd$, xp + 2)
Loop
SM ToSock, ":" + ServerHost + " 376 " + Clients(ToSock).Nick + " :End of /MOTD"
End Sub

Public Function FullHost$(SockVal As Integer)
FullHost$ = Clients(SockVal).Nick + "!" + Clients(SockVal).User + "@" + Clients(SockVal).Host
End Function

Public Function UnixTimeEnc&(dateval As String, ByVal timeval As Long)
'On Error Resume Next
Dim temptime, month, day, year As Long
timezone% = -6
tempdate$ = dateval
divider = InStr(1, tempdate$, "-"): month = Val(Left$(tempdate$, divider - 1))
tempdate$ = Mid$(tempdate$, divider + 1)
divider = InStr(1, tempdate$, "-"): day = Val(Left$(tempdate$, divider - 1))
tempdate$ = Mid$(tempdate$, divider + 1)
year = Val(tempdate$)

temptime = 0
For calcval = 1970 To year - 1
   'determine if a year is a leap year
   If (calcval + 2) / 4 = Fix((calcval + 2) / 4) Then
      daysinyear = 366
   Else
      daysinyear = 365
   End If

   temptime = temptime + (daysinyear * 86400)
Next calcval

For calcval = 1 To month - 1
   Select Case calcval
      Case 1, 3, 5, 7, 8, 10, 11, 12 '31 day months
         temptime = temptime + (31 * 86400)

      Case 2 'Feb = 28 days in common year, 29 days in leap year
        If (year + 2) / 4 = Fix((year + 2) / 4) Then 'it's a leap year
          temptime = temptime + (29 * 86400)
        Else 'it's a common year
          temptime = temptime + (28 * 86400)
        End If

      Case Else
         temptime = temptime + (30 * 86400)
   End Select
Next calcval

temptime = temptime + ((day - 1) * 86400) + timeval + (-timezone * 3600)
UnixTimeEnc& = temptime - (3600 * timezone%)
End Function

Public Sub DoPing(SockVal As Integer)
Clients(SockVal).GotPong = 0
pingtime& = UnixTimeEnc&(Date$, Timer)
Clients(SockVal).LastPing = pingtime&
SM SockVal, "PING :" + ServerHost
End Sub

Public Sub QuitClient(SockVal As Integer, Reason As String)
SM SockVal, "ERROR :Closing Link: " + Clients(SockVal).Nick + "[" + Clients(SockVal).Host + " (" + Reason + ")"
KillFlag(SockVal) = 1
frmMain.Sock(SockVal).Close
For n = 1 To 1000
  If Clients(SockVal).Chans(n) = 1 Then
    Clients(SockVal).Chans(n) = 0
    userchancount% = 0
    For nu = 0 To frmMain.Sock.UBound
      If nu > frmMain.Sock.UBound Then Exit For
      If Clients(nu).Chans(n) = 1 Then userchancount% = userchancount% + 1
    Next nu
    If userchancount% > 0 Then
      If InStr(1, Reason, ":") < 1 Then extra$ = ":" Else extra$ = ""
      BroadcastChan Int(n), ":" + FullHost(SockVal) + " QUIT " + extra$ + Reason, SockVal
    Else
      ClearChanSlot Int(n)
    End If
  End If
Next n
End Sub

Public Function JoinChan(SockVal As Integer, ChannelName As String, ChannelKey As String) As Integer
Dim newchan As Integer
newchan = ChanExists(ChannelName)
If newchan = 0 Then newchan = CreateChan(ChannelName): Chans(newchan).Key = ChannelKey: makeop = 1 Else makeop = 0
If Clients(SockVal).Chans(newchan) = 0 Then
  For n = 0 To 1000 'see if banned
    If Len(Chans(newchan).Bans(n)) > 0 And MatchMask%(FullHost(SockVal), Chans(newchan).Bans(n)) = 1 Then SendError SockVal, 474, Chans(newchan).Title + " :Cannot join channel (+b)": Exit Function
  Next n
  If IsChanMode(newchan, "k") Then
    If Chans(newchan).Key <> ChannelKey Then SendError SockVal, 475, Chans(newchan).Title + " :Cannot join channel (+k)": Exit Function
  End If
  If IsChanMode(newchan, "i") Then
    If Clients(SockVal).Invite(newchan) = 0 Then SendError SockVal, 473, Chans(newchan).Title + " :Cannot join channel (+i)": Exit Function
  End If
  If IsChanMode(newchan, "l") Then
    If Chans(newchan).ULimit > 0 Then
      usersinchan% = 0
      For nu = 0 To frmMain.Sock.UBound
        If nu > frmMain.Sock.UBound Then Exit For
        If Clients(nu).Chans(newchan) = 1 Then usersinchan% = usersinchan% + 1
      Next nu
      If usersinchan% >= Chans(newchan).ULimit Then SendError SockVal, 471, Chans(newchan).Title + " :Cannot join channel (+l)": Exit Function
    End If
  End If
  Clients(SockVal).Chans(newchan) = 1
  JoinChan = newchan
  SM SockVal, ":" + ServerHost + " 366 " + Clients(SockVal).Nick + " " + Chans(newchan).Title + " :End of /NAMES list"
  If makeop = 1 Then
    Chans(newchan).Created = UnixTimeEnc&(Date$, Timer)
    Chans(newchan).IndexDB = 0
    For n = 1 To ChanDBCount
      If LCase$(ChannelName) = LCase$(Chans(n).Title) Then Chans(newchan).IndexDB = n: Exit For
    Next n
    If Chans(newchan).IndexDB > 0 Then
      Chans(newchan).Topic = ChanDB(Chans(newchan).IndexDB).Topic
      Chans(newchan).TopicUser = ChanDB(Chans(newchan).IndexDB).TopicUser
      Chans(newchan).TopicTime = ChanDB(Chans(newchan).IndexDB).TopicTime
    End If
    SetChanMode newchan, ServerHost, "n"
    SetUserChanMode newchan, ServerHost, SockVal, "o"
  End If
  BroadcastChan newchan, ":" + FullHost(SockVal) + " JOIN " + Chans(newchan).Title, -1
  'SendModes SockVal, newchan
  SendNames SockVal, newchan
  If Len(Chans(newchan).Topic) > 0 Then SendTopic SockVal, newchan
  End If
'If Chans(newchan).IndexDB > 0 Then
ChanMaint newchan
End Function

Public Function CreateChan(ChannelName As String)
For n = 1 To 1000
  If Len(Chans(n).Title) = 0 Then Chans(n).Title = ChannelName: CreateChan = n: Exit Function
Next n
End Function

Public Function ChanExists(ChannelName As String)
For n = 1 To 1000
  If LCase$(Chans(n).Title) = LCase$(ChannelName) Then ChanExists = n: Exit Function
Next n
End Function

Public Sub SendTopic(SockVal As Integer, ChanVal As Integer)
If Len(Chans(ChanVal).Topic) = 0 Then
    SM SockVal, ":" + ServerHost + " 331 " + Clients(SockVal).Nick + " " + Chans(ChanVal).Title + " :No topic is set"
  Else
    SM SockVal, ":" + ServerHost + " 332 " + Clients(SockVal).Nick + " " + Chans(ChanVal).Title + Chans(ChanVal).Topic
    SM SockVal, ":" + ServerHost + " 333 " + Clients(SockVal).Nick + " " + Chans(ChanVal).Title + " " + Chans(ChanVal).TopicUser + Str$(Chans(ChanVal).TopicTime)
End If
End Sub

Public Sub BroadcastChan(ChanVal As Integer, Message As String, Exclude As Integer)
Dim n As Integer
If ChanVal = 0 Then Exit Sub
For n = 0 To 1000
  If Clients(n).Chans(ChanVal) = 1 And n <> Exclude Then SM n, Message: DoEvents
Next n
End Sub

Public Sub PartChan(SockVal As Integer, ChanVal As Integer, Reason As String)
BroadcastChan ChanVal, ":" + FullHost(SockVal) + " PART " + Chans(ChanVal).Title + " :" + Reason, -1
Clients(SockVal).Chans(ChanVal) = 0
Chans(ChanVal).UMode(SockVal) = ""

For n = 0 To frmMain.Sock.UBound
  For nu = 1 To 1000
    If Clients(n).Chans(nu) = 1 Then Exit Sub
  Next nu
Next n

'since nobody is left, we reset all channel number info here
ClearChanSlot ChanVal
End Sub

Public Sub SendNames(ToSock As Integer, ChanVal As Integer)
ct = 0
tmpnames$ = ":" + ServerHost + " 353 " + Clients(ToSock).Nick + " = " + Chans(ChanVal).Title + " :"
For n = 0 To frmMain.Sock.UBound
  If Clients(n).Chans(ChanVal) = 1 Then
    tmpnick$ = Clients(n).Nick
    If InStr(1, Chans(ChanVal).UMode(n), "v") Then tmpnick$ = "+" + Clients(n).Nick
    If InStr(1, Chans(ChanVal).UMode(n), "o") Then tmpnick$ = "@" + Clients(n).Nick
    If InStr(1, Chans(ChanVal).UMode(n), "q") Then tmpnick$ = "~" + Clients(n).Nick
    tmpnames$ = tmpnames$ + tmpnick$
    ct = ct + 1
    If ct = 12 Then
      SM ToSock, tmpnames$
      tmpnames$ = ":" + ServerHost + " 353 " + Clients(ToSock).Nick + " = " + Chans(ChanVal).Title + " :"
      ct = 0
    Else
      tmpnames$ = tmpnames$ + " "
    End If
  End If
Next n
If ct < 12 Then SM ToSock, tmpnames$
End Sub

Public Function IsChanMode(ChanVal As Integer, Modes As String)
For n = 1 To Len(Modes)
  If InStr(1, Chans(ChanVal).Modes, Mid$(Modes, n, 1)) Then IsChanMode = 1: Exit Function
Next n
End Function

Public Function IsUserChanMode(SockVal As Integer, ChanVal As Integer, Modes As String)
For n = 1 To Len(Modes)
  If InStr(1, Chans(ChanVal).UMode(SockVal), Mid$(Modes, n, 1)) > 0 Then IsUserChanMode = 1: Exit Function
Next n
End Function

Public Function IsUserMode(SockVal As Integer, WhatMode As String)
For n = 1 To Len(WhatMode)
  cc$ = Mid$(WhatMode, n, 1)
  If InStr(1, Clients(SockVal).Modes, cc$) Then IsUserMode = 1: Exit Function
Next n
End Function

Public Sub ChangeNick(SockVal As Integer, NewNick As String)
tmpub = frmMain.Sock.UBound
For n = 0 To tmpub
  Temp(n) = 0
Next n
For n = 0 To tmpub
  For nu = 1 To 1000
    If Clients(n).Chans(nu) = Clients(SockVal).Chans(nu) Then Temp(n) = 1
  Next nu
Next n
Temp(SockVal) = 1
For n = 0 To tmpub
  If Temp(n) = 1 Then SM Int(n), ":" + FullHost(SockVal) + " NICK " + NewNick
Next n
Clients(SockVal).NickTimer = 0
Clients(SockVal).Nick = NewNick
End Sub

Public Sub SetChanMode(ChanVal As Integer, ByWho As String, WhatMode As String)
BroadcastChan ChanVal, ":" + ByWho + " MODE " + Chans(ChanVal).Title + " :+" + WhatMode + extra$, -1
For n = 1 To Len(WhatMode)
  cc$ = Mid$(WhatMode, n, 1)
  If InStr(1, Chans(ChanVal).Modes, cc$) < 1 Then Chans(ChanVal).Modes = Chans(ChanVal).Modes + cc$
  Select Case cc$
    Case "s"
      UnSetChanMode ChanVal, ServerHost, "p"
    Case "p"
      UnSetChanMode ChanVal, ServerHost, "s"
  End Select
Next n
'If ByWho <> "ChanServ" Then ChanMaint ChanVal
End Sub

Public Sub SetUserMode(UserVal As Integer, ByWho As String, WhatMode As String)
SM UserVal, ":" + ByWho + " MODE " + Clients(UserVal).Nick + " +" + WhatMode
For n = 1 To Len(WhatMode)
  cc$ = Mid$(WhatMode, n, 1)
  If InStr(1, Clients(UserVal).Modes, cc$) < 1 Then Clients(UserVal).Modes = Clients(UserVal).Modes + cc$
Next n
End Sub

Public Sub SetUserChanMode(ChanVal As Integer, ByWho As String, UserVal As Integer, WhatMode As String)
If UserVal < 0 Then Exit Sub
For n = 1 To Len(WhatMode)
  cc$ = Mid$(WhatMode, n, 1)
  If InStr(1, Chans(ChanVal).UMode(UserVal), cc$) < 1 Then
    BroadcastChan ChanVal, ":" + ByWho + " MODE " + Chans(ChanVal).Title + " +" + cc$ + " " + Clients(UserVal).Nick, -1
    Chans(ChanVal).UMode(UserVal) = Chans(ChanVal).UMode(UserVal) + cc$
  End If
Next n
'If ByWho <> "ChanServ" Then ChanMaint ChanVal
End Sub

Public Sub UnSetUserChanMode(ChanVal As Integer, ByWho As String, UserVal As Integer, WhatMode As String)
If UserVal < 0 Then Exit Sub
For n = 1 To Len(WhatMode)
  cc$ = Mid$(WhatMode, n, 1)
  xp = InStr(1, Chans(ChanVal).UMode(UserVal), cc$)
  If xp > 0 Then
    'If xp = 1 Then
    '  Chans(ChanVal).UMode(UserVal) = Mid$(Chans(ChanVal).UMode(UserVal), 2)
    'Else
      BroadcastChan ChanVal, ":" + ByWho + " MODE " + Chans(ChanVal).Title + " -" + cc$ + " " + Clients(UserVal).Nick, -1
      Chans(ChanVal).UMode(UserVal) = Left$(Chans(ChanVal).UMode(UserVal), xp - 1) + Mid$(Chans(ChanVal).UMode(UserVal), xp + 1)
    'End If
  End If
Next n
'If ByWho <> "ChanServ" Then ChanMaint ChanVal
End Sub

Public Sub UnSetChanMode(ChanVal As Integer, ByWho As String, WhatMode As String)
For n = 1 To Len(WhatMode)
  cc$ = Mid$(WhatMode, n, 1)
  xp = InStr(1, Chans(ChanVal).Modes, cc$)
  If xp > 0 Then
    'If xp = 1 Then
    '  Chans(ChanVal).Modes = Mid$(Chans(ChanVal).Modes, 2)
    'Else
      BroadcastChan ChanVal, ":" + ByWho + " MODE " + Chans(ChanVal).Title + " -" + WhatMode, -1
      Chans(ChanVal).Modes = Left$(Chans(ChanVal).Modes, xp - 1) + Mid$(Chans(ChanVal).Modes, xp + 1)
    'End If
  End If
Next n
'If ByWho <> "ChanServ" Then ChanMaint ChanVal
End Sub

Public Sub UnSetUserMode(UserVal As Integer, ByWho As String, WhatMode As String)
For n = 1 To Len(WhatMode)
  cc$ = Mid$(WhatMode, n, 1)
  xp = InStr(1, Clients(UserVal).Modes, cc$)
  If xp > 0 Then
    'If xp = 1 Then
    '  Clients(UserVal).Modes = Mid$(Clients(UserVal).Modes, 2)
    'Else
      SM UserVal, ":" + ByWho + " MODE " + Clients(UserVal).Nick + " -" + WhatMode
      Clients(UserVal).Modes = Left$(Clients(UserVal).Modes, xp - 1) + Mid$(Clients(UserVal).Modes, xp + 1)
    'End If
  End If
Next n
End Sub

Public Function FindClient(NickName As String)
For n = 0 To frmMain.Sock.UBound
  If LCase$(Clients(n).Nick) = LCase$(NickName) Then FindClient = n: Exit Function
Next n
FindClient = -1
End Function

Public Sub SendList(SockVal As Integer, Contains As String)
SM SockVal, ":" + ServerHost + " 321 " + Clients(SockVal).Nick + " Channel :Users  Name"
For n = 1 To 1000
  If Len(Contains) = 0 Then
    If Len(Chans(n).Title) > 0 Then showit = 1 Else showit = 0
  Else
    If InStr(1, LCase$(Chans(n).Title), LCase$(Contains)) Then showit = 1 Else showit = 0
  End If
  If IsChanMode(Int(n), "sp") Then showit = 0
  If showit = 1 Then
    usersinchan% = 0
    For nu = 0 To frmMain.Sock.UBound
      If Clients(nu).Chans(n) = 1 Then usersinchan% = usersinchan% + 1
    Next nu
    SM SockVal, ":" + ServerHost + " 322 " + Clients(SockVal).Nick + " " + Chans(n).Title + Str$(usersinchan%) + Chans(n).Topic
  End If
Next n
SM SockVal, ":" + ServerHost + " 323 " + Clients(SockVal).Nick + " :End of /LIST"
End Sub


Public Function MatchMask%(check$, against$)
chk$ = against$
tmp$ = check$

Match% = 1
For chkm = 1 To Len(chk$)
  curchar$ = Mid$(chk$, chkm, 1)
  nextchar$ = Mid$(chk$, chkm + 1, 1)
  If curchar$ = "*" Then
    If Len(nextchar$) = 0 Then Exit For
    xpm = InStr(tmp$, nextchar$)
    If xpm < 1 Then Match% = 0: Exit For
    tmp$ = Mid$(tmp$, xpm + 1)
    chkm = chkm + 1
  Else
    If Left$(tmp$, 1) <> curchar$ Then Match% = 0: Exit For
    tmp$ = Mid$(tmp$, 2)
  End If
  If Len(tmp$) = 0 Then Exit For
Next chkm
MatchMask% = Match%
End Function

Public Sub ClearChanSlot(ChanVal As Integer)
Chans(ChanVal).Modes = ""
Chans(ChanVal).Title = ""
Chans(ChanVal).Topic = ""
Chans(ChanVal).TopicTime = 0
Chans(ChanVal).TopicUser = ""
Chans(ChanVal).Created = 0
Chans(ChanVal).Key = ""
Chans(ChanVal).ULimit = 0
Chans(ChanVal).IndexDB = 0
For n = 0 To 1000
  Chans(ChanVal).UMode(n) = ""
  Chans(ChanVal).Bans(n) = ""
Next n
End Sub

Public Function IsRsvd(CheckNick As String)
'For n = 1 To SQct
'  If LCase$(CheckNick) = LCase$(Reserved(n)) Then IsRsvd = 1: Exit Function
'Next n
Select Case LCase$(CheckNick)
  Case "nickserv", "chanserv", "operserv", "memoserv", "hostserv"
    IsRsvd = 1
  Case Else
    IsRsvd = 0
End Select
End Function

Public Sub Shutdown(Reason As String)
For n = 0 To frmMain.sckListen.UBound
  frmMain.sckListen(n).Close
Next n
For n = 0 To frmMain.Sock.UBound
  If n > frmMain.Sock.UBound Then Exit For
  If frmMain.Sock(n).State = sckConnected Then
    SM Int(n), ":" + ServerHost + " NOTICE " + Clients(n).Nick + " :" + Reason
    QuitClient Int(n), "Server shutdown."
  End If
Next n
frmMain.DieTimer.Enabled = True
End Sub

Public Sub SendModes(ToSock As Integer, ChanVal As Integer)
For n = 1 To Len(Chans(ChanVal).Modes)
  Select Case Mid$(Chans(ChanVal).Modes, n, 1)
    Case "l"
      extra$ = extra$ + Str$(Chans(ChanVal).ULimit)
  End Select
Next n
SM ToSock, ":" + ServerHost + " 324 " + Clients(ToSock).Nick + " " + Chans(ChanVal).Title + " +" + Chans(ChanVal).Modes + extra$
End Sub
