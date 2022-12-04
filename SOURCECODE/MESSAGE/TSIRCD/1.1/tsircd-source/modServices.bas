Attribute VB_Name = "modServices"
Type NickType
  Nick As String
  Pass As String
  Modes As String
  LastUsed As Long
End Type

Type ChanEntry
  Nick As String
  Level As Byte
End Type

Type ChanType
  Title As String
  Pass As String
  Modes As String
  EnforceModes As Byte
  LastUsed As Long
  FOUNDER As String
  EntryCount As Integer
  Entry(1 To 1000) As ChanEntry
  Topic As String
  TopicUser As String
  TopicTime As Long
  Limit As Integer
  Description As String
End Type

Type NSOptions
  AutoProtect As Byte
  GuestPrefix As String
End Type

Public NickDB(1 To 1000) As NickType
Public NickDBCount As Integer
Public ChanDB(1 To 1000) As ChanType
Public ChanDBCount As Integer
Public NS As NSOptions
Const AKICK As Byte = 1
Const AVOICE As Byte = 2
Const AOP As Byte = 4
Const COFOUNDER As Byte = 8
Const FOUNDER As Byte = 128

Public Sub InitServices()
If ServType <> 1 Then Exit Sub

nf = FreeFile
Open WorkDir + "nick.db" For Binary As #nf: Close #nf
Open WorkDir + "nick.db" For Input As #nf

NickDBCount = 0
Do Until EOF(nf)
  Input #nf, tmpnick$: If EOF(nf) Then Exit Do
  Input #nf, tmppass$: If EOF(nf) Then Exit Do
  Input #nf, tmpmodes$: If EOF(nf) Then Exit Do
  Input #nf, tmplastused&
  NickDBCount = NickDBCount + 1
  NickDB(NickDBCount).Nick = tmpnick$
  NickDB(NickDBCount).Pass = tmppass$
  NickDB(NickDBCount).Modes = tmpmodes$
  NickDB(NickDBCount).LastUsed = tmplastused&
  If NickDBCount = 1000 Then Exit Do
Loop
Close #nf

Open WorkDir + "chan.db" For Binary As #nf: Close #nf
Open WorkDir + "chan.db" For Input As #nf

ChanDBCount = 0
Do Until EOF(nf)
  Input #nf, tmpchan$: If EOF(nf) Then Exit Do
  Input #nf, tmppass$: If EOF(nf) Then Exit Do
  Input #nf, tmpmodes$: If EOF(nf) Then Exit Do
  Input #nf, tmpenforce%: If EOF(nf) Then Exit Do
  Input #nf, tmplastused&: If EOF(nf) Then Exit Do
  Input #nf, tmpfounder$: If EOF(nf) Then Exit Do
  Input #nf, tmplimit%: If EOF(nf) Then Exit Do
  Input #nf, tmpentries%: If EOF(nf) Then Exit Do
  Line Input #nf, tmpdesc$: If EOF(nf) Then Exit Do
  Line Input #nf, tmptopic$: If EOF(nf) Then Exit Do
  Input #nf, tmptopicby$: If EOF(nf) Then Exit Do
  Input #nf, tmptopictime&: If EOF(nf) = True And tmpentries% > 0 Then Exit Do
  ChanDBCount = ChanDBCount + 1
  ChanDB(ChanDBCount).Title = tmpchan$
  ChanDB(ChanDBCount).Pass = tmppass$
  ChanDB(ChanDBCount).Modes = tmpmodes$
  ChanDB(ChanDBCount).EnforceModes = tmpenforce%
  ChanDB(ChanDBCount).LastUsed = tmplastused&
  ChanDB(ChanDBCount).EntryCount = tmpentries%
  ChanDB(ChanDBCount).Limit = tmplimit%
  ChanDB(ChanDBCount).FOUNDER = tmpfounder$
  ChanDB(ChanDBCount).Description = tmpdesc$
  ChanDB(ChanDBCount).Topic = tmptopic$
  ChanDB(ChanDBCount).TopicUser = tmptopicby$
  ChanDB(ChanDBCount).TopicTime = tmptopictime&
  For n = 1 To tmpentries%
    Input #nf, ChanDB(ChanDBCount).Entry(n).Nick: If EOF(nf) = True Then Exit Do
    Input #nf, ChanDB(ChanDBCount).Entry(n).Level
  Next n
  If ChanDBCount = 1000 Then Exit Do
Loop
Close #nf
End Sub

Public Sub NickCheck(SockVal As Integer, Nick As String)
For n = 1 To NickDBCount
  If LCase$(NickDB(n).Nick) = LCase$(Nick) Then
    If frmMain.Sock(SockVal).State <> sckConnected Then Exit For
    SM SockVal, ":NickServ!Services@" + ServerHost + " NOTICE " + Clients(SockVal).Nick + " :This is a registered and protected nickname. If it is your"
    SM SockVal, ":NickServ!Services@" + ServerHost + " NOTICE " + Clients(SockVal).Nick + " :nick, type " + Chr$(2) + "/msg NickServ IDENTIFY password" + Chr$(2) + ". Otherwise,"
    SM SockVal, ":NickServ!Services@" + ServerHost + " NOTICE " + Clients(SockVal).Nick + " :please choose a different nick."
    If NS.AutoProtect <> 0 Then SM SockVal, ":NickServ!Services@" + ServerHost + " NOTICE " + Clients(SockVal).Nick + " :If you do not change within one minute, I will change your nick."
    Clients(SockVal).NickTimer = UnixTimeEnc&(Date$, Timer)
    Clients(SockVal).NickMatch = n
    Exit For
  End If
Next n
End Sub

Public Sub ExpireNick(SockVal As Integer)
If NS.AutoProtect = 0 Then Exit Sub
Do
  tmpnick$ = NS.GuestPrefix + Mid$(Str$(Int(Timer)), 2) + Mid$(Str$(Int(Rnd * 1000)), 2)
  If FindClient(tmpnick$) = -1 Then ChangeNick SockVal, tmpnick$: Clients(SockVal).NickTimer = 0: Exit Sub
Loop
End Sub

Public Sub HandleServiceMsg(SockVal As Integer, ServName As String, Message As String)
tmpmsg$ = Message
xp = InStr(1, tmpmsg$, ":")
If xp > 0 Then tmpmsg$ = Left$(tmpmsg$, xp - 1) + Mid$(tmpmsg$, xp + 1)
tmpmsg$ = LTrim$(tmpmsg$)
Select Case LCase$(ServName)
  Case "nickserv"
    NickServMsg SockVal, tmpmsg$
  Case "chanserv"
    ChanServMsg SockVal, tmpmsg$
  Case Else
    SM SockVal, ":" + ServerHost + " NOTICE " + Clients(SockVal).Nick + " :" + Chr$(2) + ServName + Chr$(2) + " is not a valid IRC service."
End Select
End Sub

Public Sub NickServMsg(SockVal As Integer, Message As String)
Dim args(20) As String
carg = 0
tmpmsg$ = Message
Do
  xp = InStr(1, tmpmsg$, " ")
  If xp < 1 Then xp = Len(tmpmsg$) + 1
  args(carg) = LTrim$(RTrim$(Left$(tmpmsg$, xp - 1)))
  tmpmsg$ = Mid$(tmpmsg$, xp + 1)
  carg = carg + 1
Loop Until carg = 21 Or Len(tmpmsg$) = 0

Select Case LCase$(args(0))
  Case "identify"
    If Clients(SockVal).NickMatch = 0 Then ServSay SockVal, "NickServ", "Your nick is not registered.": Exit Sub
    If IsUserMode(SockVal, "r") Then ServSay SockVal, "NickServ", "You have already identified.": Exit Sub
    If NickDB(Clients(SockVal).NickMatch).Pass = Hash.DigestStrToHexStr(args(1)) Then
      Clients(SockVal).NickTimer = 0
      ServSay SockVal, "NickServ", "Password accepted - you are now recognized."
      SetUserMode SockVal, "NickServ!Services@" + ServerHost, "r"
      For n = 1 To 1000
        If Clients(SockVal).Chans(n) = 1 Then ChanMaint Int(n)
      Next n
    Else
      ServSay SockVal, "NickServ", "Password incorrect."
    End If
    
  Case "register"
    If Len(args(1)) = 0 Then
      ServSay SockVal, "NickServ", "You have to supply a registration password."
    Else
      isused = 0
      For n = 1 To NickDBCount
        If LCase$(NickDB(n).Nick) = LCase$(Clients(SockVal).Nick) Then
          isused = 1
          ServSay SockVal, "NickServ", "The nickname " + Chr$(2) + NickDB(n).Nick + Chr$(2) + " is already registered."
          Exit For
        End If
      Next n
      If isused = 0 Then
        If NickDBCount = 1000 Then ServSay SockVal, "NickServ", "Sorry, the NickServ database is full.": Exit Sub
        For n = 1 To NickDBCount
          If Len(NickDB(n).Nick) = 0 Then newentry = n: Exit For
        Next n
        If n > NickDBCount Then newentry = NickDBCount + 1: NickDBCount = newentry
          
          NickDB(newentry).Nick = Clients(SockVal).Nick
          NickDB(newentry).Pass = Hash.DigestStrToHexStr(args(1))
          NickDB(newentry).Modes = "none"
          NickDB(newentry).LastUsed = UnixTimeEnc&(Date$, Timer)
          ServMaint
          SetUserMode SockVal, "NickServ!Services@" + ServerHost, "r"
          ServSay SockVal, "NickServ", "Success! You have registered the nickname " + Chr$(2) + NickDB(newentry).Nick + Chr$(2) + " with the password " + Chr$(2) + args(1) + Chr$(2) + "."
          ServSay SockVal, "NickServ", "Write down your password and keep it somewhere safe, as your nick will be inaccessible without it! It is case-sensitive!"
        End If
      End If

  Case "drop"
    If IsUserMode(SockVal, "r") = 0 Then
      ServSay SockVal, "NickServ", "Your nickname is not registered, or you have not yet identified. Unable to drop."
    Else
      isused = 0
      For n = 1 To NickDBCount
        If LCase$(NickDB(n).Nick) = LCase$(Clients(SockVal).Nick) Then: isused = n: Exit For
      Next n
      If isused > 0 Then
        NickDB(isused).Nick = ""
        NickDB(isused).Pass = ""
        NickDB(isused).Modes = ""
        NickDB(isused).LastUsed = 0
        ServMaint
        UnSetUserMode SockVal, "NickServ!Services@" + ServerHost, "r"
        ServSay SockVal, "NickServ", "Success! Your nickname is no longer registered with NickServ."
      Else
        ServSay SockVal, "NickServ", "Oops! There is a problem with the registration database that has caused an internal error."
      End If
    End If

  Case "ghost"
    If Len(args(2)) = 0 Then ServSay SockVal, "NickServ", "GHOST requires additional parameters.": Exit Sub
    newuser% = FindClient(args(1))
    If newuser% = -1 Then ServSay SockVal, "NickServ", "There is no such nickname currently logged into this server.": Exit Sub
    For n = 1 To NickDBCount
      If LCase$(args(1)) = LCase$(NickDB(n).Nick) Then
        If NickDB(n).Pass = Hash.DigestStrToHexStr(args(2)) Then
          QuitClient newuser%, "GHOST command used by " + Clients(SockVal).Nick
        Else
          ServSay SockVal, "NickServ", "Incorrect password."
        End If
        Exit Sub
      End If
    Next n
    If n > NickDBCount Then ServSay SockVal, "NickServ", "That nickname is not registered with NickServ."

  Case "help"
    If Len(args(1)) = 0 Then 'help overview
      ServSay SockVal, "NickServ", Chr$(2) + "NickServ" + Chr$(2) + " allows you to register a nickname and"
      ServSay SockVal, "NickServ", "prevent others from using it. The following commands"
      ServSay SockVal, "NickServ", "allow for registration and maintenance of a nick."
      ServSay SockVal, "NickServ", " "
      ServSay SockVal, "NickServ", "To use them, type " + Chr$(2) + "/msg NickServ command" + Chr$(2) + ". For more information on a"
      ServSay SockVal, "NickServ", "specific command, type " + Chr$(2) + "/msg NickServ help command" + Chr$(2) + "."
      ServSay SockVal, "NickServ", " "
      ServSay SockVal, "NickServ", "    REGISTER   Register a nickname"
      ServSay SockVal, "NickServ", "    IDENTIFY   Identify yourself with your password"
      ServSay SockVal, "NickServ", "    DROP       Cancel the registration of a nickname"
      ServSay SockVal, "NickServ", "    GHOST      Kill a connected client that is registered"
      ServSay SockVal, "NickServ", " "
      ServSay SockVal, "NickServ", "Nicknames that are not used for 30 days will be automatically"
      ServSay SockVal, "NickServ", "dropped from the registration database."
    Else 'detailed help
      Select Case LCase$(args(1))
        Case "register"
          ServSay SockVal, "NickServ", "NickServ REGISTER syntax:"
          ServSay SockVal, "NickServ", "    " + Chr$(2) + "/msg NickServ REGISTER password" + Chr$(2)
        Case "identify"
          ServSay SockVal, "NickServ", "NickServ IDENTIFY syntax:"
          ServSay SockVal, "NickServ", "    " + Chr$(2) + "/msg NickServ IDENTIFY password" + Chr$(2)
        Case "drop"
          ServSay SockVal, "NickServ", "NickServ DROP syntax:"
          ServSay SockVal, "NickServ", "    " + Chr$(2) + "/msg NickServ DROP" + Chr$(2)
          ServSay SockVal, "NickServ", "No arguments are required, but you must first be logged in"
          ServSay SockVal, "NickServ", "and identified as the nick you want to drop."
        Case "ghost"
          ServSay SockVal, "NickServ", "NickServ GHOST syntax:"
          ServSay SockVal, "NickServ", "    " + Chr$(2) + "/msg NickServ GHOST nickname password" + Chr$(2)
        Case Else
          ServSay SockVal, "NickServ", "Unknown command - " + Chr$(2) + args(1) + Chr$(2)
      End Select
    End If
  Case Else
    ServSay SockVal, "NickServ", "Unknown command " + Chr$(2) + args(0) + Chr$(2) + ". " + Chr$(34) + "/msg NickServ HELP" + Chr$(34) + " for help."
End Select
End Sub

Public Sub ServSay(SockVal As Integer, ServiceNick As String, Message As String)
SM SockVal, ":" + ServiceNick + "!Services@" + ServerHost + " NOTICE " + Clients(SockVal).Nick + " :" + Message
End Sub

Public Sub ServMaint()
nf = FreeFile
Open WorkDir + "nick.db" For Output As #nf
For n = 1 To NickDBCount
  If Len(NickDB(n).Nick) > 0 Then Print #nf, NickDB(n).Nick + "," + NickDB(n).Pass + "," + NickDB(n).Modes + "," + Str$(NickDB(n).LastUsed)
Next n
Close #nf
'Exit Sub

Open WorkDir + "chan.db" For Output As #nf
For n = 1 To ChanDBCount
  tmpdesc$ = ChanDB(n).Description
  If Len(tmpdesc$) = 0 Then tmpdesc$ = " "
  Print #nf, ChanDB(n).Title + "," + ChanDB(n).Pass + "," + ChanDB(n).Modes + "," + Str$(ChanDB(n).EnforceModes) + "," + Str$(ChanDB(n).LastUsed) + "," + ChanDB(n).FOUNDER + "," + Str$(ChanDB(n).Limit) + "," + Str$(ChanDB(n).EntryCount) + "," + tmpdesc$
  Print #nf, ChanDB(n).Topic
  Print #nf, ChanDB(n).TopicUser + "," + Str$(ChanDB(n).TopicTime)
  For nu = 1 To ChanDB(n).EntryCount
    Print #nf, ChanDB(n).Entry(nu).Nick + "," + Str$(ChanDB(n).Entry(nu).Level);
    If nu < ChanDB(n).EntryCount Then Print #nf, ","; Else Print #nf, ""
  Next nu
Next n
Close #nf
End Sub

Public Sub ChanServMsg(SockVal As Integer, Message As String)
Dim newchan As Integer
Dim args(20) As String
carg = 0
tmpmsg$ = Message
Do
  xp = InStr(1, tmpmsg$, " ")
  If xp < 1 Then xp = Len(tmpmsg$) + 1
  args(carg) = LTrim$(RTrim$(Left$(tmpmsg$, xp - 1)))
  tmpmsg$ = Mid$(tmpmsg$, xp + 1)
  carg = carg + 1
Loop Until carg = 21 Or Len(tmpmsg$) = 0

Select Case LCase$(args(0))
   Case "info"
     If Len(args(1)) = 0 Then ServSay SockVal, "ChanServ", "I need more parameters.": Exit Sub
     IndexDB = IsChanReg(args(1))
     If IndexDB = 0 Then ServSay SockVal, "ChanServ", "Channel " + Chr$(2) + args(1) + Chr$(2) + " is not registered with ChanServ.": Exit Sub
     ServSay SockVal, "ChanServ", "Information for channel " + Chr$(2) + ChanDB(IndexDB).Title + Chr$(2) + ":"
     ServSay SockVal, "ChanServ", "        Founder: " + ChanDB(IndexDB).FOUNDER
     ServSay SockVal, "ChanServ", "    Description: " + ChanDB(IndexDB).Description
     ServSay SockVal, "ChanServ", "     Registered: N/A"
     ServSay SockVal, "ChanServ", "      Last used: N/A"
     tmptopic$ = Mid$(ChanDB(IndexDB).Topic, 2)
     xp = InStr(1, tmptopic$, ":")
     If xp > 0 Then tmptopic$ = Left$(tmptopic$, xp - 1) + Mid$(tmptopic$, xp + 1)
     ServSay SockVal, "ChanServ", "     Last topic: " + tmptopic$
     ServSay SockVal, "ChanServ", "   Topic set by: " + ChanDB(IndexDB).TopicUser
   Case "help"
     Select Case LCase$(args(1))
       Case ""
         ServSay SockVal, "ChanServ", Chr$(2) + "ChanServ" + Chr$(2) + " allows you to register and control various"
         ServSay SockVal, "ChanServ", "aspects of channels. Chanserv can help prevent malicious"
         ServSay SockVal, "ChanServ", "users from taking over a channel by limiting who"
         ServSay SockVal, "ChanServ", "is allowed channel operator priveleges. Available"
         ServSay SockVal, "ChanServ", "commands are listed below; to use them, type"
         ServSay SockVal, "ChanServ", Chr$(2) + "/msg ChanServ command" + Chr$(2) + ". For more information on a"
         ServSay SockVal, "ChanServ", "specific command, type " + Chr$(2) + "/msg ChanServ HELP command" + Chr$(2) + "."
         ServSay SockVal, "ChanServ", " "
         ServSay SockVal, "ChanServ", "    REGISTER   Register a channel"
         ServSay SockVal, "ChanServ", "    ACCESS     Modify the list of privileged users"
         ServSay SockVal, "ChanServ", "    OP         Give op status to a selected nick on a channel"
         ServSay SockVal, "ChanServ", "    DEOP       Deops a selected nick on a channel"
         ServSay SockVal, "ChanServ", "    LIST       Lists all registered channels matching the given pattern"
         ServSay SockVal, "ChanServ", "    INFO       Lists information about the named registered channel"
         ServSay SockVal, "ChanServ", "    DROP       Cancel the registration of a channel"
         ServSay SockVal, "ChanServ", " "
         ServSay SockVal, "ChanServ", "Note that any channel which is not used for 14 days"
         ServSay SockVal, "ChanServ", "(i.e. which no user on the channel's access list enters"
         ServSay SockVal, "ChanServ", "for that period of time) will be automatically dropped."
       Case "register"
         ServSay SockVal, "ChanServ", "ChanServ REGISTER syntax:"
         ServSay SockVal, "ChanServ", "    " + Chr$(2) + "/msg NickServ REGISTER password" + Chr$(2)
       Case "access"
         ServSay SockVal, "ChanServ", "ChanServ access syntax:"
         ServSay SockVal, "ChanServ", "    " + Chr$(2) + "/msg ChanServ ACCESS channel [ADD|LIST|DEL] {nickname} {level}" + Chr$(2)
         ServSay SockVal, "ChanServ", "Valid level values and their meanings listed below:"
         'ServSay SockVal, "ChanServ", "  1 - Automatically kick nickname from channel"
         ServSay SockVal, "ChanServ", "  2 - Automatically give nickname voice (+v) in channel"
         ServSay SockVal, "ChanServ", "  4 - Automatically give nickname op status (+o) in channel"
         ServSay SockVal, "ChanServ", "  12 - Co-founder. Gives nickname permission to modify the channel access list"
         
       Case "op"
          ServSay SockVal, "ChanServ", "ChanServ OP syntax:"
          ServSay SockVal, "ChanServ", "    " + Chr$(2) + "/msg ChanServ OP {nickname}" + Chr$(2)
       Case "deop"
          ServSay SockVal, "ChanServ", "ChanServ DEOP syntax:"
          ServSay SockVal, "ChanServ", "    " + Chr$(2) + "/msg ChanServ DEOP {nickname}" + Chr$(2)
       'Case "list"
          'ServSay SockVal, "ChanServ", "ChanServ LIST syntax:"
          'ServSay SockVal, "ChanServ", "    " + Chr$(2) + "/msg ChanServ REGISTER password" + Chr$(2)
       Case "info"
          ServSay SockVal, "ChanServ", "ChanServ INFO syntax:"
          ServSay SockVal, "ChanServ", "    " + Chr$(2) + "/msg ChanServ INFO channel" + Chr$(2)
       Case "drop"
          ServSay SockVal, "ChanServ", "ChanServ DROP syntax:"
          ServSay SockVal, "ChanServ", "    " + Chr$(2) + "/msg NickServ DROP" + Chr$(2)
          ServSay SockVal, "ChanServ", "Can only be used by channel founder, who must be identified"
          ServSay SockVal, "ChanServ", "with NickServ first."
       Case Else
         ServSay SockVal, "ChanServ", "Unknown command - " + Chr$(2) + args(1) + Chr$(2)
    End Select
  
  Case "register"
    If Len(args(2)) = 0 Then ServSay SockVal, "ChanServ", "I need more parameters.": Exit Sub
    newchan = ChanExists(args(1))
    If newchan = 0 Then ServSay SockVal, "ChanServ", "Channel " + Chr$(2) + args(1) + Chr$(2) + " does not exist.": Exit Sub
    If IsUserMode(SockVal, "r") = 0 Then ServSay SockVal, "ChanServ", "Sorry, you must be registered and identified with NickServ to perform this command.": Exit Sub
    If IsChanReg(args(1)) = 0 Then
      If IsUserChanMode(SockVal, newchan, "qo") = 0 Then
        ServSay SockVal, "ChanServ", "You must have channel operator (+o) status to register that channel!"
      Else
        RegChan NickDB(Clients(SockVal).NickMatch).Nick, newchan, args(2), "N/A"
        ServSay SockVal, "ChanServ", "You have registered " + Chr$(2) + args(1) + Chr$(2) + " and have the status of channel founder."
        ChanMaint newchan
      End If
    Else
      ServSay SockVal, "ChanServ", "Channel " + Chr$(2) + args(1) + Chr$(2) + " has already been registered."
    End If
    ServMaint
    
  Case "access"
    If Len(args(2)) = 0 Then ServSay SockVal, "ChanServ", "I need more parameters.": Exit Sub
    newchan = ChanExists(args(1))
    'If newchan = 0 Then ServSay SockVal, "ChanServ", "Channel " + Chr$(2) + args(1) + Chr$(2) + " does not exist.": Exit Sub
    IndexDB = IsChanReg(args(1))
    If IndexDB = 0 Then ServSay SockVal, "ChanServ", "Channel " + Chr$(2) + args(1) + Chr$(2) + " is not registered with ChanServ.": Exit Sub
    Select Case LCase$(args(2))
      Case "add"
        If Len(args(4)) = 0 Then ServSay SockVal, "ChanServ", "I need more parameters.": Exit Sub
        For n = 1 To ChanDB(IndexDB).EntryCount
          If LCase$(ChanDB(IndexDB).Entry(n).Nick) = LCase$(Clients(SockVal).Nick) Then
            mylevel = ChanDB(IndexDB).Entry(n).Level
            If mylevel < 12 Or IsUserMode(SockVal, "r") = 0 Then ServSay SockVal, "ChanServ", "You do not have sufficient priveleges to perform this action, or you have not identified with NickServ.": Exit Sub
          End If
        Next n
        
        isvalid = 0
        For n = 1 To NickDBCount
          If LCase$(NickDB(n).Nick) = LCase$(args(3)) Then isvalid = 1: Exit For
        Next n
        If isvalid = 0 Then ServSay SockVal, "ChanServ", "The nickname " + Chr$(2) + args(3) + Chr$(2) + " is not registered with NickServ.": Exit Sub
        modifying = 0
        For n = 1 To ChanDB(IndexDB).EntryCount
          If LCase$(args(3)) = LCase$(ChanDB(IndexDB).Entry(n).Nick) Then modifying = n: Exit For
        Next n
        Select Case Val(args(4))
          'Case 1 'AKICK
          '  newlevel = AKICK
          Case 2 'AVOICE
            newlevel = AVOICE
          Case 4 'AOP
            newlevel = AOP
          Case 12
            newlevel = COFOUNDER + AOP
          Case Else
            ServSay SockVal, "ChanServ", "Invalid access level parameter.": Exit Sub
        End Select
        If modifying = 0 Then
          NewChanDBEntry Int(IndexDB), args(3), Int(newlevel)
        Else
          oldlevel = ChanDB(IndexDB).Entry(modifying).Level
          If oldlevel >= mylevel Then
            ServSay SockVal, "ChanServ", "You are not allowed to modify the access level of somebody equal to or greater than your own."
            Exit Sub
          Else
            If oldlevel >= FOUNDER Then ServSay SockVal, "ChanServ", "The access level of the channel's founder can not be modified!": Exit Sub
            ChanDB(IndexDB).Entry(modifying).Level = newlevel
          End If
        End If
        ServSay SockVal, "ChanServ", "The nickname " + Chr$(2) + args(3) + Chr$(2) + " has been granted level" + Str$(newlevel) + " on channel " + Chr$(2) + args(1) + Chr$(2) + "."
        ServMaint
        If newchan > 0 Then ChanMaint newchan
        
      Case "list"
        'If Len(args(2)) = 0 Then ServSay SockVal, "ChanServ", "I need more parameters.": Exit Sub
        'newchan = ChanExists(args(1))
        'If newchan = 0 Then ServSay SockVal, "ChanServ", "Channel " + Chr$(2) + args(1) + Chr$(2) + " does not exist.": Exit Sub
        IndexDB = IsChanReg(args(1))
        ServSay SockVal, "ChanServ", "Channel access list for " + Chr$(2) + args(1) + Chr$(2) + ":"
        For n = 1 To ChanDB(IndexDB).EntryCount
          ServSay SockVal, "ChanServ", Left$(Str$(ChanDB(IndexDB).Entry(n).Level) + Space$(8), 8) + "- " + ChanDB(IndexDB).Entry(n).Nick
        Next n
        ServSay SockVal, "ChanServ", "End of access list."
        
      Case "del"
        If Len(args(4)) = 0 Then ServSay SockVal, "ChanServ", "I need more parameters.": Exit Sub
        For n = 1 To ChanDB(IndexDB).EntryCount
          If LCase$(ChanDB(IndexDB).Entry(n).Nick) = LCase$(Clients(SockVal).Nick) Then
            mylevel = ChanDB(IndexDB).Entry(n).Level
            If mylevel < COFOUNDER Or IsUserMode(SockVal, "r") = 0 Then ServSay SockVal, "ChanServ", "You do not have sufficient priveleges to perform this action, or you have not identified with NickServ.": Exit Sub
          End If
        Next n
        
        isvalid = 0
        For n = 1 To NickDBCount
          If LCase$(NickDB(n).Nick) = LCase$(args(3)) Then isvalid = 1: Exit For
        Next n
        If isvalid = 0 Then ServSay SockVal, "ChanServ", "The nickname " + Chr$(2) + args(3) + Chr$(2) + " is not registered with NickServ.": Exit Sub
        modifying = 0
        For n = 1 To ChanDB(IndexDB).EntryCount
          If LCase$(args(3)) = LCase$(ChanDB(IndexDB).Entry(n).Nick) Then modifying = n: Exit For
        Next n
        If modifying = 0 Then ServSay SockVal, "ChanServ", "There is no such nickname currently in that channel's access list.": Exit Sub
        If modifying = 1 Then ServSay SockVal, "ChanServ", "The channel founder cannot be removed from the access list. They must cancel it's registation."
        For n = modifying + 1 To ChanDB(IndexDB).EntryCount
          ChanDB(IndexDB).Entry(n - 1).Nick = ChanDB(IndexDB).Entry(n).Nick
          ChanDB(IndexDB).Entry(n - 1).Level = ChanDB(IndexDB).Entry(n).Level
        Next n
        ChanDB(IndexDB).Entry(ChanDB(IndexDB).EntryCount).Nick = ""
        ChanDB(IndexDB).Entry(ChanDB(IndexDB).EntryCount).Level = 0
        ChanDB(IndexDB).EntryCount = ChanDB(IndexDB).EntryCount - 1
        ServMaint
        
      Case Else
        ServSay SockVal, "ChanServ", "Unknown parameter - " + args(1)
    End Select
    
  Case Else
  ServSay SockVal, "ChanServ", "Unknown command - " + Chr$(2) + args(0) + Chr$(2) + ". " + Chr$(34) + "/msg ChanServ HELP" + Chr$(34) + " for help."
End Select
End Sub

Public Sub ChanMaint(ChanVal As Integer)
Dim n As Integer
Dim IndexDB As Integer
IndexDB = Chans(ChanVal).IndexDB
'MsgBox "but here wtf"
If IndexDB = 0 Then Exit Sub

If ChanDB(IndexDB).EnforceModes = 1 Then
  For n = 1 To Len(Chans(ChanVal).Modes)
    cc$ = Mid$(Chans(ChanVal).Modes, n, 1)
    If InStr(1, ChanDB(IndexDB).Modes, cc$) < 1 Then UnSetChanMode ChanVal, "ChanServ", cc$
  Next n
  For n = 1 To Len(ChanDB(IndexDB).Modes)
    cc$ = Mid$(ChanDB(IndexDB).Modes, n, 1)
    If InStr(1, Chans(IndexDB).Modes, cc$) < 1 Then SetChanMode ChanVal, "ChanServ", cc$
  Next n
End If

For n = 0 To frmMain.Sock.UBound
'MsgBox "yes it does"
  If n > frmMain.Sock.UBound Then Exit For
    If Clients(n).Chans(ChanVal) = 1 Then
      For nu = 1 To ChanDB(IndexDB).EntryCount
      hasentry = 0
        If LCase$(Clients(n).Nick) = LCase$(ChanDB(IndexDB).Entry(nu).Nick) Then
          hasentry = 1
          'if chandb(indexdb).Entry(nu).Level and akick = akick then
          'MsgBox Str$(ChanDB(IndexDB).Entry(nu).Level)
          If (ChanDB(IndexDB).Entry(nu).Level And AVOICE) = AVOICE Then
            If IsUserChanMode(n, ChanVal, "v") = 0 Then
              If IsUserMode(n, "r") Then SetUserChanMode ChanVal, "ChanServ", n, "v"
            End If
          End If
          'If (ChanDB(IndexDB).Entry(nu).Level And AOP) = AOP Then
          'MsgBox ChanDB(IndexDB).Entry(nu).Level
          If ChanDB(IndexDB).Entry(nu).Level >= 4 Then
            If IsUserMode(n, "r") <> 0 Then
              If IsUserChanMode(n, ChanVal, "o") = 0 Then SetUserChanMode ChanVal, "ChanServ", n, "o"
            Else
              UnSetUserChanMode ChanVal, "ChanServ", n, "o"
            End If
            
          Else
            If IsUserChanMode(n, ChanVal, "o") <> 0 Then UnSetUserChanMode ChanVal, "ChanServ", n, "o"
          End If
        End If
        'If hasentry = 0 Then UnSetUserChanMode ChanVal, "ChanServ", n, "o"
      Next nu
    End If
  'Next nu
Next n
End Sub

Public Function IsChanReg(ChanName As String)
For n = 1 To ChanDBCount
  If LCase$(ChanDB(n).Title) = LCase$(ChanName) Then IsChanReg = 1: Exit Function
Next n
IsChanReg = 0
End Function

Public Sub RegChan(Nick As String, ChanVal As Integer, Pass As String, Desc As String)
ChanDBCount = ChanDBCount + 1
ChanDB(ChanDBCount).EnforceModes = 0
ChanDB(ChanDBCount).FOUNDER = Nick
ChanDB(ChanDBCount).EntryCount = 1
ChanDB(ChanDBCount).Entry(1).Nick = Nick
ChanDB(ChanDBCount).Entry(1).Level = FOUNDER + AOP
ChanDB(ChanDBCount).Title = Chans(ChanVal).Title
ChanDB(ChanDBCount).LastUsed = UnixTimeEnc&(Date$, Timer)
ChanDB(ChanDBCount).Limit = Chans(ChanVal).ULimit
ChanDB(ChanDBCount).Modes = Chans(ChanVal).Modes
ChanDB(ChanDBCount).Pass = Hash.DigestStrToHexStr(Pass)
ChanDB(ChanDBCount).Description = Desc
ServMaint
End Sub

Public Sub NewChanDBEntry(IndexDB As Integer, Nick As String, AsLevel As Integer)
If ChanDB(IndexDB).EntryCount = 1000 Then Exit Sub
ChanDB(IndexDB).EntryCount = ChanDB(IndexDB).EntryCount + 1
ChanDB(IndexDB).Entry(ChanDB(IndexDB).EntryCount).Nick = Nick
ChanDB(IndexDB).Entry(ChanDB(IndexDB).EntryCount).Level = AsLevel
End Sub
