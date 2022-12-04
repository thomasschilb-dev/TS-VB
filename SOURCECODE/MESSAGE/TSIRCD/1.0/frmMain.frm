VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "tsIRCd"
   ClientHeight    =   6645
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9165
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer MaintTimer 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer DieTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   0
   End
   Begin VB.Timer PingTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Sock 
      Index           =   0
      Left            =   360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckListen 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   6615
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DieCheck As Integer

Private Sub DieTimer_Timer()
DieCheck = DieCheck + 1
If DieCheck = 11 Then End
DieTimer.Enabled = False
alldead = 1
For n = 0 To Sock.UBound
  If n > Sock.UBound Then Exit For
  If Sock(n).State <> sckClosed Then alldead = 0: Exit For Else Sock(n).Close
Next n
If alldead = 1 Then End
DieTimer.Enabled = True
End Sub
'THIS MAKES THE MENU POPUP WHEN THE FORM IS HIDDEN IN THE SYSTRAY'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Sys As Long
Sys = x / Screen.TwipsPerPixelX
Select Case Sys
Case WM_RBUTTONDOWN:
End
'Me.PopupMenu mnuSystray
End Select
End Sub

'THIS MAKES THE FOR DISSAPEAR/MINIMIZE TO THE SYSTRAY'
Private Sub Form_Resize()
If WindowState = vbMinimized Then
Me.Hide
Me.Refresh
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = Me.Caption & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
Else
Shell_NotifyIcon NIM_DELETE, nid
End If
End Sub

'THIS WILL KILL THE SYSTRAY ICON IF THE FORM IS UNLOADED'
Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub
'THIS UNLOADS THE FORM FROM THE MENU'
Private Sub mnuexit_Click()
Unload Me
End Sub
'THIS RESTORES THE FORM'
Private Sub mnuRestore_Click()
WindowState = vbNormal
Me.Show
End Sub

Private Sub Form_Load()
frmMain.Visible = True
WindowState = vbMinimized

If UCase$(Command$) = "/DEBUG" Then frmMain.Visible = True

Open "c:\log.txt" For Output As #1
InitServer
End Sub

Private Sub MaintTimer_Timer()
ServMaint
End Sub

Private Sub PingTimer_Timer()
  PingTimer.Enabled = False
  curunix& = UnixTimeEnc&(Date$, Timer)
  For n = 0 To frmMain.Sock.UBound
    If n > frmMain.Sock.UBound Then Exit For 'because sometimes that value can change after the for loop begins! this avoids errors.
    If frmMain.Sock(n).State = sckConnected Then
      Sock(n).SendData ""
      If Clients(n).NickTimer > 0 And curunix& - Clients(n).NickTimer >= 60 Then ExpireNick Int(n)
      If curunix& - Clients(n).LastPing >= PingInterval Then DoPing Int(n)
      If curunix& - Clients(n).LastPing >= PingTimeout Then
        If Clients(n).GotPong = 0 Then
          QuitClient Int(n), ":Ping timeout:" + Str$(PingTimeout) + " seconds"
          ClearClientSlot Int(n)
        End If
      End If
    End If
    DoEvents
  Next n
  PingTimer.Enabled = True
End Sub

Private Sub sckListen_ConnectionRequest(Index As Integer, ByVal requestID As Long)
For n = 0 To Sock.UBound
  On Error Resume Next
  If Sock(n).State <> sckConnected Then Exit For
Next n

If n = 4097 Then Exit Sub
If n > Sock.UBound Then Load Sock(n)

Sock(n).Close
ClearClientSlot Int(n)

Sock(n).Accept requestID
Clients(n).ConnectTime = UnixTimeEnc&(Date$, Timer)
SM Int(n), ":" + ServerHost + " NOTICE AUTH :*** Looking up your hostname..."
SM Int(n), ":" + ServerHost + " NOTICE AUTH :*** Not checking ident"
If Len(Sock(n).RemoteHost) = 0 Then
  Clients(n).Host = Sock(n).RemoteHostIP
  SM Int(n), ":" + ServerHost + " NOTICE AUTH :*** Could not determine your hostname, so using your IP " + Clients(n).Host
Else
  Clients(n).Host = Sock(n).RemoteHost
  SM Int(n), ":" + ServerHost + " NOTICE AUTH :*** Your hostname is " + Clients(n).Host
End If
DoPing Int(n)
End Sub

Private Sub Sock_Close(Index As Integer)
QuitClient Index, "Client exited"
ClearClientSlot Index
End Sub

Private Sub Sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Sock(Index).GetData ib$
'If InStr(1, ib$, Chr$(10)) > 0 And Index = SvcSock Then MsgBox "10"
'If InStr(1, ib$, Chr$(13)) > 0 And Index = SvcSock Then MsgBox "13"

ib$ = ReplaceAll$(ib$, Chr$(13), Chr$(10))
ib$ = ReplaceAll$(ib$, Chr$(10) + Chr$(10), Chr$(10))
InBuffer(Index) = InBuffer(Index) + ib$
If InStr(1, InBuffer(Index), Chr$(10)) < 1 Then Exit Sub

Do Until Len(InBuffer(Index)) = 0
  xp = InStr(1, InBuffer(Index), Chr$(10))
  If xp < 1 Then xp = Len(curline$) + 1
  curline$ = Left$(InBuffer(Index), xp - 1)
  InBuffer(Index) = Mid$(InBuffer(Index), xp + 1)
  'If Index = SvcSock And chkDebug.value = vbChecked Then txtDebug = txtDebug + curline$ + vbCrLf: txtDebug.SelStart = Len(txtDebug) + 1
  ProcessInput Index, curline$
Loop
End Sub

Private Sub Sock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Sock(Index).State <> sckClosed Then Sock(Index).Close
QuitClient Index, "Connection reset by peer"
ClearClientSlot Index
End Sub

Private Sub Sock_SendComplete(Index As Integer)
'DoEvents
If KillFlag(Index) = 1 Then Sock(Index).Close: ClearClientSlot Index
If Len(Buffer(Index)) = 0 Or InStr(1, Buffer(Index), Chr$(13)) < 1 Then Exit Sub

xp = InStr(1, Buffer(Index), Chr$(13))
If xp < 1 Then xp = Len(Buffer(Index)) + 1
a$ = Left$(Buffer(Index), xp - 1)
Buffer(Index) = Mid$(Buffer(Index), xp + 1)
If Sock(Index).State = sckConnected And Len(LTrim$(RTrim$(a$))) > 0 Then
  Sock(Index).SendData a$ + vbCrLf
  If frmMain.Visible = True Then
    Text1 = Text1 + "--->" + a$ + vbCrLf
    frmMain.Text1.SelStart = Len(frmMain.Text1) + 1
  End If
End If
End Sub

