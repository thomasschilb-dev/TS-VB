VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmSocket 
   ClientHeight    =   435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2355
   LinkTopic       =   "Form1"
   ScaleHeight     =   435
   ScaleWidth      =   2355
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrParse 
      Interval        =   1
      Left            =   1920
      Top             =   0
   End
   Begin VB.Timer tmrSend 
      Interval        =   1
      Left            =   1440
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   990
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock socketDCC 
      Left            =   480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   -15
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Socket_Connect()
    frmStatus.rtfStatus.SelText = "Connected to server!" & vbCrLf
    Socket.SendData "User " & Email & " " & Socket.LocalHostName & " " & Socket.RemoteHost & " :" & RealName & vbCrLf
    Socket.SendData "NICK " & MyNick & vbCrLf
    frmStatus.Caption = "Status: " & NickName
End Sub

Private Sub Socket_ConnectionRequest(ByVal requestID As Long)
    'Accept request ID
    Socket.Close
    Socket.Accept requestID
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Socket.GetData strData, vbString
    'Check for data and add to queue
    CheckForLine strData
End Sub

Private Sub tmrParse_Timer()
    'The timer will check for queue message every 1 millisecond.  This is secondary parse
    'process in this program.  It will parse out host, trigger, message
    
    Dim intCount As Integer
    Dim blnParsed As Boolean
    Dim strData As String
    
    Dim strFirst As String, strSecond As String, strThird As String, strFourth As String
    Dim intPos1 As Integer, intPos2 As Integer, intPos3 As Integer, intPos4 As Integer
    
    
    intCount = 1
    Do While blnParsed = False And intCount <= QueueMsg.Count
        
        strData = QueueMsg.Item(intCount)
        
        'remove first line feed if there are any
        If Mid(strData, 1, 1) = Chr(13) Or Mid(strData, 1, 1) = Chr(10) Then
            strData = Mid(strData, 2)
        End If
        
        intPos1 = InStr(1, strData, " ")
        If intPos1 Then
            strFirst = Trim(Left(strData, intPos1))
            intPos2 = InStr(intPos1 + 1, strData, " ")
            If intPos2 Then
                strSecond = Trim(Mid(strData, intPos1 + 1, (intPos2 - intPos1)))
                intPos3 = InStr(intPos2 + 1, strData, " ")
                    If intPos3 Then
                        strThird = Trim(Mid(strData, intPos2 + 1, (intPos3 - intPos2)))
                        strFourth = Trim(Right(strData, Len(strData) - intPos3))
                    Else    'no third space
                        strThird = Trim(Mid(strData, intPos2 + 1, Len(strData) - intPos2))
                    End If
            Else    'no second space, mostlikely PING or ERROR
                strFirst = Trim(Right(strData, Len(strData) - InStr(strData, ":")))
                strSecond = "PING"
                strThird = ""
                strFourth = ""
            End If
        End If
        
        'Error case
        If UCase(strFirst) = "ERROR" Then
            strFirst = ""
            strSecond = "ERROR"
        End If
        ParseMsg strFirst, strSecond, strThird, strFourth
        
        blnParsed = True
        QueueMsg.Remove intCount
        intCount = intCount + 1
    Loop
End Sub

