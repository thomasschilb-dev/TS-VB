Attribute VB_Name = "Input"
Sub xINPUT(strData As String)
    strData = RTrim(strData)
    Dim x, y As Integer
    'don't set array it'll be set with ReDim
    Dim word() As String
    Dim parms As String
    
    'split the commands into seperate words
    'ReDim Preserve statement is the KEY
    If InStr(strData, Chr(32)) Then
        Do Until InStr(strData, Chr(32)) = 0
            x = InStr(strData, Chr(32))
            If x Then
                y = y + 1
                ReDim Preserve word(y)
                word(y) = Mid(strData, 1, x - 1)
                strData = Mid(strData, x + 1)
            End If
        Loop
        ReDim Preserve word(y + 1)
        word(y + 1) = strData
        
        'DO IT
        Select Case UCase(word(1))
            Case "RAW"
                For x = 2 To UBound(word)
                    parms = parms & word(x) & Chr(32)
                Next x
                parms = RTrim(parms)
                frmSocket.Socket.SendData parms
                frmStatus.rtfStatus.SelText = "->Server: " & parms
                frmStatus.rtfStatus.SelText = "-" & vbCrLf
            Case "MSG"
                For x = 3 To UBound(word)
                    parms = parms & word(x) & Chr(32)
                Next x
                parms = RTrim(parms)
                frmSocket.Socket.SendData "PRIVMSG " & word(2) & " :" & parms & vbCrLf
                'Adam Need to call a Docolor here...
                frmStatus.rtfStatus.SelColor = vbWhite
                frmStatus.rtfStatus.SelText = "-> *" & word(2) & "* " & parms & vbCrLf
                frmStatus.rtfStatus.SelText = "-" & vbCrLf
                frmStatus.rtfStatus.SelColor = vbWhite
            Case "WHOIS"
                frmSocket.Socket.SendData "WHOIS " & word(2) & vbCrLf
            Case "SERVER"
                ' Show server:port in status window
                frmStatus.Caption = "Status (" & word(2) & ":" & word(3) & ")"
                ' Connect to server port
                Connect word(2), word(3)
                frmStatus.rtfStatus.SelColor = vbWhite
                frmStatus.rtfStatus.SelText = " *** Connecting to Server " & vbCrLf
            Case "J"
                If Left(word(2), 1) = "#" Then
                    frmSocket.Socket.SendData "JOIN " & word(2) & vbCrLf
                Else
                    frmSocket.Socket.SendData "JOIN #" & word(2) & vbCrLf
                End If
            'Case "JOIN"
            '    If Left(word(2), 1) = "#" Then
            '        frmSocket.Socket.SendData "JOIN " & word(2) & vbCrLf
            '    Else
            '        frmSocket.Socket.SendData "JOIN #" & word(2) & vbCrLf
            '    End If
            Case "PART"
                If Left(word(2), 1) = "#" Then
                    frmSocket.Socket.SendData "PART " & word(2) & vbCrLf
                Else
                    frmSocket.Socket.SendData "PART #" & word(2) & vbCrLf
                End If
            Case "NICK"
                frmSocket.Socket.SendData "NICK " & ":" & word(2) & vbCrLf
        End Select
    Else
        'words that take no parameters just one word
        Select Case UCase(strData)
            Case "LIST"
                frmSocket.Socket.SendData "LIST" & vbCrLf
            Case "MOTD"
                frmSocket.Socket.SendData "MOTD" & vbCrLf
        End Select
    End If
End Sub

