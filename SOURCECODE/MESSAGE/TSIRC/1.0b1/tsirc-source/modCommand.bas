Attribute VB_Name = "modOperators"

Option Explicit
    
Public Channel(1 To 10) As New frmChannel
Public ChannelName(1 To 10) As String

Public Messages(1 To 100) As New frmPrivate
Public messagesName As String

Public server As String

Public Function AllIsOp(strNick As String) As Boolean
    Dim lngCounter As Long
    Dim intCounter As Integer
    Dim strTemp As String
    Dim strTag As String
    
    For lngCounter = 0 To Forms.Count - 1
        strTag = LCase(Forms(lngCounter).Tag)
        If Len(strTag) <> 0 Then
            If Left(strTag, 1) = "#" Then
                If SearchList(Forms(lngCounter).lstNick, "@" & strNick) Then
                    AllIsOp = True
                    Exit Function
                End If
            End If
        End If
    Next
    AllIsOp = False
End Function

Public Function AllIsVoice(strNick As String) As Boolean
    Dim lngCounter As Long
    Dim intCounter As Integer
    Dim strTemp As String
    Dim strTag As String
    
    For lngCounter = 0 To Forms.Count - 1
        strTag = LCase(Forms(lngCounter).Tag)
        If Len(strTag) <> 0 Then
            If Left(strTag, 1) = "#" Then
                If SearchList(Forms(lngCounter).lstNick, "+" & strNick) Then
                    AllIsVoice = True
                    Exit Function
                End If
            End If
        End If
    Next
    AllIsVoice = False
End Function

'Check for op in given channel
Public Function IsOp(strChannel As String, strNick As String) As Boolean
    If FindNameAll(strChannel, "@" & strNick) Then
        IsOp = True
    Else
        IsOp = False
    End If
End Function

'Check for voice in given channel
Public Function IsVoice(strChannel As String, strNick As String) As Boolean
    If FindNameAll(strChannel, "+" & strNick) = True Then
        IsVoice = True
    Else
        IsVoice = False
    End If
End Function

'Give regular user a +
Public Sub Voice(strChannel As String, strNick As String)
'only voice when no voiced or no opped
    If Not IsOp(strChannel, strNick) Then
        If Not IsVoice(strChannel, strNick) Then
            RemoveName strChannel, strNick
            AddToList strChannel, "+" & strNick
        End If
    End If
End Sub

'Take away the + from user
Public Sub DeVoice(strChannel As String, strNick As String)
'only devoice when only voiced
    If IsVoice(strChannel, strNick) Then
        RemoveName strChannel, "+" & strNick
        AddToList strChannel, strNick
    End If
End Sub

'Give regular/voice user a @
Public Sub Op(strChannel As String, strNick As String)
'if voiced, then remove voiced then add @
    If IsOp(strChannel, strNick) = False Then
        If IsVoice(strChannel, strNick) Then
            RemoveName strChannel, "+" & strNick
        Else
            RemoveName strChannel, strNick
        End If
        
            AddToList strChannel, "@" & strNick
    End If
End Sub

'Take away the @ from user
Public Sub DeOp(strChannel As String, strNick As String)
'only when opped
    If IsOp(strChannel, strNick) Then
        RemoveName strChannel, "@" & strNick
        AddToList strChannel, strNick
    End If
End Sub

Public Sub SetNickMode(strChannel As String, strReceiver As String, strMode As String)
    Dim strSign As String
    Dim strChar As String
    
    strSign = Left(strMode, 1)
    strChar = Right(strMode, 1)
    
    If strSign = "+" Then
        If LCase(strChar) = "o" Then
            Op strChannel, strReceiver
        ElseIf LCase(strChar) = "v" Then
            Voice strChannel, strReceiver
        ElseIf LCase(strChar) = "b" Then
            RemoveName LCase(strChannel), strReceiver
        End If
    ElseIf strSign = "-" Then
        If LCase(strChar) = "o" Then
            DeOp strChannel, strReceiver
        ElseIf LCase(strChar) = "v" Then
            DeVoice strChannel, strReceiver
        End If
    End If
End Sub

