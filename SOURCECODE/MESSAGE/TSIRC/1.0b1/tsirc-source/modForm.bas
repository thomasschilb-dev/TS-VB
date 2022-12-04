Attribute VB_Name = "modForm"
Option Explicit
Public Const MAXCHAR = 20000

Public Function FindFreeIndex(hWndTemp As hWndType) As Integer
    Dim intArrayCount As Integer
    Dim lngCounter As Long
    
    If hWndTemp = Channel_hWnd Then
        intArrayCount = UBound(ChannelArray)
    ElseIf hWndTemp = Private_hWnd Then
        intArrayCount = UBound(PrivateArray)
    End If
    
    For lngCounter = 0 To intArrayCount
        If hWndTemp = Channel_hWnd Then
            If ChannelArrayFree(lngCounter) Then
                FindFreeIndex = lngCounter
                ChannelArrayFree(lngCounter) = False
                Exit Function
            End If
        ElseIf hWndTemp = Private_hWnd Then
            If PrivateArrayFree(lngCounter) Then
                FindFreeIndex = lngCounter
                PrivateArrayFree(lngCounter) = False
                Exit Function
            End If
        End If
        
    Next
    
    If hWndTemp = Channel_hWnd Then
        ReDim Preserve ChannelArray(intArrayCount + 1)
        ReDim Preserve ChannelArrayFree(intArrayCount + 1)
        FindFreeIndex = UBound(ChannelArray)
    ElseIf hWndTemp = Private_hWnd Then
        ReDim Preserve PrivateArray(intArrayCount + 1)
        ReDim Preserve PrivateArrayFree(intArrayCount + 1)
        FindFreeIndex = UBound(PrivateArray)
    End If
    
End Function

Public Sub CreateHwnd(strTag As String, hWndTemp As hWndType)
'This sub will create a new window, based on the free index return from sub FindFreeIndex
    Dim intIndex As Integer
    'If Left(strTag, 1) <> "#" Then Exit Sub
    intIndex = FindFreeIndex(hWndTemp)
    If hWndTemp = Private_hWnd Then
        PrivateArray(intIndex).Tag = strTag & "," & Val(intIndex)
        'PrivateArray(intIndex).Caption = strTag
        PrivateArray(intIndex).Show
    ElseIf hWndTemp = Channel_hWnd Then
        ChannelArray(intIndex).Tag = strTag & "," & Val(intIndex)
        'ChannelArray(intIndex).Caption = strMsg
        ChannelArray(intIndex).Show
    End If
End Sub

Public Sub CloseWindow(strTag As String)
    'This sub will close a given window (tag)
    Dim lngCounter As Long
    Dim strTemp As String
    Dim strCaption As String
    Dim intIndex As Integer
    For lngCounter = 0 To Forms.Count - 1
        strTemp = LCase(Forms(lngCounter).Tag)
        If Len(strTemp) <> 0 Then
            strCaption = LCase(Left(strTemp, InStr(strTemp, ",") - 1))
            intIndex = Val(Right(strTemp, Len(strTemp) - InStr(strTemp, ",")))
            
            If strCaption = LCase(strTag) Then
                ChannelArrayFree(intIndex) = True
                Unload Forms(lngCounter)
            End If
        End If
    Next
End Sub

Public Sub LogTextToHwnd(strTag As String, strMsg As String, hWndTemp As hWndType, strColor As String, blnNick As Boolean)
'This sub will loop through all the forms and log the text to certain textbox in the forms
    Dim lngCounter As Long
    Dim strCaption As String
    Dim strTemp As String
    Dim strNick As String
    Dim strText As String
    
    
    Dim intRed As Integer, intGreen As Integer, intBlue As Integer
    
    'If this window is already open, then goto SkipIt
    For lngCounter = 0 To Forms.Count - 1
        strTemp = LCase(Forms(lngCounter).Tag)
        If Len(strTemp) <> 0 Then
            strCaption = LCase(Left(strTemp, InStr(strTemp, ",") - 1))
            If strCaption = LCase(strTag) Then GoTo SkipIt
        End If
    Next
    
    'If the window is not found, then create a new window
    If hWndTemp = Channel_hWnd Then
        CreateHwnd strTag, Channel_hWnd
    ElseIf hWndTemp = Private_hWnd Then
        CreateHwnd strTag, Private_hWnd
    End If
    
SkipIt:
    
    'Log the text to the hWnd
    For lngCounter = 0 To Forms.Count - 1
        strTemp = LCase(Forms(lngCounter).Tag)
        If Len(strTemp) <> 0 Then
            strCaption = LCase(Left(strTemp, InStr(strTemp, ",") - 1))
            If strCaption = LCase(strTag) Then
                If blnNick = True Then
                    strNick = Trim(Left(strMsg, InStr(1, strMsg, " ") - 1))
                    strText = Trim(Right(strMsg, Len(strMsg) - InStr(1, strMsg, " ")))
                    
                    With udtColor
                        intRed = Val("&H" & Right(.colorNick, 2))
                        intGreen = Val("&H" & Mid(.colorNick, 3, 2))
                        intBlue = Val("&H" & Left(.colorNick, 2))
                    End With
                    
                    Forms(lngCounter).rtfDisplay.SelColor = RGB(intRed, intGreen, intBlue)
                    Forms(lngCounter).rtfDisplay.SelStart = Len(Forms(lngCounter).rtfDisplay.Text)
                    Forms(lngCounter).rtfDisplay.SelText = strNick
                    LogText Forms(lngCounter).rtfDisplay, " " & strText, strColor
                Else
                    LogText Forms(lngCounter).rtfDisplay, strMsg, strColor
                End If
            End If
        End If
    Next
End Sub


Public Sub ChangeCaption(strTag As String, strMsg As String)
'This sub is to change the caption of a window based on its tag
    Dim lngCounter As Long
    Dim strCaption As String
    Dim strTemp As String
    
    For lngCounter = 0 To Forms.Count - 1
        strTemp = LCase(Forms(lngCounter).Tag)
        If Len(strTemp) <> 0 Then
            strCaption = LCase(Left(strTemp, InStr(strTemp, ",") - 1))
            If strCaption = LCase(strTag) Then
                Forms(lngCounter).Caption = strMsg
            End If
        End If
    Next
End Sub

Public Sub GiveFocus(strTag As String)
    'This sub will set focus on the given hWnd
    Dim lngCounter As Long
    Dim strCaption As String
    Dim strTemp As String
    
    For lngCounter = 0 To Forms.Count - 1
        strTemp = LCase(Forms(lngCounter).Tag)
        If Len(strTemp) <> 0 Then
            strCaption = LCase(Left(strTemp, InStr(strTemp, ",") - 1))
            If strCaption = LCase(strTag) Then
                Forms(lngCounter).SetFocus
            End If
        End If
    Next
End Sub

Public Function GetCaption(strTag As String) As String
'This sub is to get the caption of a window based on it's tag
    Dim lngCounter As Long
    Dim strCaption As String
    Dim strTemp As String
    
    For lngCounter = 0 To Forms.Count - 1
        strTemp = LCase(Forms(lngCounter).Tag)
        If Len(strTemp) <> 0 Then
            strCaption = LCase(Left(strTemp, InStr(strTemp, ",") - 1))
            If strCaption = LCase(strTag) Then
                GetCaption = Forms(lngCounter).Caption
                Exit Function
            End If
        End If
    Next
End Function

Public Function GetMode(strTag As String) As String
'this function will loop throught the forms with the tag, get all the characters
'from the lstMode and return it as string
    Dim lngCounter As Long
    Dim intCounter As Long
    Dim strTemp As String
    Dim strChannel As String
    Dim strMode As String
    
    For lngCounter = 0 To Forms.Count - 1
        strTemp = Forms(lngCounter).Tag
        If Len(strTemp) <> 0 Then
            strChannel = LCase(Left(strTemp, InStr(strTemp, ",") - 1))
            If LCase(strChannel) = LCase(strTag) Then
                For intCounter = 0 To Forms(lngCounter).lstMode.ListCount - 1
                    strMode = strMode & Forms(lngCounter).lstMode.List(intCounter)
                Next
                GetMode = "[" & strMode & "]"
                Exit Function
            End If
        End If
    Next
End Function

Public Sub SaveMode(strTag As String, strMode As String)
'this sub will loop through the forms with the tag, then save every characters
'of the mode to the lstMode.  It's also trigger if + then add whatever after +
'or if - then it will remove whatever from the list.  I shoulda used array, but
'it wasn't visible for testing so i just threw in the listbox instead. ;p

    Dim lngCounter As Long
    Dim strTemp As String
    Dim strChannel As String
    Dim blnAdd As Boolean
    
    Dim intCounter As Long
    Dim strChar As String
    
    For lngCounter = 0 To Forms.Count - 1
        strTemp = Forms(lngCounter).Tag
        If Len(strTemp) <> 0 Then
            strChannel = LCase(Left(strTemp, InStr(strTemp, ",") - 1))
            If LCase(strChannel) = LCase(strTag) Then
                For intCounter = 1 To Len(strMode)
                    If Mid(strMode, intCounter, 1) = "+" Then
                        strChar = strChar & Mid(strMode, intCounter, 1)
                        Forms(lngCounter).lstMode.AddItem strChar
                        blnAdd = True
                        Debug.Print "ADD " & strChar
                        strChar = ""
                    ElseIf Mid(strMode, intCounter, 1) = "-" Then
                        strChar = strChar & Mid(strMode, intCounter, 1)
                        blnAdd = False
                        Debug.Print "REMOVE " & strChar
                        strChar = ""
                    Else
                        If blnAdd = True Then
                            strChar = strChar & Mid(strMode, intCounter, 1)
                            Forms(lngCounter).lstMode.AddItem strChar
                            strChar = ""
                        Else
                            strChar = strChar & Mid(strMode, intCounter, 1)
                            RemoveItem Forms(lngCounter).lstMode, strChar
                            strChar = ""
                        End If
                    End If
                Next
                'Clear the lstMode after set the mode
                If Forms(lngCounter).lstMode.ListCount = 1 Then Forms(lngCounter).lstMode.Clear
            End If
        End If
    Next
End Sub
Public Function GetChannel(strCaption As String) As String
    'this sub will return whatever after the colon, if there is no colon, return nothing
    
    If InStr(strCaption, ":") Then
        GetChannel = Trim(Right(strCaption, Len(strCaption) - InStr(strCaption, ":")))
    Else
        GetChannel = ""
    End If
End Function

Public Sub LogText(rtfBox As RichTextBox, strData As String, strColor As String)
    'Sub logtext to given richtextbox with hex color
        Dim strTemp As String
        Dim intRed As Integer, intGreen As Integer, intBlue As Integer
        
        strTemp = strColor
        'Parse the red, green and blue back to rgb
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        
        
        
        With rtfBox
            If Len(strData) + Len(.Text) > MAXCHAR Then
            'Scroll some text off the top to make more room
            .Text = Mid$(.Text, InStr(100 + Len(strData), .Text, vbCrLf) + 2)
            End If
            .SelStart = Len(.Text)
            .SelColor = RGB(intRed, intGreen, intBlue)
            .SelText = strData & vbCrLf
            .SelStart = Len(.Text)
        End With
End Sub

Public Sub GetList(strChannel As String, strData As String)
    'return a big string of all users in the channel, parsing and add to the list
    'based on the tag
    Dim lngCounter As Long
    Dim intCounter As Long
    Dim strTemp As String
    Dim strTag As String
    Dim strFirst As String
    
    strData = strData & " " 'Add extra space at the end so it will add the last word because of this parsing technique
    
    For intCounter = 0 To Forms.Count - 1
        strTag = LCase(Forms(intCounter).Tag)
        If Len(strTag) <> 0 Then
            strFirst = Left(strTag, InStr(strTag, ",") - 1)
            If strFirst = LCase(strChannel) Then
                For lngCounter = 1 To Len(strData)
                    If Mid(strData, lngCounter, 1) <> " " Then
                        strTemp = strTemp & Mid(strData, lngCounter, 1)
                    Else
                        Forms(intCounter).lstNick.AddItem strTemp
                        strTemp = ""
                    End If
                Next
                              
            End If
        End If
    Next
End Sub

Public Sub AddToList(strChannel As String, strNick As String)
'this sub will add a user to the listbox based on the tag
    Dim lngCounter As Long
    Dim strTemp As String
    Dim strTag As String
    
    For lngCounter = 0 To Forms.Count - 1
        strTag = LCase(Forms(lngCounter).Tag)
        If Len(strTag) <> 0 Then
            strTemp = Left(strTag, InStr(strTag, ",") - 1)
            If strTemp = LCase(strChannel) Then
                Forms(lngCounter).lstNick.AddItem strNick
            End If
        End If
    Next
    
End Sub

Public Sub RemoveName(strChannel As String, strNick As String)
'this sub will remove a user from the listbox based on the tag
    Dim lngCounter As Long
    Dim intCounter As Integer
    Dim strTemp As String
    Dim strTag As String

    For lngCounter = 0 To Forms.Count - 1
        strTag = LCase(Forms(lngCounter).Tag)
        If Len(strTag) <> 0 Then
            strTemp = Left(strTag, InStr(strTag, ",") - 1)
            If strTemp = LCase(strChannel) Then
                RemoveItem Forms(lngCounter).lstNick, strNick
            End If
        End If
    Next
    
End Sub


Public Sub RemoveNameAll(strNick As String, strMsg As String)
'This sub will remove names that quit by error and no given channel, so it will log
'to all channels that have that user

    Dim lngCounter As Long
    Dim intCounter As Integer
    Dim strTemp As String
    Dim strTag As String
    
    For lngCounter = 0 To Forms.Count - 1
        strTag = LCase(Forms(lngCounter).Tag)
        If Len(strTag) <> 0 Then
            If Left(strTag, 1) = "#" Then
                If SearchList(Forms(lngCounter).lstNick, strNick) Then
                    strTemp = LCase(Forms(lngCounter).Tag)
                    strTag = Left(strTemp, InStr(strTemp, ",") - 1)
                    RemoveItem Forms(lngCounter).lstNick, strNick
                    LogText Forms(lngCounter).rtfDisplay, "*** " & strNick & " has quit IRC (" & strMsg & ")", udtColor.colorQuit
                End If
            End If
        End If
    Next
    
End Sub

Public Sub LogToAll(strNick As String, strMsg As String, intNum As Integer)
    'log text to all windows that have the nick name
    '1 is notice
    '2 is change name
    '3 is quit
    
    Dim lngCounter As Long
    Dim intCounter As Integer
    Dim strTemp As String
    Dim strTag As String
    Dim strDummy As String          'add @ or + if the nick is @ or +
    Dim strDummy1 As String         'store the new nick name
    
    For lngCounter = 0 To Forms.Count - 1
        strTag = LCase(Forms(lngCounter).Tag)
        If Len(strTag) <> 0 Then
            'if the user is changing name and has @
            If AllIsOp(strNick) Then
                strDummy = "@" & strNick    'hold the nick for search
                strDummy1 = "@" & strMsg
            ElseIf AllIsVoice(strNick) Then
                strDummy = "+" & strNick    'dummy hold the nick for search
                strDummy1 = "+" & strNick
            Else
                strDummy = strNick          'put original
                strDummy1 = strMsg
            End If
            If Left(strTag, 1) = "#" Then
                If SearchList(Forms(lngCounter).lstNick, strDummy) Then
                    strTemp = LCase(Forms(lngCounter).Tag)
                    strTag = Left(strTemp, InStr(strTemp, ",") - 1)
                    'NOTICE
                    If intNum = 1 Then
                        LogText Forms(lngCounter).rtfDisplay, "-" & strNick & "- " & strMsg, udtColor.colorNotice
                    'CHANGE NICK
                    ElseIf intNum = 2 Then
                    
                        RemoveItem Forms(lngCounter).lstNick, strDummy
                        Forms(lngCounter).lstNick.AddItem strDummy1
                    
                        If LCase(strNick) = LCase(MyNick) Then
                            MyNick = strMsg
                            LogText Forms(lngCounter).rtfDisplay, "*** Your nick is now " & strMsg, udtColor.colorNick
                        Else
                            LogText Forms(lngCounter).rtfDisplay, "*** " & strNick & " is now known as " & strMsg, udtColor.colorNick
                        End If
                    'QUIT
                    ElseIf intNum = 3 Then
                        RemoveItem Forms(lngCounter).lstNick, strDummy
                        LogText Forms(lngCounter).rtfDisplay, "*** " & strNick & " has quit IRC (" & strMsg & ")", udtColor.colorQuit
                    End If
                End If
            End If
        End If
    Next
End Sub

Public Sub RemoveItem(lstBox As ListBox, strItem As String)
'Generic remove item sub
    Dim intCounter As Integer
    For intCounter = 0 To lstBox.ListCount - 1
        If LCase(lstBox.List(intCounter)) = LCase(strItem) Then lstBox.RemoveItem intCounter
    Next
End Sub

Public Function FindNameAll(strChannel As String, strItem As String) As Boolean
'This will search for a name from a list based on the tag, return true/false
    Dim lngCounter As Long
    Dim intCounter As Integer
    Dim strTemp As String
    Dim strTag As String
    
    For lngCounter = 0 To Forms.Count - 1
        strTag = LCase(Forms(lngCounter).Tag)
        If Len(strTag) <> 0 Then
            strTemp = LCase(Left(strTag, InStr(strTag, ",") - 1))
            If strTemp = LCase(strChannel) Then
                For intCounter = 0 To Forms(lngCounter).lstNick.ListCount - 1
                    If LCase(Forms(lngCounter).lstNick.List(intCounter)) = LCase(strItem) Then FindNameAll = True: Exit Function
                Next
            End If
        End If
    Next
    FindNameAll = False
End Function

Public Function SearchList(lstBox As ListBox, strItem As String) As Boolean
'This is generic search item in a list, return true/false
    Dim intCounter As Integer
    
    For intCounter = 0 To lstBox.ListCount - 1
        If LCase(lstBox.List(intCounter)) = LCase(strItem) Then SearchList = True: Exit Function
    Next
    SearchList = False
End Function

Public Sub DupeCheck(lstBox As ListBox)
'Check for duplication and remove it
    Dim lngCounter As Long
    For lngCounter = 0 To lstBox.ListCount - 1
        If lstBox.List(lngCounter) = lstBox.List(lngCounter - 1) Then lstBox.RemoveItem lngCounter
    Next
End Sub
