Attribute VB_Name = "modColor"
Option Explicit

Public Type udtColorVariables
    colorAction     As String * 6
    colorCTCP       As String * 6
    colorJoin       As String * 6
    colorPart       As String * 6
    colorKick       As String * 6
    colorQuit       As String * 6
    colorMode       As String * 6
    colorNotice     As String * 6
    colorOwn        As String * 6
    colorNick       As String * 6
    colorUser       As String * 6
    colorInvite     As String * 6
    colorTopic      As String * 6
    colorWhois      As String * 6
    colorChat       As String * 6
    colorOther      As String * 6
    colorListText   As String * 6
    colorEditText   As String * 6
    colorEdit       As String * 6
    colorFrame      As String * 6
    colorList       As String * 6
End Type


Public udtColor As udtColorVariables

Public Sub LoadColor()
    Dim strApp As String
    Dim intFile As Integer
    intFile = FreeFile
    ''strApp = App.Path & "\tsIRC.ini"
    strApp = "tsIRC.ini"
    
    'Load color from file
    Open strApp For Random Access Read As #intFile Len = Len(udtColor)
        Get #intFile, , udtColor
    Close #intFile
    
    'If there are no color setting, then give default color
    With udtColor
        If .colorAction = String(6, Chr(0)) Then
            .colorAction = "A22866"
            .colorCTCP = "00008F"
            .colorJoin = "0070B8"
            .colorPart = "99C2A3"
            .colorKick = "995B00"
            .colorQuit = "8F001E"
            .colorMode = "007033"
            .colorNotice = "00008F"
            .colorOwn = "2899B8"
            .colorNick = "999496"
            .colorUser = "2899B8"
            .colorInvite = "998500"
            .colorTopic = "998500"
            .colorWhois = "999496"
            .colorChat = "FFFFFF"
            .colorOther = "FFFFFF"
            .colorListText = "FFFFFF"
            .colorEditText = "FFFFFF"
            .colorEdit = "000000"
            .colorFrame = "000000"
            .colorList = "000000"
        End If
    End With
    
End Sub

Public Sub SaveColor()
    'Save current color setting to file
    Dim strApp As String
    Dim intFile As Integer
    intFile = FreeFile
    ''strApp = App.Path & "\tsIRC.ini"
    strApp = "tsIRC.ini"
    
    Open strApp For Random Access Write As #intFile Len = Len(udtColor)
        Put #intFile, , udtColor
    Close #intFile
End Sub
Public Function RGBtoHEX(RGB) As String
    'Convert rgb format to hex
    Dim strMsg As String
    Dim intCounter As Integer
    strMsg = Hex(RGB)
    intCounter = Len(strMsg)
    
    Select Case intCounter
        Case 1
            strMsg = String(5, "0") & strMsg
        Case 2
            strMsg = String(4, "0") & strMsg
        Case 3
            strMsg = String(3, "0") & strMsg
        Case 4
            strMsg = String(2, "0") & strMsg
        Case 5
            strMsg = String(1, "0") & strMsg
    End Select
    RGBtoHEX = strMsg
End Function

Public Sub ChangeObjectColor(ctrlObject As Control, strColor As String, intNum As Integer)
    'This sub is universal to change an object color.

    Dim intRed As Integer, intGreen As Integer, intBlue As Integer
    
    intRed = Val("&H" & Right(strColor, 2))
    intGreen = Val("&H" & Mid(strColor, 3, 2))
    intBlue = Val("&H" & Left(strColor, 2))
    
    Select Case intNum
        Case 1
            ctrlObject.BackColor = RGB(intRed, intGreen, intBlue)
        Case 2
            ctrlObject.ForeColor = RGB(intRed, intGreen, intBlue)
        Case 3
            ctrlObject.SelColor = RGB(intRed, intGreen, intBlue)
    End Select
End Sub
