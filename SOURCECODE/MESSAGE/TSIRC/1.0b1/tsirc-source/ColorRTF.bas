Attribute VB_Name = "ColorRTF"
Option Explicit

Public Sub DoColor(RTF As RichTextBox, a As String)
    Dim b As String
    Dim n As Integer
    Dim n2 As Integer
    Dim fgcolor As Integer
    Dim bgcolor As Integer
    Dim savefg As Integer
    Dim savebg As Integer
    Dim bReverse As Boolean
    
    Dim Color(0 To 15) As Long
    Color(0) = vbWhite 'white
    Color(1) = vbBlack 'black
    Color(2) = RGB(0, 0, 140) 'dark blue
    Color(3) = RGB(0, 140, 0) 'dark green
    Color(4) = vbRed 'red
    Color(5) = RGB(110, 65, 0) 'brown
    Color(6) = RGB(140, 0, 140) 'purple
    Color(7) = RGB(248, 146, 0) 'orange
    Color(8) = RGB(255, 255, 0) 'yellow
    Color(9) = vbGreen 'light green
    Color(10) = RGB(0, 140, 140) 'dark blue green
    Color(11) = RGB(0, 255, 255) 'light blue green
    Color(12) = vbBlue 'light blue
    Color(13) = vbMagenta 'magenta
    Color(14) = RGB(140, 140, 140) 'grey
    Color(15) = RGB(200, 200, 200) 'light grey

RTF.FontBold = False
RTF.FontUnderline = False

fgcolor = 1
bgcolor = 0
bReverse = False
savefg = fgcolor
savebg = bgcolor

For n = 1 To Len(a)
   b = Mid(a, n, 1)
   If b = Chr(3) Then
    'Parse Colours
    If IsNumeric(Mid(a, n + 1, 1)) Then
       If IsNumeric(Mid(a, n + 2, 1)) Then
        If Mid(a, n + 3, 1) = "," Then
            If IsNumeric(Mid(a, n + 4, 1)) Then
                If IsNumeric(Mid(a, n + 5, 1)) Then
                    '@##,##
                    fgcolor = CInt(Mid(a, n + 1, 2))
                    bgcolor = CInt(Mid(a, n + 4, 2))
                    n = n + 5
                Else
                    '@##,#
                    fgcolor = CInt(Mid(a, n + 1, 2))
                    bgcolor = CInt(Mid(a, n + 4, 1))
                    n = n + 4
                End If
            Else
                '@##,
                fgcolor = CInt(Mid(a, n + 1, 2))
                n = n + 3
            End If
        Else
            '@##
            fgcolor = CInt(Mid(a, n + 1, 2))
            n = n + 2
        End If
           ElseIf Mid(a, n + 2, 1) = "," Then
        If IsNumeric(Mid(a, n + 3, 1)) Then
            If IsNumeric(Mid(a, n + 4, 1)) Then
                '@#,##
                fgcolor = CInt(Mid(a, n + 1, 1))
                'bgcolor = CInt(Mid(a, n + 3, 2))
                n = n + 4
            Else
                '@#,#
                fgcolor = CInt(Mid(a, n + 1, 1))
                bgcolor = CInt(Mid(a, n + 3, 1))
                n = n + 3
            End If
        Else
            '@#,
            fgcolor = CInt(Mid(a, n + 1, 1))
            n = n + 2
        End If
           Else
        '@#
        fgcolor = CInt(Mid(a, n + 1, 1))
        n = n + 1
       End If
       If fgcolor > 15 Then
           fgcolor = 1
       End If

       If bgcolor > 15 Then
           bgcolor = 0
       End If
       RTF.FontColour = Color(fgcolor)
       RTF.FontBackColour = Color(bgcolor)
        Else
           RTF.FontColour = Color(1)
           RTF.FontBackColour = Color(0)
        End If
   ElseIf b = Chr(2) Then
    RTF.FontBold = Not (RTF.FontBold)
'   if bBold then'
'       'Turn Bold off
'       bBold = False
'       RTF.FontBold = False
'   else
'       'Turn Bold on
'       bBold = True
'       RTF.FontBold = True
'   endif
   ElseIf b = Chr(31) Then
        RTF.FontUnderline = Not (RTF.FontUnderline)
'   if bUnderline then
'       'Turn underline off
'       bUnderline = False
'       RTF.FontUnderline = False
'   else
'       'Turn underline on
'       bUnderline = True
'       RTF.FontUnderline = True
'   endif
   ElseIf b = Chr(22) Then
    'Reverse Foreground / Background colors
'    n2 = bgcolor
'    bgcolor = fgcolor
'    fgcolor = n2
'    RTF.FontColour = color(fgcolor)
'    RTF.FontBackColour = color(bgcolor)
    'Set the colors to the reverse standard colour set.
    If bReverse Then
        bReverse = False
        fgcolor = savefg
        bgcolor = savebg
    Else
        bReverse = True
        savefg = fgcolor
        savebg = bgcolor
        fgcolor = 0
        bgcolor = 1
        
    End If
    RTF.FontColour = Color(fgcolor)
    RTF.FontBackColour = Color(bgcolor)
    
   Else
    RTF.InsertContents SF_TEXT, b
   End If
Next n

    
'    For i = 1 To Len(strColor)
'        RTF.InsertContents SF_TEXT, Asc(Mid(strColor, i, 1)) & "|"
'    Next i
    
    RTF.FontColour = Color(1)
    RTF.FontBackColour = Color(0)
    RTF.FontBold = False
    RTF.FontUnderline = False

'    RTF.InsertContents SF_TEXT, vbCrLf
End Sub


