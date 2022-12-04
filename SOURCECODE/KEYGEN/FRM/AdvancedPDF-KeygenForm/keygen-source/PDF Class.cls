VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PDFClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PDFClass = "1.3"
Private Const PDFClassVersion = "PDFClass 1.0"

Private wsPathConfig As String
Private wsPathAdobe  As String

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal szClass$, ByVal szTitle$) As Long

Private Declare Function PDFReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function PDFCreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function PDFGetFileSize Lib "kernel32" Alias "GetFileSize" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function PDFCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long

Private Const WM_CLOSE = &H10

Private Type oOutlines
    sText      As String
    iLevel     As Integer
    yPos       As Double
    iPageNb    As Integer
    bPrev      As Boolean
    bNext      As Boolean
    bFirst     As Boolean
    bLast      As Boolean
    iFirst     As Integer
    iNext      As Integer
    iPrev      As Integer
    iLast      As Integer
    iParent    As Integer
End Type

Private aOutlines()         As oOutlines
Private iOutlines           As Integer
Private aPage()             As Variant

Private Type PDFRGB
    in_r       As Integer
    in_g       As Integer
    in_b       As Integer
End Type

Private Fso                 As Object
Private Strm                As Object
Private sPDFName            As String

Private Arr_Font()          As Variant

Private in_offset           As Integer
Private in_FontNum          As Integer
Private in_PagesNum         As Integer
Private in_Ech              As Double
Private in_Canvas           As Integer
Private iWidthStr           As Double

Private in_xCurrent         As Double
Private in_yCurrent         As Double

Private ImgWidth            As Double
Private ImgHeight           As Double

Private xlink               As Double
Private yLink               As Double
Private strTLink            As String
Private strTyLink           As String
Private wRect               As Long

Private str_TmpFont         As String

Private PDFTextColor        As String
Private PDFLineColor        As String
Private PDFDrawColor        As String

Private PDFstrTextColor     As String
Private PDFstrLineColor     As String
Private PDFstrDrawColor     As String
Private PDFstrTempColor     As String
Private PDFstrTempAlign     As String
Private PDFstrTempBorder    As String
Private pTempAngle          As Double
Private PDFboTempFill       As Boolean

Private bPageBreak          As Boolean

Private PDFLnStyle          As String
Private PDFLnWidth          As Double

Private PDFDrawMode         As String

Private PDFZoomMode
Private PDFLayoutMode
Private PDFViewerPref
Private bPDFViewerPref      As Boolean
Private bPDFWatermark        As Boolean
Private sPDFWatermark        As String

Private PDFAngle            As Double
Private bAngle              As Double

Private PDFFontName         As String
Private PDFFontSize         As Integer
Private PDFFontNum          As Integer

Private boPDFUnderline      As Boolean
Private boPDFItalic         As Boolean
Private boPDFBold           As Boolean
Private boPDFConfirm        As Boolean
Private boPDFView           As Boolean
Private PDFboThumbs         As Boolean
Private PDFboOutlines       As Boolean
Private PDFboImage          As Boolean

Private PDFlMargin          As Integer ' Left Margin
Private PDFtMargin          As Integer ' Top Margin
Private PDFrMargin          As Integer ' Right Margin
Private PDFbMargin          As Integer ' Bottom Margin
Private PDFcMargin          As Integer ' Center Margin
Private PDFMargin           As Integer

Private FFileName           As String
Private FTitle              As String
Private FPageNumber         As Integer
Private FPageLink           As Integer

Private FOrientation        As String
Private FAuthor             As String
Private FCreator            As String
Private FKeywords           As String
Private FSubject            As String
Private FProducer           As String
Private FFileCompress       As Boolean

Private ParentNum, ContentNum, ResourceNum, FontNum, CatalogNum, _
        FontNumber, CurrentPDFSetPageObject, NumberofImages, iOutlineRoot As Integer

Private PDFCanvasWidth()
Private PDFCanvasHeight()
Private PDFCanvasOrientation()

Private CurrentObjectNum    As Integer
Private ObjectOffset        As Long
Private ObjectOffsetList    As Variant
Private PageNumberList      As Variant
Private PageLinksList(1 To 1000, 1 To 1000) As Variant
Private LinksList           As Variant
Private PageCanvasWidth     As Variant
Private PageCanvasHeight    As Variant
Private FontNumberList      As Variant

Private Type aIMG
    in_1    As Variant
    in_2    As Variant
    in_3    As Variant
    in_4    As Variant
    in_5    As Variant
    in_6    As Variant
    in_7    As Variant
    in_8    As Variant
End Type

Private ArrIMG()            As aIMG

Private boPageLinksList     As Variant
Private NbPageLinksList     As Variant

Private CRCounter           As Long

Private ColorSpace          As String
Private ColorCount          As Byte
Private ImageStream         As String
Private TempStream          As String
Private pTempStream         As String
Private sTempStream         As String
Private cTempStream         As String
Private dTempStream         As String

Private StreamSize1, StreamSize2 As Integer

Private bScanAdobe          As Boolean

Enum PDFStyleLgn
    pPDF_SOLID = 0
    pPDF_DASH = 1
    pPDF_DASHDOT = 2
    pPDF_DASHDOTDOT = 3
End Enum

Enum PDFFontStl
    FONT_NORMAL = 0
    FONT_ITALIC = 1
    FONT_BOLD = 2
    FONT_UNDERLINE = 3
End Enum

Enum PDFFontNme
    FONT_ARIAL = 0
    FONT_COURIER = 1
    FONT_TIMES = 2
    FONT_SYMBOL = 3
    FONT_ZAPFDINGBATS = 4
End Enum

Enum PDFZoomMd
    ZOOM_FULLPAGE = 0
    ZOOM_FULLWIDTH = 1
    ZOOM_REAL = 2
    ZOOM_DEFAULT = 3
End Enum
        
Enum PDFLayoutMd
    LAYOUT_SINGLE = 0
    LAYOUT_CONTINOUS = 1
    LAYOUT_TWO = 2
    LAYOUT_DEFAULT = 3
End Enum
        
Enum PDFUnitStr
    UNIT_PT = 0
    UNIT_MM = 1
    UNIT_CM = 2
End Enum

Enum PDFOrientationStr
    ORIENT_PAYSAGE = 0
    ORIENT_PORTRAIT = 1
End Enum
                
Enum PDFFormatPgStr
    FORMAT_A4 = 0
    FORMAT_A3 = 1
    FORMAT_A5 = 2
    FORMAT_LETTER = 3
    FORMAT_LEGAL = 4
End Enum

Enum PDFDrawMd
    DRAW_NORMAL = 0
    DRAW_DRAW = 1
    DRAW_DRAWBORDER = 2
End Enum

Enum PDFAlignValue
    ALIGN_CENTER = 0
    ALIGN_LEFT = 1
    ALIGN_RIGHT = 2
    ALIGN_FJUSTIFY = 3
End Enum

Enum PDFBorderValue
    BORDER_NONE = 0
    BORDER_ALL = 1
    BORDER_TOP = 2
    BORDER_BOTTOM = 3
    BORDER_LEFT = 4
    BORDER_RIGHT = 5
End Enum

Enum PDFViewerCst
    VIEW_HIDETOOLBAR = 1
    VIEW_HIDEMENUBAR = 2
    VIEW_HIDEWINDOWUI = 3
    VIEW_FITWINDOW = 4
    VIEW_CENTERWINDOW = 5
    VIEW_DISPLAYDOCTITLE = 6
End Enum
Property Let PDFPathConfiguration(sPathConfig As String)

    wsPathConfig = sPathConfig

End Property
Property Let PDFSetViewerPreferences(pViewerPref As PDFViewerCst)

    bPDFViewerPref = True
    PDFViewerPref = pViewerPref
    
End Property
Property Let PDFWatermark(sWatermark As String)

    bPDFWatermark = True
    sPDFWatermark = sWatermark

End Property
Private Sub PDFRotationText(X As Double, Y As Double, sText As String, pAngle As Integer)

    PDFSetRotation = pAngle
        PDFTextOut sText, X, Y
    PDFSetRotation = 0

End Sub
Private Sub PDFHeader()

Dim dH As Double
Dim dL As Double

    If bPDFWatermark Then
        PDFSetFont FONT_ARIAL, 50, FONT_BOLD
        PDFSetTextColor = Array(255, 192, 203)
        
        dH = (PDFGetPageHeight + PDFGetStringWidth(sPDFWatermark, "", 50) * Sin(45)) / 2.15
        dL = (PDFGetPageWidth - PDFGetStringWidth(sPDFWatermark, "", 50) * Cos(45)) / 2.75
        
        PDFRotationText dL, dH, sPDFWatermark, 45
    End If
    
End Sub
Property Let PDFSetZoomMode(pZoomMode As PDFZoomMd)

    If pZoomMode = ZOOM_FULLPAGE Or pZoomMode = ZOOM_FULLWIDTH Or _
        pZoomMode = ZOOM_REAL Or pZoomMode = ZOOM_DEFAULT Or _
        (IsNumeric(pZoomMode) And (pZoomMode <> ZOOM_FULLPAGE Or _
                                    pZoomMode <> ZOOM_FULLWIDTH Or _
                                    pZoomMode <> ZOOM_REAL Or _
                                    pZoomMode <> ZOOM_DEFAULT)) Then
            If IsNumeric(pZoomMode) Then
                PDFZoomMode = Int(pZoomMode)
            Else
                PDFZoomMode = pZoomMode
            End If
    Else
        MsgBox "Incorrect Zoom Mode : " & pZoomMode & "." & _
                   vbNewLine & _
                   "Focus will be set to full-page zoom", vbCritical, "Zoom Mode - " & PDFClassVersion
        PDFZoomMode = ZOOM_FULLPAGE
    End If

End Property
Property Get PDFGetZoomMode() As Variant

    PDFGetZoomMode = PDFZoomMode

End Property
Property Let PDFUseThumbs(boThumbs As Boolean)

    PDFboThumbs = boThumbs

End Property
Property Let PDFUseOutlines(boOutlines As Boolean)

    PDFboOutlines = boOutlines

End Property
Property Let PDFSetLayoutMode(pLayoutMode As PDFLayoutMd)
    
    If pLayoutMode = LAYOUT_SINGLE Or pLayoutMode = LAYOUT_CONTINOUS Or _
        pLayoutMode = LAYOUT_TWO Or pLayoutMode = LAYOUT_DEFAULT Then
            PDFLayoutMode = pLayoutMode
    Else
        MsgBox "Layout incorrect : " & pLayoutMode & "." & _
                   vbNewLine & _
                   "Layout will be set to simple single page.", vbCritical, "Layout Mode - " & PDFClassVersion
        PDFLayoutMode = LAYOUT_SINGLE
    End If

End Property
Property Get PDFGetLayoutMode() As Variant

    PDFGetLayoutMode = PDFLayoutMode

End Property
Property Let PDFSetUnit(str_Unite As PDFUnitStr)

    Select Case str_Unite
        Case UNIT_PT
            in_Ech = 1
        Case UNIT_MM
            in_Ech = 72 / 25.4
        Case UNIT_CM
            in_Ech = 72 / 2.54
        Case Else
            MsgBox "Incorrect Unit of Measure : " & str_Unite & "." & _
                   vbNewLine & _
                   "Using centimeter ", vbCritical, "Error in measurement unit - " & PDFClassVersion
            in_Ech = 72 / 2.54
    End Select

End Property
Property Get PDFGetUnit() As String

    Select Case in_Ech
        Case 1
            PDFGetUnit = "pt"
        Case 72 / 25.4
            PDFGetUnit = "mm"
        Case 72 / 2.54
            PDFGetUnit = "cm"
    End Select

End Property
Property Let PDFOrientation(str_Orientation As PDFOrientationStr)

Dim tmp_PDFCanvasWidth As Integer
Dim tmp_PDFCanvasHeight As Integer

    ReDim Preserve PDFCanvasWidth(1 To in_Canvas)
    ReDim Preserve PDFCanvasHeight(1 To in_Canvas)
    ReDim Preserve PDFCanvasOrientation(1 To in_Canvas)

    tmp_PDFCanvasWidth = PDFCanvasWidth(in_Canvas)
    tmp_PDFCanvasHeight = PDFCanvasHeight(in_Canvas)

    Select Case str_Orientation
        Case ORIENT_PORTRAIT
            PDFCanvasWidth(in_Canvas) = tmp_PDFCanvasWidth
            PDFCanvasHeight(in_Canvas) = tmp_PDFCanvasHeight
            PDFCanvasOrientation(in_Canvas) = "p"
        Case ORIENT_PAYSAGE
            PDFCanvasWidth(in_Canvas) = tmp_PDFCanvasHeight
            PDFCanvasHeight(in_Canvas) = tmp_PDFCanvasWidth
            PDFCanvasOrientation(in_Canvas) = "l"
        Case Else
            MsgBox "Orientation set incorrectly: " & str_Orientation & "." & _
                   vbNewLine & _
                   "Orientation set to portrait.", vbCritical, "Error in orientation - " & PDFClassVersion
            PDFCanvasWidth(in_Canvas) = tmp_PDFCanvasWidth
            PDFCanvasHeight(in_Canvas) = tmp_PDFCanvasHeight
            PDFCanvasOrientation(in_Canvas) = "p"
    End Select

    ReDim Preserve PDFCanvasWidth(1 To in_Canvas)
    ReDim Preserve PDFCanvasHeight(1 To in_Canvas)
    ReDim Preserve PDFCanvasOrientation(1 To in_Canvas)

End Property
Property Let PDFFormatPage(str_FormatPage As Variant)

    ReDim Preserve PDFCanvasWidth(1 To in_Canvas)
    ReDim Preserve PDFCanvasHeight(1 To in_Canvas)
    ReDim Preserve PDFCanvasOrientation(1 To in_Canvas)

    Select Case TypeName(str_FormatPage)
        Case "Long"
            Select Case str_FormatPage
                Case FORMAT_A4
                    PDFCanvasWidth(in_Canvas) = 595.28
                    PDFCanvasHeight(in_Canvas) = 841.89
                Case FORMAT_A3
                    PDFCanvasWidth(in_Canvas) = 841.89
                    PDFCanvasHeight(in_Canvas) = 1190.55
                Case FORMAT_A5
                    PDFCanvasWidth(in_Canvas) = 420.94
                    PDFCanvasHeight(in_Canvas) = 595.28
                Case FORMAT_LETTER
                    PDFCanvasWidth(in_Canvas) = 612
                    PDFCanvasHeight(in_Canvas) = 792
                Case FORMAT_LEGAL
                    PDFCanvasWidth(in_Canvas) = 612
                    PDFCanvasHeight(in_Canvas) = 1008
                Case Else
                    MsgBox "Format page set incorrectly : " & str_FormatPage & "." & _
                           vbNewLine & _
                           "Format page set to A4.", vbCritical, "Format Page - " & PDFClassVersion
                    PDFCanvasWidth(in_Canvas) = 595.28
                    PDFCanvasHeight(in_Canvas) = 841.89
            End Select
        Case "Double()"
            PDFCanvasWidth(in_Canvas) = str_FormatPage(0)
            PDFCanvasHeight(in_Canvas) = str_FormatPage(1)
        Case Else
            MsgBox "Format page set incorrectly : " & str_FormatPage & "." & _
                   vbNewLine & _
                   "Format page set to A4", vbCritical, "Format Page - " & PDFClassVersion
            PDFCanvasWidth(in_Canvas) = 595.28
            PDFCanvasHeight(in_Canvas) = 841.89
    End Select

End Property
Property Get PDFPageNumber() As Integer

    PDFPageNumber = FPageNumber

End Property
Property Get PDFNbPage() As Integer

    PDFNbPage = UBound(PageNumberList)

End Property
Property Let PDFProducer(str_Producer As String)

    FProducer = str_Producer

End Property
Property Let PDFSubject(str_Subject As String)

    FSubject = str_Subject

End Property
Property Let PDFKeywords(str_Keywords As String)

    FKeywords = str_Keywords

End Property
Property Let PDFCreator(str_Creator As String)

    FCreator = str_Creator

End Property
Property Let PDFAuthor(str_Author As String)

    FAuthor = str_Author

End Property
Property Let PDFTitle(str_Title As String)

    FTitle = str_Title

End Property
Property Let PDFFileName(str_FileName As String)

Dim Items()     As String
Dim sFilePath   As String
Dim sFileName   As String
Dim hWnd        As Long
Dim retval      As Long
Dim in_i        As Long

    On Error GoTo Err_File
    
    FFileName = str_FileName
    
    Items = Split(str_FileName, "\")
    If UBound(Items) = -1 Then Exit Property
    
    sFileName = Items(UBound(Items))
    sFilePath = Left(str_FileName, Len(str_FileName) - Len(Items(UBound(Items))))
    
    sPDFName = Fso.BuildPath(sFilePath, sFileName)
    Set Strm = Fso.CreateTextFile(sPDFName, True)
    
    Exit Property
    
Err_File:
    If Err = 70 Then
        hWnd = FindWindow(vbNullString, "Adobe Reader - [" & sFileName & "]")
        retval = PostMessage(hWnd, WM_CLOSE, 0&, 0&)
        Sleep 17

        Set Strm = Fso.CreateTextFile(sPDFName, True)
        Resume Next
    End If
    
End Property
Property Get PDFGetFileName() As String

    PDFGetFileName = FFileName
    
End Property
Property Let PDFConfirm(boConfirm As Boolean)

    boPDFConfirm = boConfirm

End Property
Property Let PDFView(boView As Boolean)

    boPDFView = boView
    
End Property
Property Let PDFPageHeight(in_PageHeight As Double)

    PDFCanvasHeight(in_Canvas) = in_PageHeight

End Property
Property Get PDFGetPageHeight() As Double

    PDFGetPageHeight = PDFCanvasHeight(in_Canvas)

End Property
Property Let PDFPageWidth(in_PageWidth As Double)

    PDFCanvasWidth(in_Canvas) = in_PageWidth

End Property
Property Get PDFGetPageWidth() As Double

    PDFGetPageWidth = PDFCanvasWidth(in_Canvas)

End Property
Property Let PDFSetLeftMargin(in_left As Double)

    PDFlMargin = in_left

End Property
Property Get PDFGetLeftMargin() As Double

    PDFGetLeftMargin = PDFlMargin

End Property
Property Let PDFSetRightMargin(in_right As Double)

    PDFrMargin = in_right

End Property
Property Get PDFGetRightMargin() As Double

    PDFGetRightMargin = PDFrMargin

End Property
Property Let PDFSetTopMargin(in_top As Double)

    PDFtMargin = in_top

End Property
Property Get PDFGetTopMargin() As Double

    PDFGetTopMargin = PDFtMargin

End Property
Property Let PDFSetBottomMargin(in_bottom As Double)

    PDFbMargin = in_bottom

End Property
Property Get PDFGetBottomMargin() As Double

    PDFGetBottomMargin = PDFbMargin

End Property
Property Let PDFSetCellMargin(in_cell As Double)

    PDFcMargin = in_cell

End Property
Property Get PDFGetCellMargin() As Double

    PDFGetCellMargin = PDFcMargin

End Property
Public Sub PDFSetMargins(in_left As Integer, in_top As Integer, Optional in_right As Integer = -1, Optional in_bottom As Integer = -1)

    PDFlMargin = in_left
    PDFtMargin = in_top

    If in_right = -1 Then in_right = in_left
    If in_bottom = -1 Then in_bottom = in_top

    PDFrMargin = in_right
    PDFbMargin = in_bottom

End Sub
Property Get PDFGetX() As Integer

    PDFGetX = in_xCurrent

End Property
Property Get PDFGetY() As Integer

    PDFGetY = in_yCurrent

End Property
Property Let PDFSetLineStyle(pLineStyle As PDFStyleLgn)

    PDFLnStyle = PDFLineStyle(pLineStyle)

End Property
Property Let PDFSetLineWidth(pLineWidth As Double)

    PDFLnWidth = pLineWidth
    
End Property
Property Let PDFSetDrawMode(pDrawMode As PDFDrawMd)

Dim pTmpDrawMode As String

    pTmpDrawMode = LCase(pDrawMode)

    Select Case pTmpDrawMode
        Case DRAW_NORMAL
            PDFDrawMode = ""
        Case DRAW_DRAW
            PDFDrawMode = "D"
        Case DRAW_DRAWBORDER
            PDFDrawMode = "DB"
        Case Else
            MsgBox "Draw Mode set incorrectly : " & pDrawMode & "." & _
                    vbNewLine & _
                    "Draw mode set to normal", vbCritical, "Object Rectangle - " & PDFClassVersion
            PDFDrawMode = ""
    End Select

End Property
Private Function PDFLineStyle(pLineStyle As PDFStyleLgn) As String

Dim pTmpLineStyle As PDFStyleLgn

    PDFLineStyle = ""
    pTmpLineStyle = pLineStyle

    Select Case pTmpLineStyle
        Case pPDF_SOLID
            PDFLineStyle = "[] 0 d"
        Case pPDF_DASH
            PDFLineStyle = "[" & Int(16 * in_Ech) & " " & Int(8 * in_Ech) & " ] 0 d"
        Case pPDF_DASHDOT
            PDFLineStyle = "[" & Int(8 * in_Ech) & " " & Int(7 * in_Ech) & " " & _
                               Int(2 * in_Ech) & " " & Int(7 * in_Ech) & " ] 0 d"
        Case pPDF_DASHDOTDOT
            PDFLineStyle = "[" & Int(8 * in_Ech) & " " & Int(4 * in_Ech) & " " & _
                               Int(2 * in_Ech) & " " & Int(4 * in_Ech) & " " & _
                               Int(2 * in_Ech) & " " & Int(4 * in_Ech) & " ] 0 d"
        Case Else
            MsgBox "Line style set incorrectly : " & pLineStyle & "." & _
                   vbNewLine & _
                   "Line style set to solid.", vbCritical, "Line Style - " & PDFClassVersion
            PDFLineStyle = "[] 0 d"
    End Select

End Function
Public Sub PDFSetFont(str_Fontname As PDFFontNme, in_FontSize As Integer, Optional str_Style As PDFFontStl)

Dim str_TmpFontName As String
Dim str_TmpFontNm   As String

    If str_Fontname <> FONT_ARIAL And _
       str_Fontname <> FONT_COURIER And _
       str_Fontname <> FONT_SYMBOL And _
       str_Fontname <> FONT_TIMES And _
       str_Fontname <> FONT_ZAPFDINGBATS Then
        MsgBox "Font name set incorrectly : " & str_Style & "." & _
                vbNewLine & _
                "Font set to Times New Roman.", vbCritical, "Font name - " & PDFClassVersion
        str_TmpFontName = "TimesRoman"
        boPDFItalic = False
        boPDFBold = False
        
        PDFFontName = str_TmpFontName
        PDFFontNum = FontNum
        PDFFontSize = in_FontSize

        FontNum = FontNum + 1
        
        Exit Sub
    End If
    
    Select Case str_Fontname
        Case FONT_ARIAL
           str_TmpFontNm = "Arial"
        Case FONT_COURIER
            str_TmpFontNm = "Courier"
        Case FONT_TIMES
            str_TmpFontNm = "Times"
        Case FONT_SYMBOL
            str_TmpFontNm = "Symbol"
        Case FONT_ZAPFDINGBATS
            str_TmpFontNm = "ZapfDingbats"
    End Select

    If str_TmpFontNm = "Arial" Then
        str_TmpFontName = "Helvetica"
    Else
        str_TmpFontName = str_TmpFontNm
    End If

    boPDFItalic = False
    boPDFBold = False

    str_TmpFont = str_TmpFontName
    
    If InStr(1, str_Style, FONT_ITALIC) <> 0 Then boPDFItalic = True
    If InStr(1, str_Style, FONT_BOLD) <> 0 Then boPDFBold = True
    If InStr(1, str_Style, FONT_UNDERLINE) <> 0 Then boPDFUnderline = True
    
    If boPDFItalic = True And boPDFBold = False Then
        Select Case str_TmpFontName
            Case "Times"
                str_TmpFontName = "TimesItalic"
            Case Else
                str_TmpFontName = str_TmpFontName & "-Oblique"
        End Select
    End If

    If boPDFItalic = True And boPDFBold = True Then
        Select Case str_TmpFontName
            Case "Times"
                str_TmpFontName = str_TmpFontName & "-BoldItalic"
            Case Else
                str_TmpFontName = str_TmpFontName & "-BoldOblique"
        End Select
    End If

    If boPDFItalic = False And boPDFBold = True Then
        str_TmpFontName = str_TmpFontName & "-Bold"
    End If
    
    If boPDFItalic = False And boPDFBold = False Then
        Select Case str_TmpFontName
            Case "Times"
                str_TmpFontName = str_TmpFontName & "-Roman"
            Case Else
                str_TmpFontName = str_TmpFontName
        End Select
    End If

    PDFFontName = str_TmpFontName
    PDFFontNum = FontNum
    PDFFontSize = in_FontSize

    FontNum = FontNum + 1

End Sub
Public Sub PDFDrawEllipse(X As Double, Y As Double, rx As Double, Optional ry As Double = 0, Optional URLLink As String = "")

Dim sTempDrawMode As String

    If ry = 0 Then ry = rx
    
    Select Case PDFDrawMode
        Case "D"
            PDFOutStream sTempStream, PDFDrawColor
            sTempDrawMode = "h f"
        Case "DB"
            PDFOutStream sTempStream, PDFDrawColor
            PDFOutStream sTempStream, PDFLineColor
            sTempDrawMode = "B"
        Case ""
            PDFOutStream sTempStream, PDFLineColor
            sTempDrawMode = "s"
    End Select

    PDFOutStream sTempStream, PDFLnStyle
        PDFOutStream sTempStream, PDFFormatDouble(X * in_Ech) & " " & PDFFormatDouble(PDFCanvasHeight(in_Canvas) - (Y + ry / 2) * in_Ech) & " m"
            PDFOutStream sTempStream, PDFCurve(X * in_Ech, _
                PDFCanvasHeight(in_Canvas) - (Y + ry / 2 - ry / 2 * 11 / 20) * in_Ech, _
                (X + rx / 2 - rx / 2 * 11 / 20) * in_Ech, _
                PDFCanvasHeight(in_Canvas) - Y * in_Ech, _
                (X + rx / 2) * in_Ech, _
                PDFCanvasHeight(in_Canvas) - Y * in_Ech)
            PDFOutStream sTempStream, PDFCurve((X + rx / 2 + rx / 2 * 11 / 20) * in_Ech, _
                PDFCanvasHeight(in_Canvas) - Y * in_Ech, _
                (X + rx) * in_Ech, _
                PDFCanvasHeight(in_Canvas) - (Y + ry / 2 - ry / 2 * 11 / 20) * in_Ech, _
                (X + rx) * in_Ech, _
                PDFCanvasHeight(in_Canvas) - (Y + ry / 2) * in_Ech)
            PDFOutStream sTempStream, PDFCurve((X + rx) * in_Ech, _
                PDFCanvasHeight(in_Canvas) - (Y + ry / 2 + ry / 2 * 11 / 20) * in_Ech, _
                (X + rx / 2 + rx / 2 * 11 / 20) * in_Ech, _
                PDFCanvasHeight(in_Canvas) - (Y + ry) * in_Ech, _
                (X + rx / 2) * in_Ech, _
                PDFCanvasHeight(in_Canvas) - (Y + ry) * in_Ech)
            PDFOutStream sTempStream, PDFCurve((X + rx / 2 - rx / 2 * 11 / 20) * in_Ech, _
                PDFCanvasHeight(in_Canvas) - (Y + ry) * in_Ech, _
                X * in_Ech, _
                PDFCanvasHeight(in_Canvas) - (Y + ry / 2 + ry / 2 * 11 / 20) * in_Ech, _
                X * in_Ech, _
                PDFCanvasHeight(in_Canvas) - (Y + ry / 2) * in_Ech)
    PDFOutStream sTempStream, PDFFormatDouble(PDFLnWidth * in_Ech) & " w " & sTempDrawMode

    PDFSetTextColor = vbWhite
    strTLink = "LINK"
    strTyLink = "ELLIPSE"
    PDFSetLink URLLink, "ELLIPSE", Int((X - rx / 2)), Int((Y + ry / 2 - ry / 2 * 11 / 20))
    strTyLink = ""
    
    in_xCurrent = X
    in_yCurrent = Y + ry / 2

End Sub
Private Function PDFCurve(x1, y1, x2, y2, x3, y3 As Double) As String

  PDFCurve = PDFFormatDouble(x1) & " " & _
             PDFFormatDouble(y1) & " " & _
             PDFFormatDouble(x2) & " " & _
             PDFFormatDouble(y2) & " " & _
             PDFFormatDouble(x3) & " " & _
             PDFFormatDouble(y3) & " c"

End Function
Public Sub PDFDrawPolygon(ParamArray pParam() As Variant)

Dim sTempDrawMode As String
Dim nbP           As Double
Dim in_i          As Integer

    nbP = (UBound(pParam(0), 1) + 1) / 2
        
    Select Case PDFDrawMode
        Case "D"
            PDFOutStream sTempStream, PDFDrawColor
            sTempDrawMode = "h f"
        Case "DB"
            PDFOutStream sTempStream, PDFDrawColor
            PDFOutStream sTempStream, PDFLineColor
            sTempDrawMode = "B"
        Case ""
            PDFOutStream sTempStream, PDFLineColor
            sTempDrawMode = "s"
    End Select

    PDFOutStream sTempStream, "%DEBUT_POLY/%"
    PDFOutStream sTempStream, PDFLnStyle
    PDFPoint CDbl(pParam(0)(0)), CDbl(pParam(0)(1))
    For in_i = 2 To nbP * 2 - 1
        If in_i Mod 2 = 0 Then
            PDFLine CDbl(pParam(0)(in_i)), CDbl(pParam(0)(in_i + 1))
        End If
    Next in_i
    
    PDFLine CDbl(pParam(0)(0)), CDbl(pParam(0)(1))
    PDFOutStream sTempStream, PDFFormatDouble(PDFLnWidth * in_Ech) & " w " & sTempDrawMode
    PDFOutStream sTempStream, "%FIN_POLY/%"
    
End Sub
Private Function PDFPoint(X As Double, Y As Double)

    PDFOutStream sTempStream, PDFFormatDouble(X * in_Ech) & " " & _
                              PDFFormatDouble(PDFCanvasHeight(in_Canvas) - Y * in_Ech) & " m"

End Function
Private Function PDFLine(X As Double, Y As Double)

    PDFOutStream sTempStream, PDFFormatDouble(X * in_Ech) & " " & _
                              PDFFormatDouble(PDFCanvasHeight(in_Canvas) - Y * in_Ech) & " l"
End Function
Public Sub PDFDrawLineHor(X As Double, Y As Double, w As Double)

    If Right(PDFLineColor, 2) = "RG" Then
        PDFDrawColor = Left(PDFLineColor, Len(PDFLineColor) - 2) & "rg"
    Else
        PDFDrawColor = Left(PDFLineColor, Len(PDFLineColor) - 1) & "g"
    End If

    PDFOutStream sTempStream, "%DEBUT_LNH/%"
    PDFOutStream sTempStream, PDFLnStyle
    PDFOutStream sTempStream, PDFFormatDouble(X * in_Ech) & " " & PDFFormatDouble(PDFCanvasHeight(in_Canvas) - Y * in_Ech) & " m"
    PDFOutStream sTempStream, PDFFormatDouble((X + w) * in_Ech) & " " & PDFFormatDouble(PDFCanvasHeight(in_Canvas) - Y * in_Ech) & " l"
    PDFOutStream sTempStream, PDFLineColor
    PDFOutStream sTempStream, PDFFormatDouble(PDFLnWidth * in_Ech) & " w S"
    PDFOutStream sTempStream, "%FIN_LNH/%"
    
    in_xCurrent = X + w
    in_yCurrent = Y

End Sub
Public Sub PDFDrawLineVer(X As Double, Y As Double, h As Double)

    If Right(PDFLineColor, 2) = "RG" Then
        PDFDrawColor = Left(PDFLineColor, Len(PDFLineColor) - 2) & "rg"
    Else
        PDFDrawColor = Left(PDFLineColor, Len(PDFLineColor) - 1) & "g"
    End If
    
    PDFOutStream sTempStream, "%DEBUT_LNV/%"
    PDFOutStream sTempStream, PDFLnStyle
    PDFOutStream sTempStream, PDFFormatDouble(X * in_Ech) & " " & PDFFormatDouble(PDFCanvasHeight(in_Canvas) - Y * in_Ech) & " m"
    PDFOutStream sTempStream, PDFFormatDouble(X * in_Ech) & " " & PDFFormatDouble(PDFCanvasHeight(in_Canvas) - (Y + h) * in_Ech) & " l"
    PDFOutStream sTempStream, PDFLineColor
    PDFOutStream sTempStream, PDFFormatDouble(PDFLnWidth * in_Ech) & " w S"
    PDFOutStream sTempStream, "%FIN_LNV/%"
    
    in_xCurrent = X
    in_yCurrent = Y + h

End Sub
Public Sub PDFDrawLine(x1 As Double, y1 As Double, x2 As Double, y2 As Double)

    PDFOutStream sTempStream, "%DEBUT_LN/%"
    PDFOutStream sTempStream, PDFLnStyle
    PDFOutStream sTempStream, PDFFormatDouble(x1 * in_Ech) & " " & PDFFormatDouble(PDFCanvasHeight(in_Canvas) - y1 * in_Ech) & " m"
    PDFOutStream sTempStream, PDFFormatDouble(x2 * in_Ech) & " " & PDFFormatDouble(PDFCanvasHeight(in_Canvas) - y2 * in_Ech) & " l"
    PDFOutStream sTempStream, PDFLineColor
    PDFOutStream sTempStream, PDFFormatDouble(PDFLnWidth * in_Ech) & " w S"
    PDFOutStream sTempStream, "%FIN_LN/%"
    
    If x1 > x2 Then
        in_xCurrent = x1
    Else
        in_xCurrent = x2
    End If

    If y1 > y2 Then
        in_yCurrent = y1
    Else
        in_yCurrent = y2
    End If


End Sub
Public Sub PDFDrawRectangle(X As Double, Y As Double, w As Double, h As Double, Optional URLLink As String = "")

Dim sTempDrawMode As String
        
    PDFOutStream sTempStream, "%DEBUT_RECT/%"
    Select Case PDFDrawMode
        Case "D"
            PDFOutStream sTempStream, PDFDrawColor
            sTempDrawMode = "f"
        Case "DB"
            PDFOutStream sTempStream, PDFDrawColor
            PDFOutStream sTempStream, PDFLineColor
            sTempDrawMode = "B"
        Case ""
            PDFOutStream sTempStream, PDFLineColor
            sTempDrawMode = "s"
    End Select
    
    PDFOutStream sTempStream, PDFLnStyle
    PDFOutStream sTempStream, PDFFormatDouble(X * in_Ech) & " " & _
                              PDFFormatDouble(PDFCanvasHeight(in_Canvas) - Y * in_Ech) & " " & _
                              PDFFormatDouble(w * in_Ech) & " " & _
                              PDFFormatDouble(-1 * h * in_Ech) & " re " & sTempDrawMode
    PDFOutStream sTempStream, PDFFormatDouble(PDFLnWidth * in_Ech) & " w S"

    PDFSetTextColor = vbWhite
    
    strTLink = "LINK"
    strTyLink = "RECTANGLE"
    wRect = w
    PDFSetLink URLLink, "RECTANGLE", Int(X + 5), Int(Y + h / 2)
    PDFOutStream sTempStream, "%FIN_RECT/%"

    strTyLink = ""
    
    in_xCurrent = X
    in_yCurrent = Y + h
    
End Sub
Private Function PDFHtml2RgbColor(sColor As String) As PDFRGB

Dim sTmpColor As String

    sTmpColor = Right("000000" & sColor, 6)
    PDFHtml2RgbColor.in_r = CByte("&h" & Mid(sTmpColor, 1, 2))
    PDFHtml2RgbColor.in_g = CByte("&h" & Mid(sTmpColor, 3, 2))
    PDFHtml2RgbColor.in_b = CByte("&h" & Mid(sTmpColor, 5, 2))

End Function
Property Let PDFSetTextColor(gColor As Variant)

Dim TxtCl     As PDFRGB
Dim sColor    As String

    Select Case TypeName(gColor)
        Case "Variant()"
            TxtCl.in_r = gColor(0)
            TxtCl.in_g = gColor(1)
            TxtCl.in_b = gColor(2)
        Case "String"
           If Left(gColor, 1) <> "#" Then
                MsgBox "Invalid HTMl color set" & gColor & "." & _
                       vbNewLine & _
                       "Set color to  black.", vbCritical, "Text Color " & PDFClassVersion
                TxtCl = PDFGetRGB(vbBlack)
            Else
                TxtCl = PDFHtml2RgbColor(CStr(gColor))
            End If
        Case Else
            TxtCl = PDFGetRGB(Int(gColor))
    End Select

    PDFTextColor = PDFStreamColor(TxtCl, "TEXT")

End Property
Property Get PDFGetTextColor() As String

    PDFGetTextColor = PDFstrTextColor

End Property
Property Let PDFSetLineColor(gColor As Variant)

Dim TxtCl     As PDFRGB
Dim sColor    As String

    Select Case TypeName(gColor)
        Case "Variant()"
            TxtCl.in_r = gColor(0)
            TxtCl.in_g = gColor(1)
            TxtCl.in_b = gColor(2)
        Case "String"
           If Left(gColor, 1) <> "#" Then
                MsgBox "Invalid line color set " & gColor & "." & _
                       vbNewLine & _
                       "Setting line color to black.", vbCritical, "Line Color - " & PDFClassVersion
                TxtCl = PDFGetRGB(vbBlack)
            Else
                TxtCl = PDFHtml2RgbColor(CStr(gColor))
            End If
        Case Else
            TxtCl = PDFGetRGB(Int(gColor))
    End Select

    PDFLineColor = PDFStreamColor(TxtCl, "LINE")

End Property
Property Get PDFGetLineColor() As String

    PDFGetLineColor = PDFstrLineColor

End Property
Property Let PDFSetDrawColor(gColor As Variant)

Dim TxtCl     As PDFRGB
Dim sColor    As String

    Select Case TypeName(gColor)
        Case "Variant()"
            TxtCl.in_r = gColor(0)
            TxtCl.in_g = gColor(1)
            TxtCl.in_b = gColor(2)
        Case "String"
           If Left(gColor, 1) <> "#" Then
                MsgBox "Invalid Draw Color set " & gColor & "." & _
                       vbNewLine & _
                       "Using black.", vbCritical, "Draw Color - " & PDFClassVersion
                TxtCl = PDFGetRGB(vbBlack)
            Else
                TxtCl = PDFHtml2RgbColor(CStr(gColor))
            End If
        Case Else
            TxtCl = PDFGetRGB(Int(gColor))
    End Select
    
    PDFDrawColor = PDFStreamColor(TxtCl, "BORDER")

End Property
Property Get PDFGetDrawColor() As String

    PDFGetDrawColor = PDFstrDrawColor

End Property
Private Function PDFStreamColor(PDFRgbColor As PDFRGB, str_Type As String) As String

Dim int_r        As Integer
Dim int_g        As Integer
Dim int_b        As Integer
Dim str_TxtColor As String

    int_r = PDFRgbColor.in_r
    int_g = PDFRgbColor.in_g
    int_b = PDFRgbColor.in_b

    Select Case str_Type
        Case "TEXT", "BORDER"
            str_TxtColor = Replace(Format(int_r / 255, "0.000"), ",", ".") & " " & _
                           Replace(Format(int_g / 255, "0.000"), ",", ".") & " " & _
                           Replace(Format(int_b / 255, "0.000"), ",", ".") & " rg"
        Case "LINE"
            str_TxtColor = Replace(Format(int_r / 255, "0.000"), ",", ".") & " " & _
                           Replace(Format(int_g / 255, "0.000"), ",", ".") & " " & _
                           Replace(Format(int_b / 255, "0.000"), ",", ".") & " RG"
    End Select

    PDFStreamColor = str_TxtColor

End Function
Property Let PDFSetAlignement(gAlignement As PDFAlignValue)

    Select Case gAlignement
        Case 2
            PDFstrTempAlign = "R"
        Case 0
            PDFstrTempAlign = "C"
        Case 1
            PDFstrTempAlign = "L"
        Case 3
            PDFstrTempAlign = "FJ"
        Case Else
            MsgBox "Invalid alignment set. : " & gAlignement & "." & _
                   vbNewLine & _
                   "Using left alignment.", vbCritical, "Alignment - " & PDFClassVersion
            PDFstrTempAlign = "L"
    End Select

End Property
Property Get PDFGetAlignement() As String

Dim strTempAlign As String

    Select Case PDFstrTempAlign
        Case "C"
            strTempAlign = "Center"
        Case "R"
            strTempAlign = "Right"
        Case "L"
            strTempAlign = "Left"
        Case Else
            strTempAlign = "Left"
    End Select
    
    PDFGetAlignement = strTempAlign

End Property
Public Sub PDFLink(X As Double, Y As Double, str_Text As String, Optional str_Link As String = "")

Dim w As Integer
Dim h As Integer

    pTempAngle = 0
    
    PDFOutStream sTempStream, "%DEBUT_LINK/%"
    
    boPDFUnderline = True
    
        If PDFboImage = True Then
            PDFSetTextColor = vbBlue
            w = Int(ImgWidth)
            h = Int(ImgHeight)
            PDFTextOut "", X, Y
        Else
            Select Case strTyLink
                Case "ELLIPSE"
                    w = Int(PDFGetStringWidth(strTLink, PDFFontName, PDFFontSize))
                    h = Int(PDFFontSize)
                    PDFTextOut "", X, Y
                Case "RECTANGLE"
                    w = wRect
                    h = Int(PDFFontSize)
                    PDFTextOut "", X, Y
                Case "CELL"
                    w = Int(PDFGetStringWidth(strTLink, PDFFontName, PDFFontSize))
                    h = Int(PDFFontSize)
                    PDFTextOut "", X, Y
                Case Else
                    w = Int(PDFGetStringWidth(str_Text, PDFFontName, PDFFontSize))
                    h = Int(PDFFontSize)
                    PDFTextOut str_Text, X, Y
            End Select
        End If

    PDFboImage = False
    boPDFUnderline = False
    
    strTyLink = ""
    If str_Link = "" Then str_Link = str_Text
    
    PDFTabLinks X, Y, w, h, str_Text, str_Link

    PDFOutStream sTempStream, "%FIN_LINK/%"
    
End Sub
Private Sub PDFTabLinks(X As Double, Y As Double, w As Integer, h As Integer, str_Text As String, Optional str_Link As Variant = 0)

    FPageLink = FPageLink + 1
    ReDim Preserve LinksList(1 To FPageLink)
    LinksList(FPageLink) = Array(FPageNumber, Y, str_Link)

    If str_Link <> 0 Then
        PageLinksList(FPageNumber, FPageLink) = Array(X * in_Ech, PDFCanvasHeight(in_Canvas) - Y * in_Ech, w * in_Ech, h * in_Ech, str_Link)
    Else
        PageLinksList(FPageNumber, FPageLink) = Array(X * in_Ech, PDFCanvasHeight(in_Canvas) - Y * in_Ech, w * in_Ech, h * in_Ech, str_Text)
    End If

    ReDim Preserve boPageLinksList(1 To FPageNumber)
    ReDim Preserve NbPageLinksList(1 To FPageNumber)

    boPageLinksList(FPageNumber) = True
    NbPageLinksList(FPageNumber) = FPageLink

End Sub
Property Get PDFTextHeight() As Double

    PDFTextHeight = PDFFontSize * in_Ech
    
End Property
Property Let PDFSetRotation(pAngle As Double)

    PDFAngle = -1 * pAngle

End Property
Private Sub PDFStreamRotate(pAngle As Double, X As Double, Y As Double)

Dim dSin     As Double
Dim dCos     As Double
Dim CenterX  As Double
Dim CenterY  As Double

    If pAngle <> 0 Then
        pAngle = pAngle * 3.1416 / 180
        dCos = Cos(pAngle)
        dSin = Sin(pAngle)
        CenterX = X * in_Ech
        CenterY = PDFCanvasHeight(in_Canvas) - Y * in_Ech
        
        PDFOutStream sTempStream, PDFFormatDouble(dCos, 5) & " " & _
                                  PDFFormatDouble(-1 * dSin, 5) & " " & _
                                  PDFFormatDouble(dSin, 5) & " " & _
                                  PDFFormatDouble(dCos, 5) & " " & _
                                  PDFFormatDouble(CenterX) & " " & _
                                  PDFFormatDouble(CenterY) & " Tm"
    End If
    
    bAngle = True
    
End Sub
Public Sub PDFTextOut(str_Text As String, Optional X As Double = 0, Optional Y As Double = 0)

Dim j               As Integer
Dim in_PositionFont As Integer
Dim str_Tmp         As String
Dim str_TmpText     As String

    str_TmpText = Replace(str_Text, "\", "\\")
    str_TmpText = Replace(str_TmpText, "\\", "\\\\")
    str_TmpText = Replace(str_TmpText, "(", "\(")
    str_TmpText = Replace(str_TmpText, ")", "\)")
    
    str_Tmp = ""

    If X = 0 Then X = in_xCurrent
    If Y = 0 Then Y = in_yCurrent
    
    If PDFFontName = "" Then
        in_PositionFont = 1
    Else
        For j = 0 To UBound(Arr_Font)
            If Arr_Font(j) = PDFFontName Then
                in_PositionFont = j + 1
                Exit For
            End If
        Next j
    End If

    If PDFFontSize = 0 Then PDFFontSize = 10
    If PDFTextColor <> "" Then PDFOutStream sTempStream, "q " & PDFTextColor & " "
    If boPDFUnderline Then str_Tmp = PDFUnderline(False, str_Text, CDbl(X * in_Ech), CDbl(Y * in_Ech))
    
    PDFOutStream sTempStream, "%DEBUT_TEXT/%"
    PDFOutStream sTempStream, "BT"
    
    If PDFAngle = 0 Then
        PDFOutStream sTempStream, PDFFormatDouble((X + PDFlMargin) * in_Ech) & " " & PDFFormatDouble(PDFCanvasHeight(in_Canvas) - Y * in_Ech) & " Td"
    Else
        PDFStreamRotate PDFAngle, X, Y
        PDFAngle = 0
    End If
    
    PDFOutStream sTempStream, "/F" & in_PositionFont & " " & PDFFormatDouble(PDFFontSize) & " Tf"
    PDFOutStream sTempStream, "(" & str_TmpText & ") Tj"
    
    If PDFTextColor <> "" Then
        PDFOutStream sTempStream, "ET"

        If boPDFUnderline = True Then
            PDFOutStream sTempStream, str_Tmp
        End If

        PDFOutStream sTempStream, "Q"
    Else
        PDFOutStream sTempStream, "ET"

        If boPDFUnderline = True Then
            PDFOutStream sTempStream, str_Tmp
        End If
    End If
    
    PDFOutStream sTempStream, "%FIN_TEXT/%"
    
    boPDFUnderline = False

    in_xCurrent = X + PDFGetStringWidth(str_Text, PDFFontName, PDFFontSize)
    in_yCurrent = Y + PDFFontSize

End Sub
Property Let PDFSetBorder(gBorder As PDFBorderValue)

    PDFstrTempBorder = ""

    Select Case gBorder
        Case BORDER_ALL
            PDFstrTempBorder = "1"
        Case BORDER_NONE
            PDFstrTempBorder = "0"
        Case BORDER_TOP
            PDFstrTempBorder = "T"
        Case BORDER_BOTTOM
            PDFstrTempBorder = "B"
        Case BORDER_LEFT
            PDFstrTempBorder = "L"
        Case BORDER_RIGHT
            PDFstrTempBorder = "R"
        Case Else
            If InStr(1, gBorder, BORDER_LEFT, 1) <> 0 Then PDFstrTempBorder = PDFstrTempBorder & "L"
            If InStr(1, gBorder, BORDER_RIGHT, 1) <> 0 Then PDFstrTempBorder = PDFstrTempBorder & "R"
            If InStr(1, gBorder, BORDER_TOP, 1) <> 0 Then PDFstrTempBorder = PDFstrTempBorder & "T"
            If InStr(1, gBorder, BORDER_BOTTOM, 1) <> 0 Then PDFstrTempBorder = PDFstrTempBorder & "B"
    End Select

End Property
Property Let PDFSetFill(bFill As Boolean)

    PDFboTempFill = bFill

End Property
Public Sub PDFCell(str_Text As String, X As Double, Y As Double, w As Double, h As Double, Optional URLLink As String = "")
  
Dim WidthMax    As Double
Dim lText       As Integer
Dim sCar        As String
Dim tWidth      As Double
Dim tBorder     As String
Dim yPos        As Double
Dim bMulti      As Boolean
Dim bBorder1    As String
Dim bBorder2    As String
Dim iSep        As Integer
Dim I, j, l     As Integer
Dim nl          As Integer

    tWidth = w
    yPos = Y
    
    WidthMax = (w - 2 * PDFcMargin) * 10 / PDFFontSize
    lText = Len(str_Text)
    
    If lText > 0 And Right(str_Text, lText - 1) = vbNewLine Then
        lText = lText - 1
    End If
 
    bBorder1 = ""
        
    tBorder = PDFstrTempBorder
    If PDFstrTempBorder = "LRTB" Or PDFstrTempBorder = 1 Then
        bBorder1 = "LRT"
        bBorder2 = "LR"
    Else
        bBorder2 = ""
        If InStr(1, PDFstrTempBorder, "L", 1) <> 0 Then bBorder2 = bBorder2 & BORDER_LEFT
        If InStr(1, PDFstrTempBorder, "R", 1) <> 0 Then bBorder2 = bBorder2 & BORDER_RIGHT
        bBorder1 = IIf(InStr(1, PDFstrTempBorder, "T", 1) <> 0, bBorder2 = bBorder2 & BORDER_TOP, bBorder2)
    End If
    
    iSep = -1
    I = 1
    j = 1
    l = 0

    nl = 1
    
    PDFOutStream sTempStream, "%DEBUT_CELL/%"
    
    While I <= lText
        sCar = Mid(str_Text, I, 1)
        
        If sCar = vbCrLf Then
            PDFstrTempBorder = bBorder1
            PDFCell2 Mid(str_Text, j, I - j), X, yPos, tWidth, h
            yPos = in_yCurrent
            
            bMulti = True
            
            I = I + 1
            
            iSep = -1
            j = I
            l = 0

            nl = nl + 1
            
            If nl = 2 Then bBorder1 = bBorder2
         End If
        
        If sCar = " " Then
            iSep = I
        End If
        
        l = l + PDFGetStringWidth(sCar, PDFFontName, PDFFontSize)
        
        If l > WidthMax Then
            If iSep = -1 Then
                If I = j Then I = I + 1
                
                PDFstrTempBorder = bBorder1
                PDFCell2 Mid(str_Text, j, I - j), X, yPos, tWidth, h
                yPos = in_yCurrent
                               
                bMulti = True
            Else
                PDFstrTempBorder = bBorder1
                PDFCell2 Mid(str_Text, j, iSep - j), X - PDFcMargin, yPos, tWidth, h
                yPos = in_yCurrent
            
                bMulti = True
                I = iSep + 1
            End If
            
            iSep = -1
            
            j = I
            l = 0
            
            nl = nl + 1
            
            If nl = 2 Then bBorder1 = bBorder2
        Else
            I = I + 1
        End If
    Wend
    
    If InStr(1, tBorder, "B", 1) <> 0 Or tBorder = 1 Then
        bBorder1 = bBorder1 & "B"
        PDFstrTempBorder = bBorder1
    End If
    
    yPos = IIf(bMulti, in_yCurrent, yPos)
    PDFCell2 Mid(str_Text, j, I - j), X - PDFcMargin, yPos, tWidth, h
    
    boPDFUnderline = False
    
    If PDFstrTempAlign = "FJ" Then
        PDFOutStream sTempStream, "0 Tw"
        iWidthStr = 0
    End If
    
    PDFOutStream sTempStream, "%FIN_CELL/%"
    
End Sub
Private Function PDFGetNumberOfCar(sText As String, sCar As String) As Integer

Dim iNbCar As Integer
Dim in_i   As Integer

    iNbCar = 0
    in_i = InStr(1, sText, sCar)
    If in_i <> 0 Then iNbCar = 1
    
    Do While in_i <> 0
        in_i = InStr(in_i + 1, sText, sCar)
        If in_i <> 0 Then iNbCar = iNbCar + 1
    Loop
    
    PDFGetNumberOfCar = iNbCar
    
End Function
Private Sub PDFCell2(str_Text As String, X As Double, Y As Double, w As Double, h As Double, Optional URLLink As String = "")

Dim j               As Integer
Dim dx              As Integer
Dim ltmp            As Integer

Dim in_PositionFont As Integer
Dim str_Tmp         As String
Dim str_TmpSTR      As String
Dim str_TmpText     As String

Dim in_Px           As Integer
Dim in_Pw           As String
Dim in_Py           As String
Dim iWidthMax       As Double

Dim str_Tmp1        As String

    str_TmpText = Replace(str_Text, "\", "\\")
    str_TmpText = Replace(str_TmpText, "\\", "\\\\")
    str_TmpText = Replace(str_TmpText, "(", "\(")
    str_TmpText = Replace(str_TmpText, ")", "\)")

    str_Tmp1 = ""

    dx = 0
    'x = x + PDFcMargin

    If PDFFontName = "" Then
        in_PositionFont = 1
    Else
        For j = 0 To UBound(Arr_Font)
            If Arr_Font(j) = PDFFontName Then
                in_PositionFont = j + 1
                Exit For
            End If
        Next j
    End If

    If PDFFontSize = 0 Then PDFFontSize = 10
    If PDFLineColor <> "" Then PDFOutStream sTempStream, Trim(PDFLineColor)
    If PDFDrawColor <> "" Then PDFOutStream sTempStream, PDFDrawColor

    If PDFboTempFill = True Or PDFstrTempBorder = "1" Then
        If PDFboTempFill = True Then
            If PDFstrTempBorder = "1" Then
                str_Tmp = "B"
            Else
                str_Tmp = "f"
            End If
        Else
            str_Tmp = "S"
        End If
        
        str_TmpSTR = PDFFormatDouble(X * in_Ech) & " " & _
                     PDFFormatDouble(PDFCanvasHeight(in_Canvas) - Y * in_Ech) & " " & _
                     PDFFormatDouble(w * in_Ech) & " " & _
                     PDFFormatDouble(-h * in_Ech) & " re " & str_Tmp & vbCr
    End If

    If PDFstrTempBorder <> "0" And PDFstrTempBorder <> "1" Then
        PDFOutStream sTempStream, PDFFormatDouble(PDFLnWidth * in_Ech) & " w"
    
        If InStr(1, PDFstrTempBorder, "L", 1) <> 0 Then _
            str_TmpSTR = str_TmpSTR & PDFFormatDouble(X * in_Ech) & " " & _
                         PDFFormatDouble(PDFCanvasHeight(in_Canvas) - Y * in_Ech) & " m " & PDFFormatDouble(X * in_Ech) & " " & _
                         PDFFormatDouble(PDFCanvasHeight(in_Canvas) - (Y + h) * in_Ech) & " l S" & vbCr
        If InStr(1, PDFstrTempBorder, "T", 1) <> 0 Then _
            str_TmpSTR = str_TmpSTR & PDFFormatDouble(X * in_Ech) & " " & _
                         PDFFormatDouble(PDFCanvasHeight(in_Canvas) - Y * in_Ech) & " m " & PDFFormatDouble(X * in_Ech + w * in_Ech) & " " & _
                         PDFFormatDouble(PDFCanvasHeight(in_Canvas) - Y * in_Ech) & " l S " & vbCr
        If InStr(1, PDFstrTempBorder, "R", 1) <> 0 Then _
            str_TmpSTR = str_TmpSTR & PDFFormatDouble(X * in_Ech + w * in_Ech) & " " & _
                         PDFFormatDouble(PDFCanvasHeight(in_Canvas) - Y * in_Ech) & " m " & PDFFormatDouble(X * in_Ech + w * in_Ech) & " " & _
                         PDFFormatDouble(PDFCanvasHeight(in_Canvas) - (Y + h) * in_Ech) & " l S " & vbCr
        If InStr(1, PDFstrTempBorder, "B", 1) <> 0 Then _
            str_TmpSTR = str_TmpSTR & PDFFormatDouble(X * in_Ech) & " " & _
                         PDFFormatDouble(PDFCanvasHeight(in_Canvas) - (Y + h) * in_Ech) & " m " & PDFFormatDouble(X * in_Ech + w * in_Ech) & " " & _
                         PDFFormatDouble(PDFCanvasHeight(in_Canvas) - (Y + h) * in_Ech) & " l S " & vbCr
    End If

    PDFstrTempBorder = "0"
    
    If PDFstrTempAlign = "" Then PDFstrTempAlign = "L"
    
    Select Case PDFstrTempAlign
        Case "R"
            ltmp = PDFGetStringWidth(str_Text, PDFFontName, PDFFontSize)
            dx = w * in_Ech - PDFcMargin - Format(ltmp, "###0.00")
        Case "C"
            ltmp = PDFGetStringWidth(str_Text, PDFFontName, PDFFontSize)
            dx = (w * in_Ech - ltmp) / 2
        Case "L"
            dx = 2 * PDFcMargin
        Case "FJ"
            iWidthMax = (w * in_Ech - (PDFGetNumberOfCar(str_Text, " ") + 1) * PDFcMargin)
            iWidthStr = (iWidthMax - PDFGetStringWidth(str_Text, PDFFontName, PDFFontSize)) / IIf(PDFGetNumberOfCar(str_Text, " ") <> 0, PDFGetNumberOfCar(str_Text, " "), 1)
            PDFOutStream sTempStream, PDFFormatDouble(iWidthStr * in_Ech, 3) & " Tw"
            dx = 2 * PDFcMargin
    End Select

    If str_TmpSTR <> "" Then PDFOutStream sTempStream, str_TmpSTR

    If URLLink <> "" Then
        boPDFUnderline = True
        PDFTabLinks (X + dx), _
                (Y + 0.5 * h - 0.5 * PDFFontSize), _
                PDFGetStringWidth(str_Text, PDFFontName, PDFFontSize), _
                CDbl(PDFFontSize), _
                str_Text, URLLink
    End If

    If boPDFUnderline Then str_Tmp1 = PDFUnderline(True, str_Text, CDbl((X * in_Ech + dx)), _
                                                PDFCanvasHeight(in_Canvas) - (Y * in_Ech + 0.5 * h * in_Ech + 0.3 * PDFFontSize))

    If PDFTextColor <> "" Then
        PDFOutStream sTempStream, "q " & PDFTextColor & " "
        If boPDFUnderline = True Then
            PDFOutStream sTempStream, str_Tmp1
        End If
    End If

    xlink = 0
    xlink = X

    yLink = 0
    yLink = Y
    
    PDFOutStream sTempStream, "BT"
    PDFOutStream sTempStream, "/F" & in_PositionFont & " " & PDFFontSize & " Tf"
    PDFOutStream sTempStream, PDFFormatDouble((X * in_Ech + dx)) & " " & _
                              PDFFormatDouble((PDFCanvasHeight(in_Canvas) - (Y * in_Ech + 0.5 * h * in_Ech + 0.3 * PDFFontSize))) & _
                              " Td"
    PDFOutStream sTempStream, "(" & str_TmpText & ") Tj"

    If PDFTextColor <> "" Then
        PDFOutStream sTempStream, "ET"
        PDFOutStream sTempStream, "Q"
    Else
        PDFOutStream sTempStream, "ET"
    End If
    
    strTLink = str_Text
    strTyLink = "CELL"
    
    PDFSetLink URLLink, "CELL", xlink, yLink
    strTyLink = ""
    
    in_xCurrent = X + w
    in_yCurrent = Y + h

End Sub
Private Sub PDFSetLink(URLLink As String, OType As String, X As Double, Y As Double)

    If TypeName(URLLink) = "String" Then
        If OType = "IMAGE" Then
            PDFboImage = True
        Else
            PDFboImage = False
        End If

        If URLLink <> "" Then PDFLink X, Y, URLLink
        strTLink = ""
        PDFboImage = False
    Else
        Select Case OType
            Case "CELL"
                MsgBox "Invalid URL link : " & URLLink & "." & _
                        vbNewLine & _
                        "Unable to include link.", vbCritical, "Url Link - " & PDFClassVersion
            Case "IMAGE"
                MsgBox "Invalid URL image object: " & URLLink & "." & _
                        vbNewLine & _
                        "Unable to include URL image.", vbCritical, "Url Link Image - " & PDFClassVersion
            Case "RECT"
                MsgBox "Invalid URL rectangle: " & URLLink & "." & _
                        vbNewLine & _
                        "Unable to include URL rectangle.", vbCritical, "Url Link Rectangle - " & PDFClassVersion
            Case "ELLIPSE"
                MsgBox "Invalid URL Ellipse : " & URLLink & "." & _
                        vbNewLine & _
                        "Unable ot include URL Ellipse.", vbCritical, "Url Link Ellipse - " & PDFClassVersion
        End Select
    End If

End Sub
Public Function PDFImageWidth(pFileName As String) As Double

Dim ArrInfo  As Variant
Dim in_pos   As Integer

    in_pos = InStr(1, pFileName, ".", 1)

    If in_pos = 0 Then
        MsgBox "File " & pFileName & " does not have an extension" & _
                vbNewLine & _
                "Invalid filename specified.", vbCritical, "Image File - " & PDFClassVersion
        Exit Function
    End If

    If Right(pFileName, 3) = "jpg" Or Right(pFileName, 4) = "jpeg" Then
        ArrInfo = PDFParseJPG(pFileName)
        If TypeName(ArrInfo) = "Boolean" Then
            If ArrInfo = False Then Exit Function
        End If
    Else
        MsgBox "Image format not supported." & _
                vbNewLine & _
                "Only JPEG images are supported." & _
                vbNewLine & _
                "Impossible to include image in PDF file.", vbCritical, "Image File - " & PDFClassVersion
        Exit Function
    End If

    PDFImageWidth = ArrInfo(0)
    
End Function
Public Function PDFImageHeight(pFileName As String) As Double

Dim ArrInfo  As Variant
Dim in_pos   As Integer

    in_pos = InStr(1, pFileName, ".", 1)

    If in_pos = 0 Then
        MsgBox "File " & pFileName & " does not have an extension" & _
                vbNewLine & _
                "Invalid filename specified.", vbCritical, "Image File - " & PDFClassVersion
        Exit Function
    End If

    If Right(pFileName, 3) = "jpg" Or Right(pFileName, 4) = "jpeg" Then
        ArrInfo = PDFParseJPG(pFileName)
        If TypeName(ArrInfo) = "Boolean" Then
            If ArrInfo = False Then Exit Function
        End If
    Else
        MsgBox "Image format not supported." & _
                vbNewLine & _
                "Only JPEG images are supported." & _
                vbNewLine & _
                "Impossible to include image in PDF file.", vbCritical, "Image File - " & PDFClassVersion
        Exit Function
    End If

    PDFImageHeight = ArrInfo(1)
    
End Function
Public Sub PDFImage(pFileName As String, X As Double, Y As Double, Optional w As Double = 0, Optional h As Double = 0, Optional URLLink As String = "")

Dim in_pos   As Integer
Dim ArrInfo  As Variant

    in_pos = InStr(1, pFileName, ".", 1)

    If in_pos = 0 Then
        MsgBox "File " & pFileName & " does not have an extension" & _
                vbNewLine & _
                "Invalid filename specified.", vbCritical, "Image File - " & PDFClassVersion
        Exit Sub
    End If

    If Right(pFileName, 3) = "jpg" Or Right(pFileName, 4) = "jpeg" Then
        ArrInfo = PDFParseJPG(pFileName)
        If TypeName(ArrInfo) = "Boolean" Then
            If ArrInfo = False Then Exit Sub
        End If
    Else
        MsgBox "Image format not supported." & _
                vbNewLine & _
                "Only JPEG images are supported." & _
                vbNewLine & _
                "Impossible to include image in PDF file.", vbCritical, "Image File - " & PDFClassVersion
        Exit Sub
    End If

    If w = 0 And h = 0 Then
        w = ArrInfo(0) / in_Ech
        h = ArrInfo(1) / in_Ech
    End If

    If w = 0 Then w = h * ArrInfo(0) / ArrInfo(1)
    If h = 0 Then h = w * ArrInfo(1) / ArrInfo(0)

    NumberofImages = NumberofImages + 1
       
    PDFOutStream sTempStream, "q"
        
    PDFOutStream sTempStream, PDFFormatDouble(w * in_Ech) & " 0 0 " & _
                              PDFFormatDouble(h * in_Ech) & " " & _
                              PDFFormatDouble(X * in_Ech) & " " & _
                              PDFFormatDouble(PDFCanvasHeight(in_Canvas) - (Y + h) * in_Ech) & " cm /ImgJPEG" & _
                              NumberofImages & " Do Q"
    
    ImgWidth = w
    ImgHeight = h

    PDFSetLink URLLink, "IMAGE", X, Y

    in_xCurrent = (X + w) * in_Ech
    in_yCurrent = (Y + h) * in_Ech

End Sub
Private Function PDFParseJPG(pFileName As String) As Variant

Const OPEN_EXISTING = 3
Const FILE_SHARE_READ = &H1
Const GENERIC_READ = &H80000000
Const FILE_BEGIN = 0

Dim in_File    As Long
Dim in_Bytes   As Long

Dim str_TChar  As String
Dim in_res     As Long

Dim sIMG       As Long
Dim inIMG

Dim in_PEnd     As Long
Dim in_idx      As Long
Dim str_SegmMk  As String
Dim in_SegmSz   As Long
Dim bChar       As Byte
Dim in_TmpColor As Long
Dim in_bpc      As Long

Dim ArrBFile()  As Byte

    ReDim Preserve ArrIMG(1 To NumberofImages + 1)

    ' Extract info from a JPEG file
    inIMG = FreeFile

    in_File = PDFCreateFile(pFileName, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, 0, 0)
    sIMG = PDFGetFileSize(in_File, 0)

    If sIMG < 250 Then
        MsgBox "File Image is non JPEG" & _
                vbNewLine & _
                "Cannot add image to PDF file.", vbCritical, "File Image - " & PDFClassVersion
        PDFParseJPG = False
        PDFCloseHandle in_File
        Exit Function
    End If

    ArrIMG(NumberofImages + 1).in_8 = sIMG

    ReDim Preserve ArrBFile(1 To 1, 1 To sIMG) As Byte
    in_res = PDFReadFile(in_File, ArrBFile(1, 1), sIMG, in_Bytes, ByVal 0&)

    in_PEnd = UBound(ArrBFile, 2) - 1

    If PDFIntAsHex(ArrBFile, 1) <> "FFD8" Or PDFIntAsHex(ArrBFile, in_PEnd) <> "FFD9" Then
        MsgBox "Invalid JPEG marker" & _
                vbNewLine & _
                "Cannot add iamge to PDF file.", vbCritical, "File Image - " & PDFClassVersion
        PDFParseJPG = False
        PDFCloseHandle in_File
        Exit Function
    End If

    in_idx = 3
    Do While in_idx < in_PEnd
        str_SegmMk = PDFIntAsHex(ArrBFile, in_idx)
        in_SegmSz = PDFIntVal(ArrBFile, in_idx + 2)

        If str_SegmMk = "FFFF" Then
            Do While ArrBFile(1, in_idx + 1) = &HFF
                in_idx = in_idx + 1
            Loop
            in_SegmSz = PDFIntVal(ArrBFile, in_idx + 2)
        End If

        Select Case str_SegmMk
            Case "FFE0"
                bChar = ArrBFile(1, in_idx + 11)
                If bChar = 0 Then
                    ArrIMG(NumberofImages + 1).in_7 = "Dots"
                ElseIf bChar = 1 Then
                    ArrIMG(NumberofImages + 1).in_7 = "Dots/inch (DPI)"
                ElseIf bChar = 2 Then
                    ArrIMG(NumberofImages + 1).in_7 = "Dots/cm"
                Else
                    MsgBox "Invalid image resolution" & bChar & _
                            "Valid resolution is: 0, 1, 2." & _
                            vbNewLine & _
                            "Cannot add image to PDF file.", vbCritical, "File Image - " & PDFClassVersion
                    PDFParseJPG = False
                    PDFCloseHandle in_File
                    Exit Function
                End If
            Case "FFC0", "FFC1", "FFC2", "FFC3", "FFC5", "FFC6", "FFC7"
                ArrIMG(NumberofImages + 1).in_1 = PDFIntVal(ArrBFile, in_idx + 7)
                ArrIMG(NumberofImages + 1).in_2 = PDFIntVal(ArrBFile, in_idx + 5)

                in_TmpColor = ArrBFile(1, in_idx + 9) * 8

                If in_TmpColor = 8 Then
                    ArrIMG(NumberofImages + 1).in_3 = "DeviceGray"
                ElseIf in_TmpColor = 24 Then
                    ArrIMG(NumberofImages + 1).in_3 = "DeviceRGB"
                ElseIf in_TmpColor = 32 Then
                    ArrIMG(NumberofImages + 1).in_3 = "DeviceCMYK"
                Else
                    ArrIMG(NumberofImages + 1).in_4 = in_TmpColor
                End If
        End Select

        in_idx = in_idx + in_SegmSz + 2
    Loop

    PDFCloseHandle in_File

    If ArrIMG(NumberofImages + 1).in_4 <> "" Then
        in_bpc = ArrIMG(NumberofImages + 1).in_4
    Else
        in_bpc = 8
        ArrIMG(NumberofImages + 1).in_4 = 8
    End If

    ArrIMG(NumberofImages + 1).in_5 = "DCTDecode"
    ArrIMG(NumberofImages + 1).in_6 = ""

    Open pFileName For Binary As #inIMG
        str_TChar = String(sIMG, " ")
        Get #inIMG, , str_TChar
        ArrIMG(NumberofImages + 1).in_6 = ArrIMG(NumberofImages + 1).in_6 & str_TChar
    Close #inIMG

    PDFParseJPG = Array(ArrIMG(NumberofImages + 1).in_1, _
                        ArrIMG(NumberofImages + 1).in_2, _
                        ArrIMG(NumberofImages + 1).in_3, _
                        in_bpc, ArrIMG(NumberofImages + 1).in_5, _
                        ArrIMG(NumberofImages + 1).in_6, _
                        ArrIMG(NumberofImages + 1).in_7, _
                        ArrIMG(NumberofImages + 1).in_8)

End Function
Private Function PDFIntAsHex(ArrBF As Variant, in_Index As Long) As String

    PDFIntAsHex = Right("00" & Hex(ArrBF(1, in_Index)), 2) & _
                  Right("00" & Hex(ArrBF(1, in_Index + 1)), 2)

End Function
Private Function PDFIntVal(ArrBF As Variant, in_idx As Long) As Long

    PDFIntVal = CLng(ArrBF(1, in_idx)) * 256& + _
                CLng(ArrBF(1, in_idx + 1))

End Function
Private Sub PDFWriteImage(in_Img As Integer)

Dim TmpImg As String

    TmpImg = ArrIMG(in_Img).in_6

    CurrentObjectNum = CurrentObjectNum + 1
    TempStream = ""

    PDFOutStream sTempStream, "%DEBUT_OBJ/%"
    PDFOutStream TempStream, CurrentObjectNum & " 0 obj"

    ImageStream = ""
    PDFOutStream ImageStream, "<</Type /XObject"
    PDFOutStream ImageStream, "/Subtype /Image"
    PDFOutStream ImageStream, "/Filter [/DCTDecode ]"
    PDFOutStream ImageStream, "/Width " & ArrIMG(in_Img).in_1
    PDFOutStream ImageStream, "/Height " & ArrIMG(in_Img).in_2
    PDFOutStream ImageStream, "/ColorSpace /" & ArrIMG(in_Img).in_3
    PDFOutStream ImageStream, "/BitsPerComponent " & ArrIMG(in_Img).in_4
    PDFOutStream ImageStream, "/Length " & Len(ArrIMG(in_Img).in_6)
    PDFOutStream ImageStream, "/Name /ImgJPEG" & in_Img & ">>"
    PDFOutStream ImageStream, "stream"
    PDFOutStream ImageStream, TmpImg
    PDFOutStream ImageStream, "endstream"
    PDFOutStream ImageStream, "endobj"
    PDFOutStream sTempStream, "%FIN_OBJ/%"
    
    TempStream = TempStream & ImageStream

    PDFAddToOffset Len(TempStream)

    Strm.WriteLine TempStream

End Sub
Public Sub PDFBeginDoc()

    FPageNumber = 1

    in_offset = 1
    
    NumberofImages = 0
    CurrentObjectNum = 0
    ObjectOffset = 0
    CurrentPDFSetPageObject = 0
    CRCounter = 0
    FontNumber = 0

    ReDim ObjectOffsetList(1 To 1)
    ReDim PageNumberList(1 To 1)
    ReDim PageCanvasHeight(1 To 1)
    ReDim PageCanvasWidth(1 To 1)

    ReDim boPageLinksList(1 To 1)
    ReDim NbPageLinksList(1 To 1)
    ReDim LinksList(1 To 1)
    ReDim FontNumberList(1 To 1)

    TempStream = ""
    ImageStream = ""

    PDFSetHeader
    PDFSetDocInfo
    PDFStartStream

End Sub
Public Sub PDFEndDoc()

Dim iRet As Long
Dim in_i As Integer

    PDFHeader
    
    PDFEndStream
    PDFSetFontType
    PDFSetPages
    PDFSetArray

    For in_i = 1 To NumberofImages
        PDFWriteImage (in_i)
    Next in_i

    For in_i = 1 To FPageNumber
        PDFSetPageObject (in_i)
    Next in_i

    PDFSetBookmarks

    PDFSetCatalog
    PDFSetXref

    Strm.WriteLine "%%EOF"
    Strm.Close

    If boPDFConfirm Then MsgBox "PDF file generated.", vbOKOnly, "Generated PDF file - " & PDFClassVersion
    If boPDFView Then
        PDFScanRepAdobe "C:\Program Files\", 0
        If wsPathAdobe <> "" Then
            iRet = Shell(wsPathAdobe & " " & PDFGetFileName, vbMaximizedFocus)
        End If
    End If
    
End Sub
Public Sub PDFEndPage()

    in_Canvas = in_Canvas + 1

    ReDim Preserve PDFCanvasWidth(1 To in_Canvas)
    ReDim Preserve PDFCanvasHeight(1 To in_Canvas)
    ReDim Preserve PDFCanvasOrientation(1 To in_Canvas)

    If PDFCanvasWidth(in_Canvas) = "" Then
        PDFCanvasWidth(in_Canvas) = PDFCanvasWidth(in_Canvas - 1)
        PDFCanvasHeight(in_Canvas) = PDFCanvasHeight(in_Canvas - 1)
        PDFCanvasOrientation(in_Canvas) = PDFCanvasOrientation(in_Canvas - 1)
    End If

    PDFHeader
    
End Sub
Public Sub PDFNewPage()

Dim TempSize As Long

    in_xCurrent = PDFlMargin
    in_yCurrent = PDFtMargin

    FPageNumber = FPageNumber + 1
    FPageLink = 0

    TempStream = TempStream & sTempStream
    If dTempStream <> "" Then TempStream = TempStream & dTempStream
    sTempStream = ""
    dTempStream = ""

    PDFOutStream TempStream, "endstream"
    PDFOutStream TempStream, "endobj"
    PDFOutStream TempStream, "%FIN_OBJ/%"
    
    StreamSize2 = 6

    PDFAddToOffset Len(TempStream)
        Strm.WriteLine TempStream

    TempSize = Len(TempStream) - StreamSize1 - StreamSize2 - Len("Stream") - Len("endstream") - 6
    ContentNum = CurrentObjectNum
    CurrentObjectNum = CurrentObjectNum + 1

    TempStream = ""

    PDFOutStream TempStream, "%DEBUT_OBJ/%"
    PDFOutStream TempStream, CurrentObjectNum & " 0 obj"
    PDFOutStream TempStream, CStr(TempSize)
    PDFOutStream TempStream, "endobj"
    PDFOutStream TempStream, "%FIN_OBJ/%"
    
    PDFAddToOffset Len(TempStream)
        Strm.WriteLine TempStream

    ContentNum = CurrentObjectNum
    CurrentObjectNum = CurrentObjectNum + 1

    TempStream = ""

    PDFOutStream sTempStream, "%DEBUT_OBJ/%"
    PDFOutStream TempStream, CurrentObjectNum & " 0 obj"
    PDFOutStream TempStream, "<< /Length " & (CurrentObjectNum + 1) & " 0 R"

    PDFOutStream TempStream, " >>"

    StreamSize1 = Len(TempStream)

    PDFOutStream TempStream, "stream"

    PDFHeader
    
End Sub
Private Sub PDFSetHeader()

    CurrentObjectNum = 0

    Strm.WriteLine "%PDF-" & PDFClass
    PDFAddToOffset Len("%PDF-" & PDFClass)

End Sub
Private Sub PDFSetDocInfo()

    CurrentObjectNum = CurrentObjectNum + 1
    TempStream = ""

    PDFOutStream sTempStream, "%DEBUT_OBJ/%"
    PDFOutStream TempStream, CurrentObjectNum & " 0 obj"
    PDFOutStream TempStream, "<<"
    PDFOutStream TempStream, "/Producer (" + FProducer + ")"
    PDFOutStream TempStream, "/Author (" + FAuthor + ")"
    PDFOutStream TempStream, "/CreationDate (D:" + Format(Now, "YYYYMMDDHHmmSS") + ")"
    PDFOutStream TempStream, "/Creator (" + FCreator + ")"
    PDFOutStream TempStream, "/Keywords (" + FKeywords + ")"
    PDFOutStream TempStream, "/Subject (" + FSubject + ")"
    PDFOutStream TempStream, "/Title (" + FTitle + ")"
    PDFOutStream TempStream, "/ModDate ()"
    PDFOutStream TempStream, ">>"
    PDFOutStream TempStream, "endobj"
    PDFOutStream sTempStream, "%FIN_OBJ/%"
    
    PDFAddToOffset Len(TempStream)
        Strm.WriteLine TempStream

End Sub
Private Sub PDFSetArray()

Dim I As Integer

    CurrentObjectNum = CurrentObjectNum + 1
    ResourceNum = CurrentObjectNum

    TempStream = ""

    PDFOutStream sTempStream, "%DEBUT_OBJ/%"
    PDFOutStream TempStream, CurrentObjectNum & " 0 obj"
    PDFOutStream TempStream, "<< /ProcSet [ /PDF /Text /ImageC]"
    PDFOutStream TempStream, "/XObject << "

    For I = 1 To NumberofImages
        PDFOutStream TempStream, "/ImgJPEG" & I & " " & (CurrentObjectNum + I) & " 0 R"
    Next I

    PDFOutStream TempStream, ">>"
    PDFOutStream TempStream, "/Font << "

    For I = 1 To FontNumber
        PDFOutStream TempStream, "/F" & I & " " & FontNumberList(I) & " 0 R "
    Next I

    PDFOutStream TempStream, ">>"
    PDFOutStream TempStream, ">>"
    PDFOutStream TempStream, "endobj"
    PDFOutStream sTempStream, "%FIN_OBJ/%"

    PDFAddToOffset Len(TempStream)
        Strm.WriteLine TempStream

End Sub
Private Sub PDFSetFontType()

Dim in_i As Integer

    For in_i = 0 To UBound(Arr_Font)
        PDFCreateFont "Type1", Arr_Font(in_i), "WinAnsiEncoding"
    Next in_i

End Sub
Private Sub PDFSetPages()

Dim I, PageObjNum As Integer

    CurrentObjectNum = CurrentObjectNum + 1
    ParentNum = CurrentObjectNum
    'TempStream = ""

    PDFOutStream TempStream, "%DEBUT_OBJ/%"
    PDFOutStream TempStream, CurrentObjectNum & " 0 obj"
    PDFOutStream TempStream, "<< /Type /Pages"
    PDFOutStream TempStream, "/Kids ["

    PageObjNum = 2
    For I = 1 To FPageNumber
        PDFOutStream TempStream, (CurrentObjectNum + I + 1 + NumberofImages) & " 0 R"

        ReDim Preserve PageNumberList(1 To in_PagesNum)
        ReDim Preserve PageCanvasHeight(1 To in_PagesNum)
        ReDim Preserve PageCanvasWidth(1 To in_PagesNum)

        ReDim Preserve boPageLinksList(1 To FPageNumber)
        ReDim Preserve NbPageLinksList(1 To FPageNumber)

        PageCanvasHeight(in_PagesNum) = PDFCanvasHeight(in_PagesNum)
        PageCanvasWidth(in_PagesNum) = PDFCanvasWidth(in_PagesNum)

        PageNumberList(in_PagesNum) = PageObjNum
        in_PagesNum = in_PagesNum + 1

        PageObjNum = PageObjNum + 2
    Next I

    PDFOutStream TempStream, "]"
    PDFOutStream TempStream, "/Count " & FPageNumber
    PDFOutStream TempStream, ">>"
    PDFOutStream TempStream, "endobj"
    PDFOutStream sTempStream, "%FIN_OBJ/%"
    
    PDFAddToOffset Len(TempStream)
        Strm.WriteLine TempStream

End Sub
Private Sub PDFSetPageObject(in_pg As Integer)

Dim I             As Integer
Dim str_Rect      As String
Dim str_Annots    As String
Dim str_TmpAnnots As String

    ContentNum = ContentNum + 1
    CurrentObjectNum = CurrentObjectNum + 1
    TempStream = ""

    ReDim Preserve aPage(1 To in_pg)
    aPage(in_pg) = CurrentObjectNum
    
    PDFOutStream sTempStream, "%DEBUT_OBJ/%"
    PDFOutStream TempStream, CurrentObjectNum & " 0 obj"
    PDFOutStream TempStream, "<< /Type /Page"
    PDFOutStream TempStream, "/Parent " & ParentNum & " 0 R"
    PDFOutStream TempStream, "/MediaBox [ 0 0 " & PageCanvasWidth(CurrentPDFSetPageObject + 1) & " " & PageCanvasHeight(CurrentPDFSetPageObject + 1) & "]"
    PDFOutStream TempStream, "/Resources " & ResourceNum & " 0 R"

    If boPageLinksList(in_pg) = True Then
        str_Annots = "/Annots ["
        For I = 1 To NbPageLinksList(in_pg)
            str_Rect = ""
            str_Rect = PageLinksList(in_pg, I)(0) & " " & _
                      PageLinksList(in_pg, I)(1) & " " & _
                      PageLinksList(in_pg, I)(0) + PageLinksList(in_pg, I)(2) & " " & _
                      PageLinksList(in_pg, I)(1) - PageLinksList(in_pg, I)(3)
            str_Annots = str_Annots & "<</Type /Annot /Subtype /Link /Rect [" & str_Rect & "] /Border [0 0 0] "

            If TypeName(PageLinksList(in_pg, I)(4)) = "String" And PageLinksList(in_pg, I)(4) <> "" Then
                str_TmpAnnots = PageLinksList(in_pg, I)(4)
                
                str_TmpAnnots = Replace(str_TmpAnnots, "\", "\\")
                str_TmpAnnots = Replace(str_TmpAnnots, "\\", "\\\\")
                str_TmpAnnots = Replace(str_TmpAnnots, "(", "\(")
                str_TmpAnnots = Replace(str_TmpAnnots, ")", "\)")
    
                str_Annots = str_Annots & "/A <</S /URI /URI (" & str_TmpAnnots & ")>>>>" & vbCr & vbLf
            End If
        Next I

        PDFOutStream TempStream, str_Annots & "]"
        'MsgBox str_Annots
    End If

    PDFOutStream TempStream, "/Contents " & PageNumberList(CurrentPDFSetPageObject + 1) & " 0 R"
    PDFOutStream TempStream, ">>"
    PDFOutStream TempStream, "endobj"
    PDFOutStream TempStream, "%FIN_OBJ/%"
    
    PDFAddToOffset Len(TempStream)
        Strm.WriteLine TempStream
        
    CurrentPDFSetPageObject = CurrentPDFSetPageObject + 1
    
End Sub
Private Sub PDFSetCatalog()

    CurrentObjectNum = CurrentObjectNum + 1
    CatalogNum = CurrentObjectNum
    TempStream = ""

    PDFOutStream sTempStream, "%DEBUT_OBJ/%"
    PDFOutStream TempStream, CurrentObjectNum & " 0 obj"
    PDFOutStream TempStream, "<<"
    PDFOutStream TempStream, "/Type /Catalog"
    PDFOutStream TempStream, "/Pages " & ParentNum & " 0 R"
    
    If PDFZoomMode = ZOOM_FULLPAGE Then
        PDFOutStream TempStream, "/OpenAction [3 0 R /Fit]"
    ElseIf PDFZoomMode = ZOOM_FULLWIDTH Then
        PDFOutStream TempStream, "/OpenAction [3 0 R /FitH null]"
    ElseIf PDFZoomMode = ZOOM_REAL Then
        PDFOutStream TempStream, "/OpenAction [3 0 R /XYZ null null 1]"
    ElseIf IsNumeric(PDFZoomMode) Then
        PDFOutStream TempStream, "/OpenAction [3 0 R /XYZ null null " & PDFFormatDouble(PDFZoomMode / 100) & "]"
    End If

    If PDFLayoutMode = LAYOUT_SINGLE Then
        PDFOutStream TempStream, "/PageLayout /SinglePage"
    ElseIf PDFLayoutMode = LAYOUT_CONTINOUS Then
        PDFOutStream TempStream, "/PageLayout /OneColumn"
    ElseIf PDFLayoutMode = LAYOUT_TWO Then
        PDFOutStream TempStream, "/PageLayout /TwoColumnLeft"
    End If

    If PDFboThumbs = True Then
        PDFOutStream TempStream, "/PageMode /UseThumbs"
    End If
    
    If PDFboOutlines = True Then
        PDFOutStream TempStream, "/Outlines " & iOutlines & " 0 R"
        PDFOutStream TempStream, "/PageMode /UseOutlines"
    End If
    
    If bPDFViewerPref Then
        PDFOutStream TempStream, "/ViewerPreferences<<"
        If InStr(1, PDFViewerPref, VIEW_HIDEMENUBAR) <> 0 Then PDFOutStream TempStream, "/HideMenubar true"
        If InStr(1, PDFViewerPref, VIEW_HIDETOOLBAR) <> 0 Then PDFOutStream TempStream, "/HideToolbar true"
        If InStr(1, PDFViewerPref, VIEW_HIDEWINDOWUI) <> 0 Then PDFOutStream TempStream, "/HideWindowUI true"
        If InStr(1, PDFViewerPref, VIEW_DISPLAYDOCTITLE) <> 0 Then PDFOutStream TempStream, "/DisplayDocTitle true"
        If InStr(1, PDFViewerPref, VIEW_CENTERWINDOW) <> 0 Then PDFOutStream TempStream, "/CenterWindow true"
        If InStr(1, PDFViewerPref, VIEW_FITWINDOW) <> 0 Then PDFOutStream TempStream, "/FitWindow true"
        PDFOutStream TempStream, ">>"
    End If
    
    PDFOutStream TempStream, ">>"
    PDFOutStream TempStream, "endobj"
    PDFOutStream sTempStream, "%FIN_OBJ/%"
    
    PDFAddToOffset Len(TempStream)
        Strm.WriteLine TempStream

End Sub
Private Sub PDFStartStream()

    ContentNum = CurrentObjectNum
    CurrentObjectNum = CurrentObjectNum + 1

    TempStream = ""

    PDFOutStream sTempStream, "%DEBUT_OBJ/%"
    PDFOutStream TempStream, CurrentObjectNum & " 0 obj"
    PDFOutStream TempStream, "<< /Length " & (CurrentObjectNum + 1) & " 0 R"
    PDFOutStream TempStream, " >>"

    StreamSize1 = Len(TempStream)

    PDFOutStream TempStream, "stream"
    sTempStream = ""
    dTempStream = ""

End Sub
Private Sub PDFEndStream()

Dim TempSize As Long

    TempStream = TempStream & sTempStream
    If dTempStream <> "" Then TempStream = TempStream & dTempStream
    sTempStream = ""
    dTempStream = ""

    PDFOutStream TempStream, "endstream"
    PDFOutStream TempStream, "endobj"
    PDFOutStream sTempStream, "%FIN_OBJ/%"
    
    StreamSize2 = 6

    PDFAddToOffset Len(TempStream)
        Strm.WriteLine TempStream

    TempSize = Len(TempStream) - StreamSize1 - StreamSize2 - Len("Stream") - Len("endstream") - 6
    ContentNum = CurrentObjectNum
    CurrentObjectNum = CurrentObjectNum + 1
    TempStream = ""

    PDFOutStream sTempStream, "%DEBUT_OBJ/%"
    PDFOutStream TempStream, CurrentObjectNum & " 0 obj"
    PDFOutStream TempStream, CStr(TempSize)
    PDFOutStream TempStream, "endobj"
    PDFOutStream sTempStream, "%FIN_OBJ/%"
    
    PDFAddToOffset Len(TempStream)
        Strm.WriteLine TempStream

End Sub
Private Sub PDFSetXref()

Dim I As Integer

    CurrentObjectNum = CurrentObjectNum + 1
    TempStream = ""

    PDFOutStream TempStream, "xref"
    PDFOutStream TempStream, "0 " & CurrentObjectNum
    PDFOutStream TempStream, "0000000000 65535 f"

    For I = 1 To CurrentObjectNum - 1
        PDFOutStream TempStream, PDFGetOffsetNumber(Trim(ObjectOffsetList(I))) + " 00000 n"
    Next I

    PDFOutStream TempStream, "trailer"
    PDFOutStream TempStream, "<< /Size " & CurrentObjectNum
    PDFOutStream TempStream, "/Root " & CatalogNum & " 0 R"
    PDFOutStream TempStream, "/Info 1 0 R"
    PDFOutStream TempStream, ">>"
    PDFOutStream TempStream, "startxref"
    PDFOutStream TempStream, Trim(ObjectOffsetList(CurrentObjectNum))

    Strm.WriteLine TempStream

End Sub
Private Function PDFUnderline(boCell As Boolean, str_Text As String, X As Double, Y As Double) As String

Dim in_wUp          As Integer
Dim in_wUt          As Integer
Dim in_wTxt         As String

Dim in_Px           As Integer
Dim in_Pw           As String
Dim in_Py           As String

Dim str_TmpUnderl   As String

Dim str_xLeft       As String
Dim str_yTop        As String
Dim str_wText       As String
Dim str_hLine       As String
Dim iNbSpace        As Integer

    str_TmpUnderl = ""

    in_wUp = PDFGetStringWidth("up", PDFFontName, PDFFontSize)
    in_wUt = 2

    iNbSpace = PDFGetNumberOfCar(str_Text, " ")
    in_wTxt = PDFGetStringWidth(str_Text, PDFFontName, PDFFontSize) + _
              iNbSpace * PDFGetStringWidth(" ", PDFFontName, PDFFontSize) + _
              iWidthStr * iNbSpace - _
              IIf(iWidthStr <> 0, (iNbSpace + 1) * PDFcMargin, 0)

    in_Px = X + PDFlMargin * in_Ech
    in_Pw = (PDFCanvasHeight(in_Canvas) - (Y - in_wUp / 1000 * PDFFontSize) - 2)
    in_Py = -in_wUt / 1000 * in_wTxt
    str_hLine = PDFFormatDouble(in_Py)

    If boCell = False Then
        str_wText = PDFFormatDouble(in_wTxt)
        str_xLeft = PDFFormatDouble(in_Px)
        str_yTop = PDFFormatDouble(in_Pw)

        str_TmpUnderl = str_xLeft & " " & str_yTop & " " & str_wText & " " & str_hLine & " re f"
    Else
        str_wText = PDFFormatDouble(in_wTxt - PDFcMargin)
        str_xLeft = PDFFormatDouble(X)
        str_yTop = PDFFormatDouble(Y - 3)
        
        str_TmpUnderl = str_xLeft & " " & str_yTop & " " & str_wText & " " & str_hLine & " re f"
    End If

    PDFUnderline = str_TmpUnderl

End Function
Private Sub PDFCreateFont(Subtype, BaseFont, Encoding As String)

    FontNumber = FontNumber + 1
    CurrentObjectNum = CurrentObjectNum + 1

    ReDim Preserve FontNumberList(1 To in_FontNum)
    FontNumberList(in_FontNum) = CurrentObjectNum
    in_FontNum = in_FontNum + 1

    TempStream = ""

    PDFOutStream sTempStream, "%DEBUT_OBJ/%"
    PDFOutStream TempStream, CurrentObjectNum & " 0 obj"
    PDFOutStream TempStream, "<< /Type /Font"
    PDFOutStream TempStream, "/Subtype /" & Subtype
    PDFOutStream TempStream, "/Name /F" & FontNumber
    PDFOutStream TempStream, "/BaseFont /" & BaseFont
    PDFOutStream TempStream, "/Encoding /" + Encoding
    PDFOutStream TempStream, ">>"
    PDFOutStream TempStream, "endobj"
    PDFOutStream sTempStream, "%FIN_OBJ/%"
    
    PDFAddToOffset Len(TempStream)
        Strm.WriteLine TempStream

End Sub
Private Function PDFGetOffsetNumber(offset As String) As String
Dim X, Y As Long

    X = Len(offset)
    For Y = 1 To 10 - X
        PDFGetOffsetNumber = PDFGetOffsetNumber + "0"
    Next Y

    PDFGetOffsetNumber = PDFGetOffsetNumber + offset

End Function
Private Sub PDFOutStream(ms As String, S As String)

    CRCounter = CRCounter + 2
    ms = ms & S & vbCrLf

End Sub
Private Sub PDFAddToOffset(offset As Long)

    ReDim Preserve ObjectOffsetList(1 To in_offset)

    ObjectOffset = ObjectOffset + offset
    ObjectOffsetList(in_offset) = ObjectOffset

    in_offset = in_offset + 1

    CRCounter = 0

End Sub
Public Function PDFGetStringWidth(str_Txt As String, Optional str_FName As String, Optional in_FSize As Integer) As Double

Dim str_TmpINI As String
Dim in_Tmp     As Long
Dim in_i       As Integer
Dim in_j       As Integer
Dim ArrFNT()   As Integer
Dim in_Asc     As Integer
Dim Fso        As Object
Dim f          As Object
Dim aTempFNT   As Variant
Dim bWX        As Boolean
Dim iAscMin    As Integer
Dim iAscMax    As Integer
Dim aAsc       As Variant
Dim aWX        As Variant
Dim sReadLine  As String

    If str_FName = "" Then
        str_FName = PDFFontName
    End If
    
    ReDim ArrFNT(1 To 255)
    
    iAscMin = 0
    iAscMax = 0
    
    bWX = False
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set f = Fso.OpenTextFile(wsPathConfig & "\" & str_FName & ".font", 1, 0)
        Do While f.AtEndOfStream <> True
            sReadLine = f.ReadLine
            
            If InStr(1, sReadLine, "StartCharMetrics") <> 0 Then
                bWX = True
                sReadLine = f.ReadLine
            End If
            
            If InStr(1, sReadLine, "-1 ;") <> 0 Or _
               InStr(1, sReadLine, "EndCharMetrics") <> 0 Then
                        iAscMax = aAsc(1)
                    Exit Do
            End If
            
            If bWX = True Then
                aTempFNT = Split(sReadLine, ";")
                    aAsc = Split(Trim(aTempFNT(0)), " ")
                    If iAscMin = 0 Then iAscMin = aAsc(1)
                    
                    aWX = Split(Trim(aTempFNT(1)), " ")
                    
                    ArrFNT(aAsc(1)) = Int(aWX(1))
            End If
        Loop
    f.Close

    For in_i = 1 To 255
        If in_i < iAscMin Then ArrFNT(in_i) = 0
        If in_i > iAscMax Then ArrFNT(in_i) = 0
    Next in_i

    in_Tmp = 0
    For in_i = 1 To Len(str_Txt)
        in_Asc = Asc(Mid(str_Txt, in_i, 1))
        in_Tmp = in_Tmp + Int(ArrFNT(in_Asc)) ' + FontBBoxAbout
    Next in_i

    PDFGetStringWidth = (in_Tmp * in_FSize) / 1000

End Function
Private Function PDFGetRGB(lColor As Long) As PDFRGB

With PDFGetRGB
    .in_b = CByte(Int(lColor / 65536))
    .in_g = CByte(Int((lColor - CLng(.in_b) * 65536) / 256))
    .in_r = CByte(lColor - CLng(.in_b) * 65536 - CLng(.in_g) * 256)
End With

End Function
Private Function PDFFormatDouble(in_dbl As Variant, Optional nZero As Integer = 2) As String

Dim sZero As String

    sZero = String(nZero, "0")
    PDFFormatDouble = Replace(Format(in_dbl, "###0." & sZero), ",", ".")

End Function
Private Sub Class_Initialize()

    PDFInit

End Sub
Property Let PDFLoadFont(sPathFont As String)

Dim Fso     As Object
Dim oRep    As Object
Dim oFiles  As Object
Dim in_Font As Integer

    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set oRep = Fso.GetFolder(sPathFont)
        
    in_Font = -1
    For Each oFiles In oRep.Files
        If InStr(1, LCase(oFiles.Path), ".font") <> 0 Then
            in_Font = in_Font + 1
            ReDim Preserve Arr_Font(0 To in_Font)
                Arr_Font(in_Font) = Mid(oFiles.Name, 1, Len(oFiles.Name) - 4)
        End If
    Next oFiles
    
    If in_Font <> -1 Then wsPathConfig = sPathFont
    
End Property
Private Function PDFScanRepAdobe(sPathBegin As String, iIndexFolder As Long) As Boolean

Dim Fso     As Object
Dim oRep    As Object
Dim oSubRep As Object
Dim oFolder As Object
Dim oFiles  As Object

    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set oRep = Fso.GetFolder(sPathBegin)
    
    For Each oFolder In oRep.SubFolders
        iIndexFolder = iIndexFolder + 1
    
        If oFolder.Attributes <> 22 Then
            For Each oFiles In oFolder.Files
                If InStr(1, oFiles.Path, "AcroRd32.exe") <> 0 Then
                    wsPathAdobe = oFiles.Path
                    bScanAdobe = True
                    Exit For
                End If
            Next oFiles
        End If
        
        If bScanAdobe = True Then Exit For
    Next oFolder

    For Each oSubRep In oRep.SubFolders
        If bScanAdobe = True Then Exit For
        PDFScanRepAdobe oSubRep.Path, iIndexFolder
    Next oSubRep
   
    Set Fso = Nothing
    If bScanAdobe = True Then Exit Function
    
End Function
Public Sub PDFInit()
    
    bScanAdobe = False
    Set Fso = CreateObject("scripting.filesystemobject")
    
    If wsPathConfig = "" Then wsPathConfig = App.Path
    PDFLoadFont = wsPathConfig
    
    ObjectOffsetList = Array()
    PageNumberList = Array()
    PageCanvasWidth = Array()
    PageCanvasHeight = Array()

    boPageLinksList = Array()
    NbPageLinksList = Array()
    LinksList = Array()

    FontNumberList = Array()

    in_offset = 1
    in_FontNum = 1
    in_PagesNum = 1
    in_Canvas = 1
    FPageLink = 0

    boPDFUnderline = False
    boPDFBold = False
    boPDFItalic = False

    ' Unité de mesure par défaut : cm
        in_Ech = 72 / 2.54

    ' Marges de la page (1 cm)
    PDFMargin = in_Ech / 28.35
    PDFSetMargins PDFMargin, PDFMargin

    ' Marge interieure des cellules (1 mm)
    PDFcMargin = in_Ech * (PDFMargin / 10)

    ' Largeur de ligne (0.2 mm)
    PDFLnWidth = 0.567

    in_xCurrent = PDFlMargin
    in_yCurrent = PDFtMargin

    TempStream = ""
    ImageStream = ""
    pTempStream = ""
    sTempStream = ""
    cTempStream = ""
    dTempStream = ""

    FontNum = 1

    ' Définition dzes couleurs par défaut
        PDFLineColor = "0 G"
        PDFDrawColor = "0 g"
        PDFTextColor = "0 g"

    ' Format d'orientation de page par défaut : A4
        ReDim Preserve PDFCanvasWidth(1 To in_Canvas)
        ReDim Preserve PDFCanvasHeight(1 To in_Canvas)
        ReDim Preserve PDFCanvasOrientation(1 To in_Canvas)

        PDFCanvasWidth(in_Canvas) = 595.28
        PDFCanvasHeight(in_Canvas) = 841.89
        PDFCanvasOrientation(in_Canvas) = "p"

    FProducer = ""
    FAuthor = ""
    FCreator = ""

    FKeywords = ""
    FSubject = ""

    Exit Sub
    
End Sub
Function PDFSetBookmark(str_Text As String, Optional iLevel As Integer = 0, Optional Y As Double = -1)

    If Y = -1 Then Y = in_yCurrent
    
    ReDim Preserve aOutlines(0 To iOutlines)
    
    aOutlines(iOutlines).sText = str_Text
    aOutlines(iOutlines).iLevel = iLevel
    aOutlines(iOutlines).yPos = Y
    aOutlines(iOutlines).iPageNb = PDFPageNumber

    iOutlines = iOutlines + 1
    
End Function
Private Function PDFSetBookmarks()

Dim iNbBookMrk  As Integer
Dim aTemp()     As Variant
Dim iLevel      As Integer
Dim in_i        As Integer
Dim iParent     As Integer
Dim iFirst      As Integer
Dim iPrev       As Integer
Dim iNb         As Integer
Dim iPageOut    As Integer

    On Error Resume Next
    iNbBookMrk = UBound(aOutlines)
    If iNbBookMrk = 0 Then Exit Function
    On Error GoTo 0

    iLevel = 0
    For in_i = 0 To iNbBookMrk
        If aOutlines(in_i).iLevel > 0 Then
            iParent = aTemp(aOutlines(in_i).iLevel - 1)

            aOutlines(in_i).iParent = iParent
            aOutlines(iParent).iLast = in_i
            aOutlines(iParent).bLast = True
            
            If aOutlines(in_i).iLevel > iLevel Then
                aOutlines(iParent).iFirst = in_i
                aOutlines(iParent).bFirst = True
            End If
        Else
            aOutlines(in_i).iParent = iNbBookMrk + 1
        End If
        
        If aOutlines(in_i).iLevel <= iLevel And in_i > 1 Then
            iPrev = aTemp(aOutlines(in_i).iLevel)
            aOutlines(iPrev).iNext = in_i
            aOutlines(iPrev).bNext = True
            aOutlines(in_i).iPrev = iPrev
            aOutlines(in_i).bPrev = True
        End If
        
        ReDim Preserve aTemp(0 To aOutlines(in_i).iLevel)
        aTemp(aOutlines(in_i).iLevel) = in_i
        iLevel = aOutlines(in_i).iLevel
    Next in_i
    
    iNb = CurrentObjectNum + 1
    iOutlineRoot = iNb
    For in_i = 0 To iNbBookMrk
        CurrentObjectNum = CurrentObjectNum + 1
        TempStream = ""
        
        PDFOutStream sTempStream, "%DEBUT_OBJ/%"
        PDFOutStream TempStream, CurrentObjectNum & " 0 obj"
        PDFOutStream TempStream, "<</Title (" & aOutlines(in_i).sText & ")"
        PDFOutStream TempStream, "/Parent " & (iNb + aOutlines(in_i).iParent) & " 0 R"
        
        If aOutlines(in_i).bPrev Then
            PDFOutStream TempStream, "/Prev " & (iNb + aOutlines(in_i).iPrev) & " 0 R"
        End If
        If aOutlines(in_i).bNext Then
            PDFOutStream TempStream, "/Next " & (iNb + aOutlines(in_i).iNext) & " 0 R"
        End If
        If aOutlines(in_i).bFirst Then
            PDFOutStream TempStream, "/First " & (iNb + aOutlines(in_i).iFirst) & " 0 R"
        End If
        If aOutlines(in_i).bLast Then
            PDFOutStream TempStream, "/Last " & (iNb + aOutlines(in_i).iLast) & " 0 R"
        End If
        
        iPageOut = aPage(aOutlines(in_i).iPageNb)
        
        PDFOutStream TempStream, "/Dest [" & iPageOut & _
                                 " 0 R /XYZ 0 " & PDFFormatDouble(PDFCanvasHeight(aOutlines(in_i).iPageNb) - aOutlines(in_i).yPos * in_Ech) & " null]"
        PDFOutStream TempStream, "/Count 0>>"
        PDFOutStream TempStream, "endobj"
        PDFOutStream sTempStream, "%FIN_OBJ/%"
    
        PDFAddToOffset Len(TempStream)
            Strm.WriteLine TempStream
    Next in_i
    
    CurrentObjectNum = CurrentObjectNum + 1
    TempStream = ""
    iOutlines = CurrentObjectNum
    
    PDFOutStream sTempStream, "%DEBUT_OBJ/%"
    PDFOutStream TempStream, CurrentObjectNum & " 0 obj"

    PDFOutStream TempStream, "<</Type /Outlines /First " & iNb & " 0 R"
    PDFOutStream TempStream, "/Last " & (iNb + aTemp(1)) & " 0 R>>"
    PDFOutStream TempStream, "endobj"
    PDFOutStream sTempStream, "%FIN_OBJ/%"
    
    PDFAddToOffset Len(TempStream)
        Strm.WriteLine TempStream
            
End Function


