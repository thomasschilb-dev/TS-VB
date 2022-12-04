VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FileDlg2.cls

' (Modified from vbAccelerator.com)

'  Use eg
'  On Form1:=


'  LOAD
'   Title$ = "Load a picture file"
'   Filt$ = "Pics bmp,jpg,gif,ico,cur,wmf,emf|*.bmp;*.jpg;*.gif;*.ico;*.cur;*.wmf;*.emf"
'   InDir$ = CurrPath$ 'Pathspec$
'   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd
'
'  SAVE
'   Title$ = "Save Mask as 2-color bmp"
'   Filt$ = "Save bmp|*.bmp"
'   InDir$ = CurrPath$ 'Pathspec$
'   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd
'
'   Len(FileSpec$)=0 for cancel

Option Explicit

Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" _
    (ByVal lpString As String) As Long

Private Const MAX_PATH = 260
Private Const MAX_FILE = 260

Private Type OPENFILENAME
    lStructSize As Long          ' UDT length
    hWndOwner As Long            ' Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Filter
    lpstrCustomFilter As String  ' Ignored
    nMaxCustFilter As Long       ' Ignored
    nFilterIndex As Long         ' FilterIndex
    lpstrFile As String          ' FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' InitDir
    lpstrTitle As String         ' Dialog Title
    flags As Long                ' Flags
    nFileOffset As Integer       ' Ignored
    nFileExtension As Integer    ' Ignored
    lpstrDefExt As String        ' DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored
    lpTemplateName As Long       ' Ignored
End Type

Public Enum ENUMOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000&
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONE2RKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

Private Declare Function GetOpenFileName Lib "COMDLG32" Alias "GetOpenFileNameA" _
    (file As OPENFILENAME) As Long

Private Declare Function GetSaveFileName Lib "COMDLG32" Alias "GetSaveFileNameA" _
    (file As OPENFILENAME) As Long

Dim m_lExtendedError As Long

' If parameter MultiSelect is True, dialog will be Drive/File box style

Function ShowOpen(Optional FileName As String, _
                  Optional DlgTitle As String, _
                  Optional Filter As String = "All (*.*)| *.*", _
                  Optional InitDir As String, _
                  Optional DefaultExt As String = "", _
                  Optional owner As Long = -1, _
                  Optional FileTitle As String, _
                  Optional FileMustExist As Boolean = True, _
                  Optional MultiSelect As Boolean = False, _
                  Optional ReadOnly As Boolean = False, _
                  Optional HideReadOnly As Boolean = False, _
                  Optional FilterIndex As Long = 1, _
                  Optional flags As Long = 0) As String

Dim typOpenFile As OPENFILENAME
Dim s As String
Dim lFlags As Long
Dim ch As String
Dim i As Integer
Dim mResult

m_lExtendedError = 0

With typOpenFile
    .lStructSize = Len(typOpenFile)

     ' Add in specific flags and strip out non-VB flags
    .flags = (-FileMustExist * OFN_FILEMUSTEXIST) Or _
         (-MultiSelect * OFN_ALLOWMULTISELECT) Or _
         (-ReadOnly * OFN_READONLY) Or _
         (-HideReadOnly * OFN_HIDEREADONLY) Or _
         (.flags And CLng(Not (OFN_ENABLEHOOK Or OFN_ENABLETEMPLATE)))
    If owner <> -1 Then .hWndOwner = owner
    .lpstrInitialDir = InitDir
    .lpstrDefExt = DefaultExt
    .lpstrTitle = DlgTitle

    ' To make Windows-style filter, replace | and : with nulls
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
             s = s & vbNullChar
        Else
             s = s & ch
        End If
    Next

    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    s = FileName & String$(MAX_PATH - Len(FileName), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE

    mResult = GetOpenFileName(typOpenFile)
    If mResult = 1 Then
         FileName = StrZToStr(.lpstrFile)
        '  If you initiate the variables,
        '  you can return the value(s)
        'FileTitle = StrZToStr(.lpstrFileTitle)
    Else
         FileName = ""
         If mResult <> 0 Then    ' 0 is Cancel, else extended error
              m_lExtendedError = CommDlgExtendedError()
         End If
    End If
End With
ShowOpen = FileName
End Function

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
Dim iStart As Long
Dim iEnd As Long
Dim s As String

iStart = 1
If sFilters = "" Then Exit Function
Do
    ' Cut out both parts marked by null character
    iEnd = InStr(iStart, sFilters, vbNullChar)
    If iEnd = 0 Then Exit Function
    iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
    If iEnd Then
        s = Mid$(sFilters, iStart, iEnd - iStart)
    Else
        s = Mid$(sFilters, iStart)
    End If
    iStart = iEnd + 1
    If iCur = 1 Then
        FilterLookup = s
        Exit Function
    End If
    iCur = iCur - 1
Loop While iCur
End Function

Private Function StrZToStr(s As String) As String
    StrZToStr = Left$(s, lstrlen(s))
End Function

            
Function ShowSave(Optional FileName As String, _
                  Optional DlgTitle As String, _
                  Optional Filter As String = "All (*.*)| *.*", _
                  Optional InitDir As String, _
                  Optional DefaultExt As String, _
                  Optional owner As Long = -1, _
                  Optional FileTitle As String, _
                  Optional OverWritePrompt As Boolean = True, _
                  Optional FilterIndex As Long = 1, _
                  Optional flags As Long) As String
            
Dim typOpenFile As OPENFILENAME
Dim s As String
Dim ch As String
Dim i As Integer
Dim mResult

m_lExtendedError = 0

With typOpenFile
    .lStructSize = Len(typOpenFile)

    ' Add in specific flags and strip out non-VB flags
    .flags = (-OverWritePrompt * OFN_OVERWRITEPROMPT) Or _
     OFN_HIDEREADONLY Or _
     (flags And CLng(Not (OFN_ENABLEHOOK Or OFN_ENABLETEMPLATE)))
    If owner <> -1 Then .hWndOwner = owner
    .lpstrInitialDir = InitDir
    .lpstrDefExt = DefaultExt
    .lpstrTitle = DlgTitle

    ' Make new filter with bars (|) replacing nulls
    ' and double null at end
    For i = 1 To Len(Filter)
         ch = Mid$(Filter, i, 1)
         If ch = "|" Or ch = ":" Then
              s = s & vbNullChar
         Else
              s = s & ch
         End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    s = FileName & String$(MAX_PATH - Len(FileName), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    ' All other fields zero

    mResult = GetSaveFileName(typOpenFile)
    If mResult = 1 Then
         FileName = StrZToStr(.lpstrFile)
         '  If you initiate the variables,
         '  you can return the value(s)
         'FileTitle = StrZToStr(.lpstrFileTitle)
    Else
         FileName = ""
         If mResult <> 0 Then   ' 0 is Cancel, else extended error
              m_lExtendedError = CommDlgExtendedError()
         End If
    End If
End With
ShowSave = FileName
End Function

