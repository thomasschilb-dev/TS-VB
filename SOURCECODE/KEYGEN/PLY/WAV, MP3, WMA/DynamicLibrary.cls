VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynamicLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CallWindowProc Lib "user32" _
Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Private Declare Function LoadLibrary Lib "kernel32" _
Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As String _
) As Long

Private Declare Function FreeLibrary Lib "kernel32" ( _
    ByVal hModule As Long _
) As Long

Private Declare Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As Long, _
    ByVal lpProcName As String _
) As Long

Private Declare Sub CpyMem Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal dwLen As Long _
)

Private hMod        As Long
Private blnIsCDECL  As Boolean

Public Property Get ModuleHandle( _
) As Long

    ModuleHandle = hMod
End Property

Public Property Get IsCDECL( _
) As Boolean

    IsCDECL = blnIsCDECL
End Property

Public Property Let IsCDECL( _
    bln As Boolean _
)

    blnIsCDECL = bln
End Property

Public Sub UnloadDLL()
    FreeLibrary hMod
    hMod = 0
End Sub

Public Function LoadDLL( _
    ByVal strDLL As String _
) As Boolean

    blnIsCDECL = False
    hMod = LoadLibrary(strDLL)
    LoadDLL = hMod <> 0
End Function

Public Function CallFunc( _
    ByVal fnc As String, _
    ParamArray args() As Variant _
) As Long

    Dim hFnc    As Long

    hFnc = GetProcAddress(hMod, fnc)
    If hFnc = 0 Then
        Err.Raise 1, , "Export nicht gefunden!"
        Exit Function
    End If

    If Not IsCDECL Then
        CallFunc = CallStd(hFnc, args)
    Else
        CallFunc = CallCdecl(hFnc, args)
    End If
End Function

Private Function CallStd( _
    ByVal fnc As Long, _
    ParamArray Params() As Variant _
) As Long

    Dim btASM(&HEC00& - 1)  As Byte
    Dim pASM                As Long
    Dim i                   As Integer

    pASM = VarPtr(btASM(0))

    AddByte pASM, &H58                  ' POP EAX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H50                  ' PUSH EAX

    If UBound(Params) = 0 Then
        If IsArray(Params(0)) Then
            For i = UBound(Params(0)) To 0 Step -1
                AddPush pASM, CLng(Params(0)(i))    ' PUSH dword
            Next
        Else
            For i = UBound(Params) To 0 Step -1
                AddPush pASM, CLng(Params(i))       ' PUSH dword
            Next
        End If
    Else
        For i = UBound(Params) To 0 Step -1
            AddPush pASM, CLng(Params(i))           ' PUSH dword
        Next
    End If

    AddCall pASM, fnc                   ' CALL rel addr
    AddByte pASM, &HC3                  ' RET

    CallStd = CallWindowProc(VarPtr(btASM(0)), _
                             0, 0, 0, 0)
End Function

' http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=62014&lngWId=1
Private Function CallCdecl( _
    ByVal lpfn As Long, _
    ParamArray args() As Variant _
) As Long

    Dim btASM(&HEC00& - 1)  As Byte
    Dim pASM                As Long
    Dim btArgSize           As Byte
    Dim i                   As Integer

    pASM = VarPtr(btASM(0))

    If UBound(args) = 0 Then
        If IsArray(args(0)) Then
            For i = UBound(args(0)) To 0 Step -1
                AddPush pASM, CLng(args(0)(i))    ' PUSH dword
                btArgSize = btArgSize + 4
            Next
        Else
            For i = UBound(args) To 0 Step -1
                AddPush pASM, CLng(args(i))       ' PUSH dword
                btArgSize = btArgSize + 4
            Next
        End If
    Else
        For i = UBound(args) To 0 Step -1
            AddPush pASM, CLng(args(i))           ' PUSH dword
            btArgSize = btArgSize + 4
        Next
    End If

    AddByte pASM, &HB8
    AddLong pASM, lpfn
    AddByte pASM, &HFF
    AddByte pASM, &HD0
    AddByte pASM, &H83
    AddByte pASM, &HC4
    AddByte pASM, btArgSize
    AddByte pASM, &HC2
    AddByte pASM, &H10
    AddByte pASM, &H0

    CallCdecl = CallWindowProc(VarPtr(btASM(0)), _
                               0, 0, 0, 0)
End Function

Private Sub AddPush( _
    pASM As Long, _
    lng As Long _
)

    AddByte pASM, &H68
    AddLong pASM, lng
End Sub

Private Sub AddCall( _
    pASM As Long, _
    addr As Long _
)

    AddByte pASM, &HE8
    AddLong pASM, addr - pASM - 4
End Sub

Private Sub AddLong( _
    pASM As Long, _
    lng As Long _
)

    CpyMem ByVal pASM, lng, 4
    pASM = pASM + 4
End Sub

Private Sub AddByte( _
    pASM As Long, _
    Bt As Byte _
)

    CpyMem ByVal pASM, Bt, 1
    pASM = pASM + 1
End Sub

Private Sub Class_Terminate()
    'UnloadDLL
End Sub
