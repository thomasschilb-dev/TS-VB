Attribute VB_Name = "modCustomConstructor"
Option Explicit

' from Paul Catons Sub Classing Code
' http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=64867&lngWId=1

Private Declare Function CallWindowProcA Lib "user32" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Private Declare Function IsBadCodePtr Lib "kernel32" ( _
    ByVal lpfn As Long _
) As Long

Private Declare Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As Long, _
    ByVal dwSize As Long, _
    ByVal flAllocationType As Long, _
    ByVal flProtect As Long _
) As Long

Private Declare Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As Long, _
    ByVal dwSize As Long, _
    ByVal dwFreeType As Long _
) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
    ByVal Destination As Long, _
    ByVal Source As Long, _
    ByVal Length As Long _
)

Private Enum VirtualFreeTypes
    MEM_DECOMMIT = &H4000
    MEM_RELEASE = &H8000
End Enum

Private Enum VirtualAllocTypes
    MEM_COMMIT = &H1000
    MEM_RESERVE = &H2000
    MEM_RESET = &H8000
    MEM_LARGE_PAGES = &H20000000
    MEM_PHYSICAL = &H100000
    MEM_WRITE_WATCH = &H200000
End Enum

Private Enum VirtualAllocPageFlags
    PAGE_EXECUTE = &H10
    PAGE_EXECUTE_READ = &H20
    PAGE_EXECUTE_READWRITE = &H40
    PAGE_EXECUTE_WRITECOPY = &H80
    PAGE_NOACCESS = &H1
    PAGE_READONLY = &H2
    PAGE_READWRITE = &H4
    PAGE_WRITECOPY = &H8
    PAGE_GUARD = &H100
    PAGE_NOCACHE = &H200
    PAGE_WRITECOMBINE = &H400
End Enum

' calls the last method of an interface
Public Sub CallCustomConstructor( _
    obj As Object, _
    ParamArray params() _
)

    CallStd zAddressOf(obj, 1), ObjPtr(obj), params
End Sub

'Return the address of the specified ordinal method
'on the oCallback object,
'1 = last private method,
'2 = second last private method, etc
Private Function zAddressOf( _
    ByVal oCallback As Object, _
    ByVal nOrdinal As Long _
) As Long

    Dim bSub  As Byte                                   'Value we expect to find pointed at by a vTable method entry
    Dim bVal  As Byte
    Dim nAddr As Long                                   'Address of the vTable
    Dim i     As Long                                   'Loop index
    Dim j     As Long                                   'Loop limit

    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4   'Get the address of the callback object's instance
    If Not zProbe(nAddr + &H1C, i, bSub) Then           'Probe for a Class method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then      'Probe for a Form method
            If Not zProbe(nAddr + &H7A4, i, bSub) Then  'Probe for a UserControl method
                Exit Function                           'Bail...
            End If
        End If
    End If
  
    i = i + 4                                           'Bump to the next entry
    j = i + 1024                                        'Set a reasonable limit, scan 256 vTable entries

    Do While i < j
        RtlMoveMemory VarPtr(nAddr), i, 4               'Get the address stored in this vTable entry

        If IsBadCodePtr(nAddr) Then                     'Is the entry an invalid code address?
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4     'Return the specified vTable entry address
            Exit Do                                                     'Bad method signature, quit loop
        End If

        RtlMoveMemory VarPtr(bVal), nAddr, 1            'Get the byte pointed to by the vTable entry
        If bVal <> bSub Then                            'If the byte doesn't match the expected value...
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4     'Return the specified vTable entry address
            Exit Do                                                     'Bad method signature, quit loop
        End If

        i = i + 4                                                       'Next vTable entry
    Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe( _
    ByVal nStart As Long, _
    ByRef nMethod As Long, _
    ByRef bSub As Byte _
) As Boolean

    Dim bVal    As Byte
    Dim nAddr   As Long
    Dim nLimit  As Long
    Dim nEntry  As Long

    nAddr = nStart                                      'Start address
    nLimit = nAddr + 32                                 'Probe eight entries
    Do While nAddr < nLimit                             'While we've not reached our probe depth
        RtlMoveMemory VarPtr(nEntry), nAddr, 4          'Get the vTable entry

        If nEntry <> 0 Then                             'If not an implemented interface
            RtlMoveMemory VarPtr(bVal), nEntry, 1       'Get the value pointed at by the vTable entry
            If bVal = &H33 Or bVal = &HE9 Then          'Check for a native or pcode method signature
                nMethod = nAddr                         'Store the vTable entry
                bSub = bVal                             'Store the found method signature
                zProbe = True                           'Indicate success
                Exit Function                           'Return
            End If
        End If

        nAddr = nAddr + 4                               'Next vTable entry
    Loop
End Function

' call a function pointer (stdcall calling convention)
Private Function CallStd( _
    ByVal fnc As Long, _
    ParamArray params() As Variant _
) As Long

    Dim pMemory             As Long
    Dim cBytesMemory        As Long
    Dim pASM                As Long
    Dim i                   As Integer
    Dim j                   As Integer

    If fnc = 0 Then
        Err.Raise "CallStd called with fnc=0!"
    End If

    cBytesMemory = &HEC00&
    pMemory = VirtualAlloc(0, cBytesMemory, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pMemory = 0 Then Exit Function

    pASM = pMemory

    AddByte pASM, &H58                  ' POP EAX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H59                  ' POP ECX
    AddByte pASM, &H50                  ' PUSH EAX

    For i = UBound(params) To 0 Step -1
        If IsArray(params(i)) Then
            For j = UBound(params(i)) To 0 Step -1
                AddPush pASM, CLng(params(i)(j))    ' PUSH dword
            Next
        Else
            AddPush pASM, CLng(params(i))           ' PUSH dword
        End If
    Next

    AddCall pASM, fnc                   ' CALL rel addr
    AddByte pASM, &HC3                  ' RET

    CallStd = CallWindowProcA(pMemory, 0, 0, 0, 0)

    VirtualFree pMemory, cBytesMemory, MEM_DECOMMIT
End Function

Private Sub AddPush(pASM As Long, lng As Long)
    AddByte pASM, &H68
    AddLong pASM, lng
End Sub

Private Sub AddCall(pASM As Long, addr As Long)
    AddByte pASM, &HE8
    AddLong pASM, addr - pASM - 4
End Sub

Private Sub AddLong(pASM As Long, lng As Long)
    RtlMoveMemory pASM, VarPtr(lng), 4
    pASM = pASM + 4
End Sub

Private Sub AddByte(pASM As Long, Bt As Byte)
    RtlMoveMemory pASM, VarPtr(Bt), 1
    pASM = pASM + 1
End Sub
