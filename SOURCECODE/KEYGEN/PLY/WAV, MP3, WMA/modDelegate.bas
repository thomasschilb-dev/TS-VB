Attribute VB_Name = "modDelegate"
Option Explicit

Private Declare Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As Long, _
    ByVal dwSize As Long, _
    ByVal flAllocType As Long, _
    ByVal flProtect As Long _
) As Long

Private Declare Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As Long, _
    ByVal dwSize As Long, _
    ByVal dwFreeType As Long _
) As Long

Private Declare Function VirtualProtect Lib "kernel32" ( _
    ByVal lpAddress As Long, _
    ByVal dwSize As Long, _
    ByVal flNewProtect As Long, _
    lpflOldProtect As Long _
) As Long

Private Declare Sub CpyMem Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    Optional ByVal dwLen As Long = 4 _
)

Public Type Memory
    address     As Long
    bytes       As Long
End Type

Public Type MethodDelegate
    hMem        As Memory
    addr        As Long
End Type

Public Enum VirtualFreeTypes
    MEM_DECOMMIT = &H4000
    MEM_RELEASE = &H8000
End Enum

Public Enum VirtualAllocTypes
    MEM_COMMIT = &H1000
    MEM_RESERVE = &H2000
    MEM_RESET = &H8000
    MEM_LARGE_PAGES = &H20000000
    MEM_PHYSICAL = &H100000
    MEM_WRITE_WATCH = &H200000
End Enum

Public Enum VirtualAllocPageFlags
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

Public Function AllocMemory( _
    ByVal bytes As Long, _
    Optional ByVal lpAddr As Long = 0, _
    Optional ByVal PageFlags As VirtualAllocPageFlags = PAGE_READWRITE _
) As Memory

    With AllocMemory
        .address = VirtualAlloc(lpAddr, bytes, MEM_COMMIT, PageFlags)
        .bytes = bytes
    End With
End Function

Public Function FreeMemory( _
    udtMem As Memory _
) As Boolean

    VirtualFree udtMem.address, udtMem.bytes, MEM_DECOMMIT

    udtMem.address = 0
    udtMem.bytes = 0
End Function

' Creates delegates for class methods
'
' Machinecode by VF-fCRO:
' http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=38801&lngWId=1
'
Public Function CreateDelegate( _
    obj As Object, _
    VTblIndexOrPointer As Long, _
    cParams As Long _
) As MethodDelegate

    Dim pADR        As Long
    Dim pASM        As Long
    Dim pOut        As Long
    Dim delegate    As MethodDelegate
    Dim i           As Integer

    If obj Is Nothing Then
        pADR = VTblIndexOrPointer
    Else
        CpyMem pADR, ByVal ObjPtr(obj), 4
        CpyMem pADR, ByVal pADR + &H1C& + (4 * VTblIndexOrPointer), 4
    End If

    If pADR = 0 Then Exit Function

    With delegate
        .hMem = AllocMemory(31 + cParams * 3, , PAGE_EXECUTE_READWRITE)
        .addr = .hMem.address
    End With

    pASM = delegate.addr
    If pASM = 0 Then Exit Function

    pOut = delegate.addr + 31 + (cParams * 3) - 4

    AddLong pASM, &H68EC8B55
    AddLong pASM, pOut

    For i = 4 + 4 * cParams To 5 Step -4
        AddInteger pASM, &H75FF
        AddByte pASM, CByte(i)
    Next i

    AddPush pASM, ObjPtr(obj)
    AddCall pASM, pADR
    AddByte pASM, &HA1
    AddLong pASM, pOut
    AddInteger pASM, &HC2C9
    AddInteger pASM, cParams * 4

    CreateDelegate = delegate
End Function

Public Sub FreeDelegate( _
    fnc As MethodDelegate _
)

    FreeMemory fnc.hMem
End Sub

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

Private Sub AddInteger( _
    pASM As Long, _
    iInt As Integer _
)

    CpyMem ByVal pASM, iInt, 2
    pASM = pASM + 2
End Sub

Private Sub AddByte( _
    pASM As Long, _
    bt As Byte _
)

    CpyMem ByVal pASM, bt, 1
    pASM = pASM + 1
End Sub
