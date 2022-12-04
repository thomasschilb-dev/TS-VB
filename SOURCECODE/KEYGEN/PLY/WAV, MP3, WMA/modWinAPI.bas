Attribute VB_Name = "modWinAPI"
Option Explicit

Public Declare Sub CpyMem Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal dwLen As Long _
)

Public Declare Sub FillMem Lib "kernel32.dll" _
Alias "RtlFillMemory" ( _
    pDst As Any, _
    ByVal length As Long, _
    ByVal Fill As Byte _
)

Public Declare Sub ZeroMem Lib "kernel32" _
Alias "RtlZeroMemory" ( _
    pDst As Any, _
    ByVal dwLen As Long _
)

Public Declare Function IsBadReadPtr Lib "kernel32" ( _
    ptr As Any, _
    ByVal ucb As Long _
) As Long

Public Declare Function IsBadWritePtr Lib "kernel32" ( _
    ptr As Any, _
    ByVal ucb As Long _
) As Long
