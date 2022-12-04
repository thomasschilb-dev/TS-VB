Attribute VB_Name = "modFileAccess"
Option Explicit

Private Declare Function CreateFile Lib "kernel32.dll" _
Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long _
) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any _
) As Long

Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Any _
) As Long

Private Declare Function SetFilePointer Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lDistanceToMove As Long, _
    ByVal lpDistanceToMoveHigh As Long, _
    ByVal dwMoveMethod As Long _
) As Long

Private Declare Function GetFileSize Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpFileSizeHigh As Long _
) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long _
) As Long


Public Enum FILE_OPEN_METHOD
    CREATE_NEW = 1
    CREATE_ALWAYS = 2
    OPEN_EXISTING = 3
    OPEN_ALWAYS = 4
End Enum

Public Enum FILE_SHARE_RIGHTS
    FILE_SHARE_READ = &H1
    FILE_SHARE_WRITE = &H2
End Enum

Public Enum FILE_ACCESS_RIGHTS
    GENERIC_READ = &H80000000
    GENERIC_WRITE = &H40000000
End Enum

Public Enum SEEK_METHOD
    FILE_BEGIN = 0
    FILE_CURRENT = 1
    FILE_END = 2
End Enum

Public Const INVALID_HANDLE As Long = -1

Public Type hFile
    handle      As Long
    path        As String
End Type

Public Function IsValidFile( _
    ByVal strFile As String _
) As Boolean

    Dim hInp    As hFile

    hInp = FileOpen(strFile, GENERIC_READ, FILE_SHARE_READ)
    IsValidFile = hInp.handle <> INVALID_HANDLE
    FileClose hInp
End Function

Public Function FileOpen( _
    ByVal strFile As String, _
    Optional access As FILE_ACCESS_RIGHTS = GENERIC_READ Or GENERIC_WRITE, _
    Optional share As FILE_SHARE_RIGHTS = FILE_SHARE_READ Or FILE_SHARE_WRITE, _
    Optional method As FILE_OPEN_METHOD = OPEN_EXISTING _
) As hFile

    FileOpen.handle = CreateFile(strFile, _
                                 access, _
                                 share, _
                                 ByVal 0&, _
                                 method, _
                                 0, 0)

    FileOpen.path = strFile
End Function

Public Sub FileClose( _
    filehandle As hFile _
)

    CloseHandle filehandle.handle
    filehandle.handle = INVALID_HANDLE
    filehandle.path = vbNullString
End Sub

Public Function FileRead( _
    filehandle As hFile, _
    ByVal ptr As Long, _
    ByVal bytes As Long _
) As Long

    Dim dwRead  As Long
    Dim lngRet  As Long

    If filehandle.handle = INVALID_HANDLE Then Exit Function

    lngRet = ReadFile(filehandle.handle, ByVal ptr, bytes, dwRead, 0&)

    If lngRet = 1 Then
        FileRead = dwRead
    Else
        FileRead = -1
    End If
End Function

Public Function FileWrite( _
    filehandle As hFile, _
    ByVal ptr As Long, _
    ByVal bytes As Long _
) As Long

    Dim dwWritten   As Long
    Dim lngRet      As Long

    If filehandle.handle = INVALID_HANDLE Then Exit Function

    lngRet = WriteFile(filehandle.handle, ByVal ptr, bytes, dwWritten, 0&)

    If lngRet = 1 Then
        FileWrite = dwWritten
    Else
        FileWrite = -1
    End If
End Function

Public Function FileSeek( _
    filehandle As hFile, _
    ByVal bytes As Long, _
    ByVal method As SEEK_METHOD _
) As Long

    FileSeek = SetFilePointer(filehandle.handle, bytes, 0, method)

End Function

Public Function FilePosition( _
    filehandle As hFile _
) As Long

    FilePosition = FileSeek(filehandle, 0, FILE_CURRENT)

End Function

Public Function FileLength( _
    filehandle As hFile _
) As Long

    FileLength = GetFileSize(filehandle.handle, 0)
End Function

Public Function FileEnd( _
    filehandle As hFile _
) As Boolean

    FileEnd = FilePosition(filehandle) >= FileLength(filehandle)
End Function
