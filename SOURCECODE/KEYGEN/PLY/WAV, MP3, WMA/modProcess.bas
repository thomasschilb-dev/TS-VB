Attribute VB_Name = "modProcess"
Option Explicit

' (c) Copyright 2003 Andrew Novick.
' You may use this code in your projects, including projects
' that you sell so long as there is substantial additional
' content. All other rights including rights to publication
' are reserved.
 

' Win32 API declarations

Private Declare Function GetCurrentProcess _
                                                    Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function SetThreadPriority Lib "kernel32" _
                                                       (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long

Private Const THREAD_BASE_PRIORITY_LOWRT As Long = 15 ' value that gets a thread to LowRealtime-1
Private Const THREAD_BASE_PRIORITY_MAX As Long = 2 ' maximum thread base priority boost
Private Const THREAD_BASE_PRIORITY_MIN As Long = -2 ' minimum thread base priority boost
Private Const THREAD_BASE_PRIORITY_IDLE As Long = -15 ' value that gets a thread to idle

Public Enum ThreadPriority
    THREAD_PRIORITY_LOWEST = -2
    THREAD_PRIORITY_BELOW_NORMAL = -1
    THREAD_PRIORITY_NORMAL = 0
    THREAD_PRIORITY_HIGHEST = 2
    THREAD_PRIORITY_ABOVE_NORMAL = 1
    THREAD_PRIORITY_TIME_CRITICAL = 15 ' THREAD_BASE_PRIORITY_LOWRT
    THREAD_PRIORITY_IDLE = -15 'THREAD_BASE_PRIORITY_IDLE
End Enum

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByVal lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                                                  ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' Used by the OpenProcess API call
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_SET_INFORMATION As Long = &H200

' Used by SetPriorityClass
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const BELOW_NORMAL_PRIORITY_CLASS = 16384
Private Const ABOVE_NORMAL_PRIORITY_CLASS = 32768
Private Const IDLE_PRIORITY_CLASS = &H40
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const REALTIME_PRIORITY_CLASS = &H100

Public Enum ProcessPriorities
    ppidle = IDLE_PRIORITY_CLASS
    ppbelownormal = BELOW_NORMAL_PRIORITY_CLASS
    ppAboveNormal = ABOVE_NORMAL_PRIORITY_CLASS
    ppNormal = NORMAL_PRIORITY_CLASS
    ppHigh = HIGH_PRIORITY_CLASS
    ppRealtime = REALTIME_PRIORITY_CLASS
End Enum

Public Function ProcessPriorityName(ByVal Priority As ProcessPriorities) As String

Dim sName As String

Select Case Priority

    Case ppidle
        sName = "Idle"

    Case ppbelownormal
        sName = "Below Normal"

    Case ppNormal
        sName = "Normal"

    Case ppAboveNormal
        sName = "Above Normal"

    Case ppHigh
        sName = "High"

    Case ppRealtime
        sName = "Realtime"

    Case Else
        sName = "Unknown:" & CStr(Priority)

End Select

ProcessPriorityName = sName

End Function

Public Function ProcessPriorityGet(Optional ByVal ProcessID As Long, Optional ByVal hWnd As Long) As Long

    ' Gets the process priority identified by an Id, a hWnd

    '  or if not identified, then the current process

    Dim hProc As Long
    Const fdwAccess As Long = PROCESS_QUERY_INFORMATION

    ' If not passed a PID, then find value from hWnd.

    If ProcessID = 0 Then
        If hWnd <> 0 Then
            Call GetWindowThreadProcessId(hWnd, ProcessID)
        Else
            ProcessID = GetCurrentProcessId()
        End If
    End If

    '   Need to open process with simple query rights,

    ' get the current setting, and close handle.
    hProc = OpenProcess(fdwAccess, 0&, ProcessID)
    ProcessPriorityGet = GetPriorityClass(hProc)

    Call CloseHandle(hProc)

End Function

Public Function ProcessPrioritySet( _
                    Optional ByVal ProcessID As Long, _
                    Optional ByVal hWnd As Long, _
                    Optional ByVal Priority As ProcessPriorities = NORMAL_PRIORITY_CLASS _
                    ) As Long

    Dim hProc As Long
    Const fdwAccess1 As Long = PROCESS_QUERY_INFORMATION Or PROCESS_SET_INFORMATION
    Const fdwAccess2 As Long = PROCESS_QUERY_INFORMATION

    ' If not passed a PID, then find value from hWnd.

    If ProcessID = 0 Then
        If hWnd <> 0 Then
            Call GetWindowThreadProcessId(hWnd, ProcessID)
        Else
            ProcessID = GetCurrentProcessId()
        End If
    End If

    ' Need to open process with setinfo rights.
    hProc = OpenProcess(fdwAccess1, 0&, ProcessID)

    If hProc Then
        ' Attempt to set new priority.
        Call SetPriorityClass(hProc, Priority)
    Else
        ' Weren't allowed to setinfo, so just open to
        ' enable return of current priority setting.
        hProc = OpenProcess(fdwAccess2, 0&, ProcessID)
    End If

    ' Get current/new setting.
    ProcessPrioritySet = GetPriorityClass(hProc)
    ' Clean up.
    Call CloseHandle(hProc)

End Function

Public Function ProcessThreadPrioritySet( _
                                Optional ByVal Priority As ThreadPriority = THREAD_PRIORITY_NORMAL _
                                ) As ThreadPriority

    Dim hThread As Long
    Dim rc As Long

    ' Set's the priority of the current thread

    hThread = GetCurrentThread()

    ' Need to open process with setinfo rights.

    rc = SetThreadPriority(hThread, Priority)
    ProcessThreadPrioritySet = GetThreadPriority(hThread)

End Function
