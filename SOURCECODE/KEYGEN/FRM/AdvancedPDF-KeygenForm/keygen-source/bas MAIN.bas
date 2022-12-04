Attribute VB_Name = "basMAIN"
Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const Ht = 2
Private Const WM_NCLDOWN = &HA1

Public Sub DragWindow(Obj As Object)
    ReleaseCapture
    SendMessage Obj.hwnd, WM_NCLDOWN, Ht, 0&
End Sub

Public Sub Main()
If App.PrevInstance = True Then
    End
Else
    Load frmTEST
    frmTEST.Show
End If
End Sub
