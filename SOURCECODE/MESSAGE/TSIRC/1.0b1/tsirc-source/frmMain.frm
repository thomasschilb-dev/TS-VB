VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "tsIRC"
   ClientHeight    =   7545
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   10335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuReConnect 
         Caption         =   "ReConnect"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
         Shortcut        =   ^D
      End
      Begin VB.Menu hy00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "View"
      Begin VB.Menu mnuOption 
         Caption         =   "Options"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuColors 
         Caption         =   "Colors..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Contents"
         Shortcut        =   ^H
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About tsIRC..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long


Private Sub MDIForm_Load()
    LoadColor
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call mnuExit_Click
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuColors_Click()
frmColor.Show
End Sub

Private Sub mnuConnect_Click()
    frmOption.Show 1
End Sub

Private Sub mnuContents_Click()
'Open "sirc.chm" For Output As #1
'Open "sirc.chm" For Binary As #1
Dim r As Long
r = StartDoc("tsIRC.chm")
End Sub
Function StartDoc(DocName As String) As Long
Dim Scr_hDC As Long
Scr_hDC = GetDesktopWindow()
'change "Open" to "Explore" to bring up file explorer
StartDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "", 1)
'end function
End Function

Private Sub mnuDisconnect_Click()
Disconnect
End Sub

Private Sub mnuExit_Click()
    Disconnect
    End
End Sub

Private Sub mnuOption_Click()
    frmOption.Show 1
End Sub

