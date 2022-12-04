VERSION 5.00
Begin VB.Form tsMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "tsMIDI-1.0"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2010
   Icon            =   "tsMIDI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   2010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   510
      Left            =   990
      TabIndex        =   1
      Top             =   0
      Width           =   1005
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1005
   End
End
Attribute VB_Name = "tsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPlay_Click()
PlayMID "music"
End Sub

Private Sub cmdStop_Click()
StopMID "music"
End Sub

Private Sub Form_Load()
Dim MidiFile As String
MidiFile = "tsMIDI.mid"
OpenMID MidiFile, "music"
PlayMID "music"
End Sub

Private Sub Form_Terminate()
CloseMID "music"
End Sub

Private Sub Form_Unload(Cancel As Integer)
CloseMID "music"
End Sub



