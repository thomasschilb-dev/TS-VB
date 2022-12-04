VERSION 5.00
Object = "{2BF273BF-D8EB-11D3-8ED1-CFAA7A653C7F}#1.0#0"; "MBFormEx.ocx"
Begin VB.Form frmMain 
   Caption         =   "MB Extended Form Control - Sample App"
   ClientHeight    =   5790
   ClientLeft      =   4185
   ClientTop       =   2100
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   5175
   Begin VB.ComboBox cmbTransparent 
      Height          =   315
      ItemData        =   "MBFormExDemo.frx":0000
      Left            =   723
      List            =   "MBFormExDemo.frx":000D
      Style           =   2  'Dropdown-Liste
      TabIndex        =   14
      Top             =   2160
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   5085
      TabIndex        =   12
      Top             =   5520
      Width           =   5145
      Begin VB.Label lblStatus 
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   6015
      End
   End
   Begin VB.ListBox lstFiles 
      Height          =   1815
      Left            =   60
      TabIndex        =   11
      Top             =   3480
      Width           =   5055
   End
   Begin VB.CheckBox chkDraggable 
      Caption         =   "Draggable"
      Height          =   255
      Left            =   723
      TabIndex        =   10
      Top             =   720
      Width           =   1695
   End
   Begin VB.CheckBox chkLimitSize 
      Caption         =   "LimitSize"
      Height          =   255
      Left            =   723
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CheckBox chkOnTop 
      Caption         =   "AlwaysOnTop"
      Height          =   255
      Left            =   723
      TabIndex        =   8
      Top             =   480
      Width           =   1695
   End
   Begin VB.CheckBox chkAcceptFiles 
      Caption         =   "AcceptFiles"
      Height          =   255
      Left            =   716
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.CheckBox chkTask 
      Caption         =   "ShowInTaskList"
      Height          =   255
      Left            =   723
      TabIndex        =   6
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CheckBox chkFlashing 
      Caption         =   "FlashingCaption"
      Height          =   255
      Left            =   723
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtInterval 
      Height          =   285
      Left            =   3483
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdInterval 
      Caption         =   "Set"
      Height          =   255
      Left            =   2643
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CheckBox chkResizeCtl 
      Caption         =   "ResizeControls"
      Height          =   255
      Left            =   723
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CheckBox chkResizeFonts 
      Caption         =   "ResizeFonts"
      Height          =   255
      Left            =   723
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdFull 
      Caption         =   "Toggle FullScreen"
      Height          =   375
      Left            =   723
      TabIndex        =   0
      Top             =   2520
      Width           =   3735
   End
   Begin MBFormEx.FormEx FormEx1 
      Left            =   2408
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      MaxHeight       =   500
      MaxWidth        =   500
      MinHeight       =   250
      MinWidth        =   250
      Draggable       =   -1  'True
      LimitSize       =   -1  'True
      AlwaysOnTop     =   -1  'True
      AcceptFiles     =   -1  'True
      Picture         =   "MBFormExDemo.frx":005B
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************
' MB Extended Form Control - Test App
' Version 1.2 - August 2000
' Copyright 2000 by Marco Bellinaso
' EMail: mbellinaso@vb2themax.com
' Web site: www.vb2themax.com
'***************************************

Option Explicit
Dim ID_HELLO As Long
Dim ID_SEP As Long
Dim ID_ABOUT As Long
'system menu items ID constants
Const SC_CLOSE = &HF060
Const SC_MAXIMIZE = &HF030
Const SC_MINIMIZE = &HF020
Const SC_MOVE = &HF010
Const SC_RESTORE = &HF120
Const SC_SIZE = &HF000


Private Sub chkAcceptFiles_Click()
    FormEx1.AcceptFiles = (chkAcceptFiles = vbChecked)
End Sub

Private Sub chkDraggable_Click()
    FormEx1.Draggable = (chkDraggable = vbChecked)
End Sub

Private Sub chkFlashing_Click()
    FormEx1.FlashingCaption = (chkFlashing = vbChecked)
End Sub

Private Sub chkLimitSize_Click()
    FormEx1.LimitSize = (chkLimitSize = vbChecked)
End Sub

Private Sub chkOnTop_Click()
    FormEx1.AlwaysOnTop = (chkOnTop = vbChecked)
End Sub

Private Sub chkResizeCtl_Click()
    FormEx1.ResizeControls = (chkResizeCtl = vbChecked)
End Sub

Private Sub chkResizeFonts_Click()
    FormEx1.ResizeFonts = (chkResizeFonts = vbChecked)
End Sub

Private Sub chkTask_Click()
    FormEx1.ShowInTaskList = (chkTask = vbChecked)
End Sub

Private Sub cmdFull_Click()
    FormEx1.FullScreen = Not (FormEx1.FullScreen)
End Sub

Private Sub cmdInterval_Click()
    FormEx1.FlashInterval = txtInterval
End Sub

Private Sub Form_Load()
    'add a separator in the system menu
    ID_SEP = FormEx1.AddSysItem("-")
    'add two items in the system menu. The return values are stored
    'because they will be used to identify the item menu in the
    'ItemSelect and SysItemClick events
    ID_HELLO = FormEx1.AddSysItem("Hello message")
    ID_ABOUT = FormEx1.AddSysItem("About FormEx ActiveX")
    'set the checkboxes' values according to the FormEx properties
    chkAcceptFiles.Value = IIf(FormEx1.AcceptFiles, vbChecked, vbUnchecked)
    chkOnTop.Value = IIf(FormEx1.AlwaysOnTop, vbChecked, vbUnchecked)
    chkDraggable.Value = IIf(FormEx1.Draggable, vbChecked, vbUnchecked)
    chkFlashing.Value = IIf(FormEx1.FlashingCaption, vbChecked, vbUnchecked)
    chkLimitSize.Value = IIf(FormEx1.LimitSize, vbChecked, vbUnchecked)
    chkResizeCtl.Value = IIf(FormEx1.ResizeControls, vbChecked, vbUnchecked)
    chkResizeFonts.Value = IIf(FormEx1.ResizeFonts, vbChecked, vbUnchecked)
    chkTask.Value = IIf(FormEx1.ShowInTaskList, vbChecked, vbUnchecked)
    txtInterval = FormEx1.FlashInterval
    cmbTransparent.ListIndex = FormEx1.Transparent
End Sub

Private Sub FormEx1_ActivateApp()
    Debug.Print "FormEx_ActivateApp"
End Sub

Private Sub FormEx1_CompactingMemory()
    Debug.Print "FormEx_CompactingMemory"
End Sub

Private Sub FormEx1_DeactivateApp()
    Debug.Print "FormEx_DeactivateApp"
End Sub

Private Sub FormEx1_DisplayChanged(ByVal Width As Long, ByVal Height As Long, ByVal NumColors As Long)
    MsgBox Width & " " & Height & " " & NumColors
End Sub

Private Sub FormEx1_DragDropFiles(FilesArray() As String)
    Dim i As Integer
    'add the file names to list
    For i = 0 To UBound(FilesArray)
        lstFiles.AddItem FilesArray(i)
    Next i
End Sub

Private Sub FormEx1_ItemSelect(ByVal ID As Long)
    'show the item's description in the lblStatus label
    Select Case ID
        Case Is = ID_ABOUT
            lblStatus = "Open the ActiveX's about box"
        Case Is = ID_HELLO
            lblStatus = "Show a message box"
        Case Is = SC_MOVE
            lblStatus = "Move the window"
        Case Is = SC_MAXIMIZE
            lblStatus = "Maximize the window"
        Case Is = SC_MINIMIZE
            lblStatus = "Minimize the window"
        Case Is = SC_RESTORE
            lblStatus = "Restore the window"
        Case Is = SC_SIZE
            lblStatus = "Resize the window"
        Case Is = SC_CLOSE
            lblStatus = "Close the window. Are you tired?"
        Case Else
            lblStatus = ""
    End Select
End Sub

Private Sub FormEx1_Move()
    Debug.Print "FormEx_Move"
End Sub

Private Sub FormEx1_Paint()
    Debug.Print "FormEx_Paint"
End Sub

Private Sub FormEx1_SysItemClick(ByVal ID As Long, sItem As String)
    Select Case ID
        Case Is = ID_ABOUT
            If FormEx1.AlwaysOnTop Then
                'disable the AlwaysOnTop, open the AboutBox and re-set the
                'AlwayOnTop. This is necessary, because if the form is always
                'on top the AboutBox can't be visible-->can't be closed
                '-->application freezed!!!
                FormEx1.AlwaysOnTop = False
                FormEx1.About
                FormEx1.AlwaysOnTop = True
            Else
                FormEx1.About
            End If
        Case Is = ID_HELLO
            MsgBox "Hello dear. Do you like FormEx ActiveX? If not mail me!", vbQuestion, "FormEx ActiveX"
    End Select
End Sub

Private Sub lstFiles_Click()
    lstFiles.Clear
End Sub

Private Sub cmbTransparent_Click()
    FormEx1.Transparent = cmbTransparent.ListIndex
End Sub
