VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmChannel 
   Caption         =   "IRC client"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin RichTextLib.RichTextBox rtfDisplay 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5106
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmChannel.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstMode 
      Height          =   255
      Left            =   1080
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lstNick 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   3240
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   -15
      Width           =   1455
   End
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   -15
      MaxLength       =   255
      TabIndex        =   1
      Top             =   2985
      Width           =   2415
   End
End
Attribute VB_Name = "frmChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    ChangeObjectColor rtfDisplay, udtColor.colorFrame, 1
    ChangeObjectColor lstNick, udtColor.colorList, 1
    ChangeObjectColor lstNick, udtColor.colorListText, 2
    ChangeObjectColor txtSend, udtColor.colorEdit, 1
    ChangeObjectColor txtSend, udtColor.colorEditText, 2
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtSend.Top = Me.Height - txtSend.Height - 620
    txtSend.Width = Me.ScaleWidth '- 70
    lstNick.Left = Me.Width - lstNick.Width - 70
    lstNick.Height = txtSend.Top + 50
    rtfDisplay.Width = Me.Width - lstNick.Width - 40
    rtfDisplay.Height = txtSend.Top - 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intIndex As Integer
    Dim strMsg As String
    Dim strName As String
    
    'notify that this current index is available for function FindFreeIndex
    strMsg = Me.Tag
    strName = Left(strMsg, Len(strMsg) - 2)
    SendData "PART " & strName
    intIndex = Right(strMsg, Len(strMsg) - InStr(strMsg, ","))
    ChannelArrayFree(intIndex) = True
End Sub
Private Sub lstNick_DblClick()
    Dim strNick As String
    strNick = lstNick.Text
    
    If Left(strNick, 1) = "@" Or Left(strNick, 1) = "+" Then
        strNick = Mid(strNick, 2)
    End If
    If GetCaption(LCase(strNick)) = LCase(strNick) Then GiveFocus LCase(strNick): Exit Sub
    CreateHwnd LCase(strNick), Private_hWnd
    ChangeCaption LCase(strNick), strNick
End Sub
Private Sub rtfDisplay_GotFocus()
    txtSend.SetFocus
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
On Error GoTo xError

    Dim strTemp As String
    Dim strTag As String
    Static LastSent As String
    If KeyAscii = 13 Then
    
        txtSend = LTrim(txtSend)
        If Left(txtSend.Text, 1) = "/" Then
            Call xINPUT(Mid(txtSend, 2))
        Else
        
        strTemp = Me.Tag
        strTag = Left(strTemp, InStr(strTemp, ",") - 1)
        SendData "PRIVMSG " & strTag & " :" & txtSend
        LogText rtfDisplay, "<" & MyNick & "> " & txtSend, udtColor.colorOwn

        End If

        txtSend = ""
        KeyAscii = 0
    End If
    If KeyAscii = 38 Then
        txtSend = LastSent
    End If
xError:
    If Err.Description <> "" Then
        MsgBox Err.Description
    End If
End Sub
