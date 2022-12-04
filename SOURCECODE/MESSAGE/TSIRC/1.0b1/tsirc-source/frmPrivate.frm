VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPrivate 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmPrivate.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox txtSend 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      MaxLength       =   255
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox rtfDisplay 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2990
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmPrivate.frx":058A
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
End
Attribute VB_Name = "frmPrivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ChangeObjectColor rtfDisplay, udtColor.colorFrame, 1
    ChangeObjectColor txtSend, udtColor.colorEdit, 1
    ChangeObjectColor txtSend, udtColor.colorEditText, 2
End Sub

Private Sub Form_Paint()
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtSend.Top = rtfDisplay.Height
    rtfDisplay.Width = Me.Width - 70
    txtSend.Width = Me.Width - 70
    rtfDisplay.Height = (Me.Height - txtSend.Height - 400)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intIndex As Integer
    Dim strMsg As String
    Dim strName As String
    'notify that this current index is available for function FindFreeIndex
    strMsg = Me.Tag
    strName = Left(strMsg, Len(strMsg) - 2)
    intIndex = Right(strMsg, Len(strMsg) - InStr(strMsg, ","))
    PrivateArrayFree(intIndex) = True
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
On Error GoTo xError
    Dim strCommand As String
    Dim strMsg As String
    Static LastSent As String
    If KeyAscii = 13 Then
    
        txtSend = LTrim(txtSend)
        If Left(txtSend.Text, 1) = "/" Then
            Call xINPUT(Mid(txtSend, 2))
        Else
            LogTextToHwnd Trim(LCase(Left(Me.Tag, InStr(Me.Tag, ",") - 1))), "<" & MyNick & "> " & txtSend.Text, Private_hWnd, udtColor.colorChat, True
            SendData "PRIVMSG " & Me.Caption & " :" & txtSend
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
