VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStatus 
   Caption         =   "Status"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   6855
   Begin RichTextLib.RichTextBox rtfStatus 
      Height          =   3855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   6800
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmStatus.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      MaxLength       =   255
      TabIndex        =   0
      Top             =   3960
      Width           =   6705
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call Form_Resize
    ChangeObjectColor rtfStatus, udtColor.colorFrame, 1
    ChangeObjectColor txtSend, udtColor.colorEdit, 1
    ChangeObjectColor txtSend, udtColor.colorEditText, 2
    ChangeObjectColor rtfStatus, udtColor.colorOther, 3
End Sub

Private Sub Form_Paint()
    Call Form_Resize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtSend.Top = rtfStatus.Height
    rtfStatus.Width = Me.ScaleWidth
    txtSend.Width = Me.ScaleWidth
    rtfStatus.Height = (Me.Height - txtSend.Height - 520)
End Sub

Private Sub rtfStatus_GotFocus()
    txtSend.SetFocus
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    On Error GoTo xError
    Static LastSent As String
    If KeyAscii = 13 Then
    
        txtSend = LTrim(txtSend)
        If Left(txtSend.Text, 1) = "/" Then
            Call xINPUT(Mid(txtSend, 2))
        End If

        'If LCase(lastsend) <> LCase(txtSend) Then
        '    LastSent = txtSend
        'End If
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
