VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colors"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstColor 
      Height          =   1035
      Left            =   4800
      TabIndex        =   31
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1110
      TabIndex        =   29
      Top             =   4050
      Width           =   855
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "OK"
      Height          =   375
      Left            =   150
      TabIndex        =   28
      Top             =   4050
      Width           =   855
   End
   Begin VB.ComboBox cmbItem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   3600
      Width           =   1695
   End
   Begin VB.PictureBox picDisplay 
      Height          =   735
      Left            =   3645
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   26
      ToolTipText     =   "click to change color"
      Top             =   3600
      Width           =   735
   End
   Begin VB.PictureBox picFrame 
      Height          =   2330
      Left            =   210
      ScaleHeight     =   2265
      ScaleWidth      =   4080
      TabIndex        =   6
      Top             =   1110
      Width           =   4140
      Begin VB.PictureBox picList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2075
         Left            =   2880
         ScaleHeight     =   2040
         ScaleWidth      =   1185
         TabIndex        =   24
         Top             =   -10
         Width           =   1215
         Begin VB.Label lblColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Listbox text"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   45
            TabIndex        =   25
            Top             =   60
            Width           =   1095
         End
      End
      Begin VB.PictureBox picText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   4065
         TabIndex        =   23
         Top             =   2040
         Width           =   4095
         Begin VB.Label lblColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Editbox text"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   75
            TabIndex        =   30
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Chat text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   1665
         TabIndex        =   22
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Other text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   1665
         TabIndex        =   21
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Invite text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   1665
         TabIndex        =   20
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Topic text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   1665
         TabIndex        =   19
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Whois text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   1665
         TabIndex        =   18
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "User text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   1665
         TabIndex        =   17
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Nick text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   1665
         TabIndex        =   16
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Own text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1665
         TabIndex        =   15
         Top             =   60
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Quit text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   105
         TabIndex        =   14
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Notice text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   105
         TabIndex        =   13
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Part text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   105
         TabIndex        =   12
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Mode text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   105
         TabIndex        =   11
         Top             =   1500
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Kick text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   105
         TabIndex        =   10
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Join text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   105
         TabIndex        =   9
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "CTCP text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Action text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   60
         Width           =   1095
      End
   End
   Begin VB.HScrollBar hsBlue 
      Height          =   255
      Left            =   2760
      Max             =   255
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4080
      Width           =   855
   End
   Begin VB.HScrollBar hsGreen 
      Height          =   255
      Left            =   2760
      Max             =   255
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3840
      Width           =   855
   End
   Begin VB.HScrollBar hsRed 
      Height          =   255
      Left            =   2760
      Max             =   255
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "tsIRC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   33
      Top             =   480
      Width           =   888
   End
   Begin VB.Label lblStatic 
      Caption         =   "Blue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   5
      Top             =   4110
      Width           =   735
   End
   Begin VB.Label lblStatic 
      Caption         =   "Green"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   4
      Top             =   3855
      Width           =   615
   End
   Begin VB.Label lblStatic 
      Caption         =   "Red"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   3
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnClick As Boolean
Private Sub InitializeColor()
    Dim intCounter As Integer
    Dim intRed As Integer, intGreen As Integer, intBlue As Integer
    Dim strTemp As String
    
    For intCounter = 0 To 17
        cmbItem.AddItem lblColor(intCounter).Caption
    Next
    
    cmbItem.AddItem "Editbox"
    cmbItem.AddItem "Background"
    cmbItem.AddItem "Listbox"
    
    With udtColor
    
        strTemp = .colorAction 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(0).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 0
        
        strTemp = .colorCTCP 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(1).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 1
        
        strTemp = .colorJoin 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(2).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 2
        
        strTemp = .colorPart 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(3).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 3
        
        strTemp = .colorQuit 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(4).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 4
        
        strTemp = .colorKick 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(5).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 5
        
        strTemp = .colorMode 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(6).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 6
        
        strTemp = .colorNotice 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(7).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 7
        
        strTemp = .colorOwn 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(8).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 8
        
        strTemp = .colorNick 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(9).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 9
        
        strTemp = .colorUser 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(10).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 10
        
        strTemp = .colorInvite 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(11).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 11
        
        strTemp = .colorTopic 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(12).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 12
        
        strTemp = .colorWhois 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(13).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 13
        
        strTemp = .colorChat 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(14).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 14
        
        strTemp = .colorOther 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(15).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 15
        
        strTemp = .colorListText 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(16).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 16
        
        strTemp = .colorEditText 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        lblColor(17).ForeColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 17
        
        strTemp = .colorEdit 'frmMain.lstColor.List(18)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        picText.BackColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 18
        
        strTemp = .colorFrame 'frmMain.lstColor.List(19)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        picFrame.BackColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 19
        
        strTemp = .colorList 'frmMain.lstColor.List(20)
        intRed = Val("&H" & Right(strTemp, 2))
        intGreen = Val("&H" & Mid(strTemp, 3, 2))
        intBlue = Val("&H" & Left(strTemp, 2))
        picList.BackColor = RGB(intRed, intGreen, intBlue)
        lstColor.AddItem strTemp, 20
    End With
End Sub
Private Sub cmbItem_Click()
    Dim strColor As String
    Dim intRed As Integer, intGreen As Integer, intBlue As Integer
        strColor = lstColor.List(cmbItem.ListIndex)
        intRed = Val("&H" & Right(strColor, 2))
        intGreen = Val("&H" & Mid(strColor, 3, 2))
        intBlue = Val("&H" & Left(strColor, 2))
   
        hsRed.Value = intRed
        hsGreen.Value = intGreen
        hsBlue.Value = intBlue
End Sub

Private Sub cmbItem_DropDown()
    blnClick = False
End Sub

Private Sub cmbItem_KeyDown(KeyCode As Integer, Shift As Integer)
    blnClick = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOkay_Click()
    Dim lngCounter As Long
    Dim strTemp As String
    
    With udtColor
        .colorAction = lstColor.List(0)
        .colorCTCP = lstColor.List(1)
        .colorJoin = lstColor.List(2)
        .colorPart = lstColor.List(3)
        .colorQuit = lstColor.List(4)
        .colorKick = lstColor.List(5)
        .colorMode = lstColor.List(6)
        .colorNotice = lstColor.List(7)
        .colorOwn = lstColor.List(8)
        .colorNick = lstColor.List(9)
        .colorUser = lstColor.List(10)
        .colorInvite = lstColor.List(11)
        .colorTopic = lstColor.List(12)
        .colorWhois = lstColor.List(13)
        .colorChat = lstColor.List(14)
        .colorOther = lstColor.List(15)
        .colorListText = lstColor.List(16)
        .colorEditText = lstColor.List(17)
        .colorEdit = lstColor.List(18)
        .colorFrame = lstColor.List(19)
        .colorList = lstColor.List(20)
    End With
    
    For lngCounter = 0 To Forms.Count - 1
        strTemp = Forms(lngCounter).Tag
        If Len(Trim(strTemp)) <> 0 Then
            With udtColor
                ChangeObjectColor Forms(lngCounter).lstNick, .colorList, 1
                ChangeObjectColor Forms(lngCounter).lstNick, .colorListText, 2
                ChangeObjectColor Forms(lngCounter).txtSend, .colorEdit, 1
                ChangeObjectColor Forms(lngCounter).txtSend, .colorEditText, 2
                ChangeObjectColor Forms(lngCounter).rtfDisplay, .colorFrame, 1
            End With
        End If
    Next

    ChangeObjectColor frmStatus.rtfStatus, udtColor.colorFrame, 1
    ChangeObjectColor frmStatus.txtSend, udtColor.colorEdit, 1
    ChangeObjectColor frmStatus.txtSend, udtColor.colorEditText, 2
    
    SaveColor
    Unload Me
End Sub

Private Sub Form_Load()
    InitializeColor
    picDisplay.BackColor = vbBlack
    cmbItem.ListIndex = 0
End Sub

Private Sub hsBlue_Change()
    picDisplay.BackColor = RGB(hsRed.Value, hsGreen.Value, hsBlue.Value)
End Sub

Private Sub hsGreen_Change()
    picDisplay.BackColor = RGB(hsRed.Value, hsGreen.Value, hsBlue.Value)
End Sub

Private Sub hsRed_Change()
    picDisplay.BackColor = RGB(hsRed.Value, hsGreen.Value, hsBlue.Value)
End Sub

Private Sub lblColor_Click(Index As Integer)
        
    Dim strColor As String
    Dim intRed As Integer, intGreen As Integer, intBlue As Integer
    
        strColor = lstColor.List(Index)
    
        intRed = Val("&H" & Right(strColor, 2))
        intGreen = Val("&H" & Mid(strColor, 3, 2))
        intBlue = Val("&H" & Left(strColor, 2))
    
        hsRed.Value = intRed
        hsGreen.Value = intGreen
        hsBlue.Value = intBlue
        cmbItem.ListIndex = Index

End Sub

Private Sub picDisplay_Click()
    If cmbItem.ListIndex > 17 Then
        Select Case cmbItem.ListIndex
            Case 18
                picText.BackColor = picDisplay.BackColor
                lstColor.RemoveItem 18
                lstColor.AddItem RGBtoHEX(picDisplay.BackColor), 18
            Case 19
                picFrame.BackColor = picDisplay.BackColor
                lstColor.RemoveItem 19
                lstColor.AddItem RGBtoHEX(picDisplay.BackColor), 19
            Case 20
                picList.BackColor = picDisplay.BackColor
                lstColor.RemoveItem 20
                lstColor.AddItem RGBtoHEX(picDisplay.BackColor), 20
        End Select
    Else
        lblColor(cmbItem.ListIndex).ForeColor = picDisplay.BackColor
        lstColor.RemoveItem cmbItem.ListIndex
        lstColor.AddItem RGBtoHEX(picDisplay.BackColor), cmbItem.ListIndex
    End If
End Sub

Private Sub picText_Click()
    Dim strColor As String
    Dim intRed As Integer, intGreen As Integer, intBlue As Integer
    
    strColor = lstColor.List(18)
    intRed = Val("&H" & Right(strColor, 2))
    intGreen = Val("&H" & Mid(strColor, 3, 2))
    intBlue = Val("&H" & Left(strColor, 2))
    
    hsRed.Value = intRed
    hsGreen.Value = intGreen
    hsBlue.Value = intBlue
    cmbItem.ListIndex = 18
    
End Sub

Private Sub picFrame_Click()
    Dim strColor As String
    Dim intRed As Integer, intGreen As Integer, intBlue As Integer
    
    strColor = lstColor.List(19)
    intRed = Val("&H" & Right(strColor, 2))
    intGreen = Val("&H" & Mid(strColor, 3, 2))
    intBlue = Val("&H" & Left(strColor, 2))
    
    hsRed.Value = intRed
    hsGreen.Value = intGreen
    hsBlue.Value = intBlue
    cmbItem.ListIndex = 19
    
End Sub

Private Sub picList_Click()
    Dim strColor As String
    Dim intRed As Integer, intGreen As Integer, intBlue As Integer
    
    strColor = lstColor.List(20)
    intRed = Val("&H" & Right(strColor, 2))
    intGreen = Val("&H" & Mid(strColor, 3, 2))
    intBlue = Val("&H" & Left(strColor, 2))
    
    hsRed.Value = intRed
    hsGreen.Value = intGreen
    hsBlue.Value = intBlue
    cmbItem.ListIndex = 20

End Sub

