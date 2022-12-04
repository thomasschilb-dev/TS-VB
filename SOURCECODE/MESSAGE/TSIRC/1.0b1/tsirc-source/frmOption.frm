VERSION 5.00
Begin VB.Form frmOption 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Connect"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frmOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   4200
      TabIndex        =   40
      Text            =   "6667"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtRealName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   960
      TabIndex        =   32
      Text            =   "Thomas Schilb"
      Top             =   2160
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   2640
      Width           =   5055
   End
   Begin VB.TextBox txtServer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   960
      TabIndex        =   30
      Text            =   "irc.thomasschilb.net"
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtNick 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   960
      TabIndex        =   29
      Text            =   "Guest"
      Top             =   1440
      Width           =   4215
   End
   Begin VB.TextBox txtMail 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   960
      TabIndex        =   28
      Text            =   "thomas_schilb@outlook.com"
      Top             =   1800
      Width           =   4215
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   1
      Left            =   210
      ScaleHeight     =   3135
      ScaleWidth      =   3495
      TabIndex        =   0
      Top             =   5760
      Width           =   3495
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect to IRC Server"
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   570
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Invisible mode"
         Height          =   195
         Left            =   1320
         TabIndex        =   11
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtFullName 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ComboBox cmbServers 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label lblStatic 
         Caption         =   "Alternative"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblStatic 
         Caption         =   "Nickname:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblStatic 
         Caption         =   "Email Address:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblStatic 
         Caption         =   "Full Name:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblStatic 
         Caption         =   "IRC Servers"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   150
         Width           =   975
      End
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   0
      Left            =   3750
      ScaleHeight     =   3135
      ScaleWidth      =   3495
      TabIndex        =   13
      Top             =   5760
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   2040
         TabIndex        =   27
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Okay"
         Height          =   315
         Left            =   480
         TabIndex        =   26
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Text            =   "6667"
         Top             =   2280
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Caption         =   "When connecting:"
         Height          =   1455
         Left            =   480
         TabIndex        =   16
         Top             =   720
         Width           =   2415
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   720
            TabIndex        =   22
            Text            =   "99"
            Top             =   600
            Width           =   375
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Try next server in group"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   720
            TabIndex        =   18
            Text            =   "99"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblStatic 
            Caption         =   "second(s)"
            Height          =   255
            Index           =   8
            Left            =   1200
            TabIndex        =   24
            Top             =   660
            Width           =   735
         End
         Begin VB.Label lblStatic 
            Caption         =   "Delay:"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   23
            Top             =   660
            Width           =   495
         End
         Begin VB.Label lblStatic 
            Caption         =   "time(s)"
            Height          =   255
            Index           =   6
            Left            =   1200
            TabIndex        =   21
            Top             =   300
            Width           =   495
         End
         Begin VB.Label lblStatic 
            Caption         =   "Retry:"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   20
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Reconnect on disconnection"
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Connect on start up"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblStatic 
         Caption         =   "Default port:"
         Height          =   255
         Index           =   9
         Left            =   960
         TabIndex        =   25
         Top             =   2325
         Width           =   975
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Port:"
      Height          =   255
      Left            =   3720
      TabIndex        =   39
      Top             =   1080
      Width           =   375
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
      TabIndex        =   37
      Top             =   480
      Width           =   888
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fullname:"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email:"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nick:"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Server:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   1080
      Width           =   735
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
      Left            =   -120
      TabIndex        =   38
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MyNick = txtNick
    Email = txtMail
    RealName = txtRealName
    Port = txtPort
    
    If Len(txtServer) = 0 Then Exit Sub
    Connect txtServer, txtPort
    frmStatus.Caption = "Status (" & txtServer & ":" & txtPort & ")"
    frmStatus.rtfStatus.SelColor = vbWhite
    frmStatus.rtfStatus.SelText = " *** Connecting to Server " & vbCrLf
Unload Me
End Sub
