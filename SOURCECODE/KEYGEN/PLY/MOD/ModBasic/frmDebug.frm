VERSION 5.00
Begin VB.Form frmDebug 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Debug:"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   360
      Left            =   4455
      TabIndex        =   29
      Top             =   2730
      Width           =   1380
   End
   Begin VB.TextBox sm 
      Height          =   315
      Index           =   3
      Left            =   4755
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   2280
      Width           =   915
   End
   Begin VB.TextBox sm 
      Height          =   315
      Index           =   2
      Left            =   3735
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   2280
      Width           =   915
   End
   Begin VB.TextBox sm 
      Height          =   315
      Index           =   1
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2280
      Width           =   915
   End
   Begin VB.TextBox sm 
      Height          =   315
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   2280
      Width           =   915
   End
   Begin VB.TextBox fxp 
      Height          =   315
      Index           =   3
      Left            =   4755
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1845
      Width           =   915
   End
   Begin VB.TextBox fxp 
      Height          =   315
      Index           =   2
      Left            =   3735
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1845
      Width           =   915
   End
   Begin VB.TextBox fxp 
      Height          =   315
      Index           =   1
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1845
      Width           =   915
   End
   Begin VB.TextBox fxp 
      Height          =   315
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1845
      Width           =   915
   End
   Begin VB.TextBox fxv 
      Height          =   315
      Index           =   3
      Left            =   4755
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1410
      Width           =   915
   End
   Begin VB.TextBox fxv 
      Height          =   315
      Index           =   2
      Left            =   3735
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1410
      Width           =   915
   End
   Begin VB.TextBox fxv 
      Height          =   315
      Index           =   1
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1410
      Width           =   915
   End
   Begin VB.TextBox fxv 
      Height          =   315
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1410
      Width           =   915
   End
   Begin VB.TextBox notef 
      Height          =   315
      Index           =   3
      Left            =   4755
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   945
      Width           =   915
   End
   Begin VB.TextBox notef 
      Height          =   315
      Index           =   2
      Left            =   3735
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   945
      Width           =   915
   End
   Begin VB.TextBox notef 
      Height          =   315
      Index           =   1
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   945
      Width           =   915
   End
   Begin VB.TextBox notef 
      Height          =   315
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   945
      Width           =   915
   End
   Begin VB.TextBox note 
      Height          =   315
      Index           =   3
      Left            =   4755
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   450
      Width           =   915
   End
   Begin VB.TextBox note 
      Height          =   315
      Index           =   2
      Left            =   3735
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   450
      Width           =   915
   End
   Begin VB.TextBox note 
      Height          =   315
      Index           =   1
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   450
      Width           =   915
   End
   Begin VB.TextBox note 
      Height          =   315
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   450
      Width           =   915
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "SampleN:"
      Height          =   195
      Left            =   105
      TabIndex        =   8
      Top             =   2205
      Width           =   690
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Effekt parm:"
      Height          =   195
      Left            =   105
      TabIndex        =   7
      Top             =   1794
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Effekt val:"
      Height          =   195
      Left            =   105
      TabIndex        =   6
      Top             =   1386
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Note(FREQ):"
      Height          =   195
      Left            =   105
      TabIndex        =   5
      Top             =   978
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Note:"
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   570
      Width           =   390
   End
   Begin VB.Line Line1 
      X1              =   1170
      X2              =   1170
      Y1              =   270
      Y2              =   2985
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Channel 4"
      Height          =   195
      Left            =   4875
      TabIndex        =   3
      Top             =   90
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Channel 3"
      Height          =   195
      Left            =   3810
      TabIndex        =   2
      Top             =   90
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Channel 2"
      Height          =   195
      Left            =   2730
      TabIndex        =   1
      Top             =   90
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Channel 1"
      Height          =   195
      Left            =   1665
      TabIndex        =   0
      Top             =   90
      Width           =   720
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Me.Hide
End Sub

