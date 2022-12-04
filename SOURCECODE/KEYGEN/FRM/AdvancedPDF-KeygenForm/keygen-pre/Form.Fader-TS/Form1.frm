VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fader"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Fade Black"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fade Color "
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   3840
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Matrix
'Qais Ghalib
'Qais60@hotmail.com
'2001

Private Sub Command1_Click()
Dim w, b, i, Y
Me.BackColor = &H0&
Me.DrawStyle = 6
Me.DrawMode = 13
Me.DrawWidth = 2
Me.ScaleMode = 3
Me.ScaleHeight = (256 * 2)
For w = 255 To 0 Step -1
Me.Line (0, 0)-(Me.Width, b + 1), RGB(255, 255 - w, 0), BF
b = b + 100
Next w
End Sub

Private Sub Command2_Click()
Dim w, b, i, Y
Me.BackColor = &H0&
Me.DrawStyle = 6
Me.DrawMode = 13
Me.DrawWidth = 2
Me.ScaleMode = 3
Me.ScaleHeight = (256 * 2)
For w = 255 To 0 Step -1
Me.Line (0, b)-(Me.Width, b + 1), RGB(w + 1, w, w * 1), BF
b = b + 100
Next w
For i = 255 To 0 Step -1
Me.Line (0, 0)-(Me.Width, Y + 1), RGB(i + 1, i, i * 1), BF
Y = Y + 100
Next i
End Sub

Private Sub Form_Load()

End Sub
