VERSION 5.00
Begin VB.Form frm3dstars 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " 3D Stars 1.0"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6315
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "3dstars10.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   120
   End
End
Attribute VB_Name = "frm3dstars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' 3D Stars v1.0 (in a form)
'
' © 2017 // thomas.schilb@live.de
'
Dim X(100), Y(100), Z(100) As Integer
Dim tmpX(100), tmpY(100), tmpZ(100) As Integer
Dim K As Integer
Dim Zoom As Integer
Dim Speed As Integer
Private Sub Form_Activate()
Speed = -1
K = 2038
Zoom = 256
Timer1.Interval = 1
For i = 0 To 100
X(i) = Int(Rnd * 1024) - 512
Y(i) = Int(Rnd * 1024) - 512
Z(i) = Int(Rnd * 512) - 256
Next i
End Sub
Private Sub Timer1_Timer()
For i = 0 To 100
Circle (tmpX(i), tmpY(i)), 5, BackColor
Z(i) = Z(i) + Speed
If Z(i) > 255 Then Z(i) = -255
If Z(i) < -255 Then Z(i) = 255
tmpZ(i) = Z(i) + Zoom
tmpX(i) = (X(i) * K / tmpZ(i)) + (frm3dstars.Width / 2)
tmpY(i) = (Y(i) * K / tmpZ(i)) + (frm3dstars.Height / 2)
Radius = 1
StarColor = 256 - Z(i)
Circle (tmpX(i), tmpY(i)), 5, RGB(StarColor, StarColor, StarColor)
Next i
End Sub
