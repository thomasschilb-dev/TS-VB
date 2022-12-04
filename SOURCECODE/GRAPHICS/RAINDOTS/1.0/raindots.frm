VERSION 5.00
Begin VB.Form frmStarField 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Star Field"
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000E&
   Icon            =   "raindots.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerStarField 
      Interval        =   1
      Left            =   5040
      Top             =   3600
   End
End
Attribute VB_Name = "frmStarField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Chapter 1
'Starfield
'


Option Explicit

Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Star Type
Private Type Star
    X As Long
    Y As Long
    Speed As Long
    Size As Long
    Color As Long
End Type

'Star field array
Dim Stars(49) As Star
Const MaxSize As Long = 5
Const MaxSpeed As Long = 25

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Unload Me

End Sub

Private Sub Form_Load()
Dim I As Long

Randomize
'Generate the 100 stars
For I = LBound(Stars) To UBound(Stars)
    
    Stars(I).X = Me.ScaleWidth * Rnd + 1
    Stars(I).Y = Me.ScaleHeight * Rnd + 1
    Stars(I).Size = MaxSize * Rnd + 1
    Stars(I).Speed = MaxSpeed * Rnd + 1
    Stars(I).Color = RGB(Rnd * 255 + 1, Rnd * 255 + 1, Rnd * 255 + 1)
Next I

End Sub

Private Sub TimerStarField_Timer()
Dim I As Long

'clear the form
BitBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, 0, vbBlackness

For I = 0 To UBound(Stars)
    
    'Move the star
    Stars(I).Y = (Stars(I).Y Mod Me.ScaleHeight) + Stars(I).Speed
    'Relocate the X position
    If Stars(I).Y > Me.ScaleHeight Then
      Stars(I).X = Me.ScaleWidth * Rnd + 1
      Stars(I).Speed = MaxSpeed * Rnd + 1
    End If
    'Set the color
    Me.FillColor = Stars(I).Color
    Me.ForeColor = Stars(I).Color
    'Draw the star
    Ellipse Me.hdc, Stars(I).X, Stars(I).Y, Stars(I).X + Stars(I).Size, Stars(I).Y + Stars(I).Size

Next I

Me.Refresh

End Sub
