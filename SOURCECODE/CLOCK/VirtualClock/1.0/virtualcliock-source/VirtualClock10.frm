VERSION 5.00
Begin VB.Form VirtualClock10 
   Caption         =   "Virtual Clock 1.0"
   ClientHeight    =   7500
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7500
   FillColor       =   &H00E0E0E0&
   FillStyle       =   0  'Solid
   Icon            =   "VirtualClock10.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   375
   ScaleMode       =   2  'Point
   ScaleWidth      =   375
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3480
      Top             =   3240
   End
   Begin VB.Menu MnOptions 
      Caption         =   "&Options"
      Begin VB.Menu MnDial 
         Caption         =   "&Dial"
         Begin VB.Menu MnCircle 
            Caption         =   "Circle"
            Checked         =   -1  'True
         End
         Begin VB.Menu MnSquare 
            Caption         =   "Square"
         End
      End
      Begin VB.Menu MnMovement 
         Caption         =   "&Movement"
         Begin VB.Menu MnGraduations 
            Caption         =   "Graduations"
            Checked         =   -1  'True
         End
         Begin VB.Menu MnUniform 
            Caption         =   "Uniform"
         End
      End
      Begin VB.Menu MnHours 
         Caption         =   "&Hours"
         Begin VB.Menu Mn24h 
            Caption         =   "24h"
            Checked         =   -1  'True
         End
         Begin VB.Menu Mn12h 
            Caption         =   "12h"
         End
      End
   End
   Begin VB.Menu MnLanguage 
      Caption         =   "&Language"
      Begin VB.Menu MnEnglish 
         Caption         =   "English"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnGerman 
         Caption         =   "German"
      End
   End
   Begin VB.Menu MnAbout 
      Caption         =   "&About..."
   End
End
Attribute VB_Name = "VirtualClock10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PI = 3.14159265358979

Private Sub Form_Load()
    Icon = LoadPicture(App.Path + "\VirtualClock10.ico")
    ScaleMode = 2
    FillColor = &HE0E0E0
    FillStyle = 0
    BackColor = &HFFFFFF
    Language
End Sub

Private Sub Form_Resize()
    'I prefer to have small names
    Dim X As Double, Y As Double
    Dim Radius As Double
    Dim H As Integer, M As Integer, S As Integer
    Dim SH As String, SM As String, SS As String
    X = ScaleWidth
    Y = ScaleHeight
    H = Hour(Time)
    M = Minute(Time)
    S = Second(Time)
    Cls
    'The largest radius is choosen with the smallest of the ScaleWidth or'
    'the ScaleHeight not to draw out of the screen
    If X <= Y Then Radius = (X / 2) * (1 - 10 / 100) Else Radius = (Y / 2) * _
    (1 - 10 / 100)
    Call Dial(Radius) 'this function draws the dial
    If MnGraduations.Checked = True Then
        'The needles increases only when the previous needle has reached its
        'maximum (ex:minutes increases only when secondes reaches 60)
        Call Needle(Radius * (1 - 40 / 100), H / 12, QBColor(0), 3)
        Call Needle(Radius * (1 - 6 / 100), M / 60, QBColor(0), 2)
        Call Needle(Radius * (1 - 6 / 100), S / 60, QBColor(12), 1)
    Else
        'All the needles are moving on each second
        Call Needle(Radius * (1 - 40 / 100), H / 12 + M / 60 / 12 + _
        S / 60 / 60 / 12, QBColor(0), 3)
        Call Needle(Radius * (1 - 6 / 100), M / 60 + S / 60 / 60, QBColor(0), 2)
        Call Needle(Radius * (1 - 6 / 100), S / 60, QBColor(12), 1)
    End If
    'Here is being made the hour display for the form's title
    If Mn12h.Checked = True Then H = H Mod (12)
    SH = H
    SM = M
    SS = S
    If H < 10 Then SH = "0" & H
    If M < 10 Then SM = "0" & M
    If S < 10 Then SS = "0" & S
    If Mn12h.Checked = True Then
        If Hour(Time) > H Then
            Caption = SH & ":" & SM & ":" & SS & " PM"
        Else
            Caption = SH & ":" & SM & ":" & SS & " AM"
        End If
    Else
        Caption = SH & ":" & SM & ":" & SS
    End If
End Sub

Private Sub Mn12h_Click()
    If Mn12h.Checked = False Then
        Mn12h.Checked = True
        Mn24h.Checked = False
    End If
End Sub

Private Sub Mn24h_Click()
    If Mn24h.Checked = False Then
        Mn24h.Checked = True
        Mn12h.Checked = False
    End If
End Sub

Private Sub MnAbout_Click()
    If MnEnglish.Checked = True Then
        Call MsgBox("Virtual Clock 1.0" + (Chr(13) & Chr(10)) + "Contact: thomas_schilb@outlook.com", 64, "About")
    Else
        Call MsgBox("Virtual Clock 1.0" + (Chr(13) & Chr(10)) + "Kontakt: thomas_schilb@outlook.com", 64, "Über")
    End If
End Sub

Private Sub MnEnglish_Click()
    If MnEnglish.Checked = False Then
        MnEnglish.Checked = True
        MnGerman.Checked = False
        Language
    End If
End Sub

Private Sub MnGerman_Click()
    If MnGerman.Checked = False Then
        MnGerman.Checked = True
        MnEnglish.Checked = False
        Language
    End If
End Sub

Private Sub MnSquare_Click()
    If MnSquare.Checked = False Then
        MnSquare.Checked = True
        MnCircle.Checked = False
    End If
End Sub

Private Sub MnCircle_Click()
    If MnCircle.Checked = False Then
        MnCircle.Checked = True
        MnSquare.Checked = False
    End If
End Sub

Private Sub MnGraduations_Click()
    If MnGraduations.Checked = False Then
        MnGraduations.Checked = True
        MnUniform.Checked = False
    End If
End Sub

Private Sub MnUniform_Click()
If MnUniform.Checked = False Then
        MnUniform.Checked = True
        MnGraduations.Checked = False
    End If
End Sub

Private Sub Timer1_Timer()
    'I've add HasChanged function because I don't use AutoRedraw
    If HasChanged Then Form_Resize
End Sub

Private Static Function HasChanged() As Boolean
    'This function is made for the display not to have too many jerks
    Dim OldTime As Variant
    If Time = OldTime Then
        HasChanged = False
    Else
        OldTime = Time
        HasChanged = True
    End If
End Function

Private Function Needle(ByVal Radius As Double, ByVal Fraction As Double, ByVal Colour As Long, ByVal Thickness As Integer)
    Dim X As Double, Y As Double
    X = ScaleWidth
    Y = ScaleHeight
    DrawWidth = Thickness
    'Draws the parametered needle
    Line (X / 2, Y / 2)-(X / 2 + Radius * Cos(PI + PI / 2 + 2 * PI * Fraction), _
    Y / 2 + Radius * Sin(-PI / 2 - 2 * PI * Fraction)), Colour
    DrawWidth = 1
End Function

Private Function Dial(ByVal Radius As Double)
    Dim X As Double, Y As Double
    X = ScaleWidth
    Y = ScaleHeight
    DrawWidth = 1
    If MnCircle.Checked = True Then
        'For circle dial
        Circle (X / 2, Y / 2), Radius, QBColor(0)
        For i = 1 To 60
            If i Mod 5 = 0 Then DrawWidth = 2 Else DrawWidth = 1
            Line (X / 2 + (1 - 4 / 100) * Radius * Cos(PI + PI / 2 + 2 * PI * (i / 60)), _
            Y / 2 + (1 - 4 / 100) * Radius * Sin(-PI / 2 - 2 * PI * (i / 60)))- _
            (X / 2 + (1 - 0.5 / 100) * Radius * Cos(PI + PI / 2 + 2 * PI * (i / 60)), _
            Y / 2 + (1 - 0.5 / 100) * Radius * Sin(-PI / 2 - 2 * PI * (i / 60))), _
            QBColor(0)
        Next i
    Else
        'For square dial
        Line (X / 2 - Radius, Y / 2 - Radius)-(X / 2 + Radius, Y / 2 + Radius), QBColor(0), B
        For j = 0 To 1
            For i = -7 To 7
                If i Mod 5 = 0 Then DrawWidth = 2 Else DrawWidth = 1
                Line (X / 2 + (1 - 4 / 100) * Radius / Tan(PI / 2 - 2 * PI * (i + j * 30) / 60), _
                Y / 2 + (1 - 4 / 100) * Radius * Cos(j * PI))- _
                (X / 2 + (1 - 0.5 / 100) * Radius / Tan(PI / 2 - 2 * PI * (i + j * 30) / 60), _
                Y / 2 + (1 - 0.5 / 100) * Radius * Cos(j * PI)), _
                QBColor(0)
            Next i
            For i = 8 To 22
                If i Mod 5 = 0 Then DrawWidth = 2 Else DrawWidth = 1
                Line (X / 2 + (1 - 4 / 100) * Radius * Cos(j * PI), _
                Y / 2 + (1 - 4 / 100) * Radius / Tan(2 * PI * (i + j * 30) / 60))- _
                (X / 2 + (1 - 0.5 / 100) * Radius * Cos(j * PI), _
                Y / 2 + (1 - 0.5 / 100) * Radius / Tan(2 * PI * (i + j * 30) / 60)), _
                QBColor(0)
            Next i
        Next j
    End If
End Function

Private Function Language()
    If MnEnglish.Checked = True Then
        MnDial.Caption = "&Dial"
        MnCircle.Caption = "Circle"
        MnSquare.Caption = "Square"
        MnMovement.Caption = "&Movement"
        MnGraduations.Caption = "Graduations"
        MnUniform.Caption = "Uniform"
        MnHours.Caption = "&Hours"
        MnLanguage.Caption = "&Language"
        MnEnglish.Caption = "English"
        MnGerman.Caption = "German"
        MnAbout.Caption = "&About..."
    Else
        MnDial.Caption = "&Aussehen"
        MnCircle.Caption = "Kreis"
        MnSquare.Caption = "Quadrat"
        MnMovement.Caption = "&Bewegung"
        MnGraduations.Caption = "Crans"
        MnUniform.Caption = "Uniform"
        MnHours.Caption = "&Stunden"
        MnLanguage.Caption = "&Sprache"
        MnEnglish.Caption = "Englisch"
        MnGerman.Caption = "Deutsch"
        MnAbout.Caption = "&Über..."
    End If
End Function
