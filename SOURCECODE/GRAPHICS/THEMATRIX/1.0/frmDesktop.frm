VERSION 5.00
Begin VB.Form frmDesktop 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10125
   BeginProperty Font 
      Name            =   "Wingdings 2"
      Size            =   9.75
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00004000&
   Icon            =   "frmDesktop.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   443
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMatrix 
      Interval        =   20
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Lookups & variables
Private xm_alPreviousChar() As Long
Private xm_alCurrentRow() As Long   'Current row index for each column
Private xm_alPosCol() As Long       'X position, screen coordinate for a column
Private xm_alPosRow() As Long       'Y position, screen coordinate for a row
Private xm_alColumnInitd() As Long
Private xm_lCountCol As Long        'Number of columns, depends on resolution
Private xm_lCountRow As Long        'Number of rows, depends on resolution
Private xm_bInit As Boolean

Private Sub FormOnTop(frmForm As Form, Optional blnTop As Boolean = True)
    If blnTop Then
        SetWindowPos frmForm.Hwnd, -1, 0, 0, 0, 0, 3
    Else
        SetWindowPos frmForm.Hwnd, -2, 0, 0, 0, 0, 3
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'Press Escape, ask for exit
    If KeyAscii = 27 Then
        If MsgBox("Really want to exit?", vbQuestion + vbYesNo, "The Matrix 1.0") = vbYes Then
            tmrMatrix.Enabled = False
            DoEvents
            Unload frmStrip
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    'Position the form.
    Me.Left = 0
    Me.Top = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height - 15
    'Load the small form strip in the upper left corner
    frmStrip.Show
    'Make it on top of all windows
    FormOnTop frmStrip
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase xm_alPreviousChar, xm_alCurrentRow, xm_alPosCol, xm_alPosRow, xm_alColumnInitd
End Sub

Private Sub tmrMatrix_Timer()
    Const UPDATE_CHANCE    As Long = 4  'percent
    Const NON_SPACE_CHANCE As Long = 10 'percent
    Const SPACE_CHANCE     As Long = 3 'percent
    Const CHARACTER_WIDTH  As Long = 12 'pixels
    Const CHARACTER_HEIGHT As Long = 15 'pixels
    Const TIMER_INTERVAL   As Long = 30 'miliseconds
        
    'Stop timer while we draw
    tmrMatrix.Enabled = False
    
    Dim lX As Long, lDice As Long
    DoEvents
    
    'If it's the first run, initialize all info
    If xm_bInit = False Then
        tmrMatrix.Interval = TIMER_INTERVAL
        Randomize Timer
        
        'Find out how many columns & rows
        xm_lCountCol = Screen.Width \ CHARACTER_WIDTH \ 15 + 1
        xm_lCountRow = Screen.Height \ CHARACTER_HEIGHT \ 15 + 1
        
        'Resize arrays to hold data
        ReDim xm_alPreviousChar(xm_lCountCol - 1)
        ReDim xm_alPosCol(xm_lCountCol - 1)
        ReDim xm_alPosRow(xm_lCountRow - 1)
        ReDim xm_alColumnInitd(xm_lCountCol - 1)
        ReDim xm_alCurrentRow(xm_lCountCol - 1)
        
        'Fill in the data
        For lX = 0 To xm_lCountCol - 1
            xm_alPosCol(lX) = lX * CHARACTER_WIDTH
            xm_alPreviousChar(lX) = -1
        Next lX
        For lX = 0 To xm_lCountRow - 1
            xm_alPosRow(lX) = lX * CHARACTER_HEIGHT
        Next lX
        
        'Done
        xm_bInit = True
    End If
    
    'The main draw parts...
    'Go through every column
    For lX = 0 To xm_lCountCol - 1
    
        'If column is not updating
        If xm_alPreviousChar(lX) = -1 Then
            'Roll dice to see if it should be updated on the next check
            lDice = Rnd * 100 + 1
            If lDice < UPDATE_CHANCE Then
                xm_alPreviousChar(lX) = 0
            End If
            
        Else 'Its updating
        
            'First, blank out area to be drawn
            Me.Line (xm_alPosCol(lX), xm_alPosRow(xm_alCurrentRow(lX)))-(xm_alPosCol(lX) + CHARACTER_WIDTH, xm_alPosRow(xm_alCurrentRow(lX)) + CHARACTER_HEIGHT), Me.BackColor, BF
  'If VB complains about this line when you press F5, delete this dash ^, click on another line to generate an error
  'then replace it with another dash. I don't know why that works, maybe coz I used the Line function like in VB3?
            
            'Draw over the previous char with the darker color
            If xm_alCurrentRow(lX) > 0 Then
                'Select a random darker color
                lDice = Rnd * 100 + 1
                If lDice < 33 Then
                    Me.ForeColor = &H36B72E
                ElseIf lDice < 66 Then
                    Me.ForeColor = &H172C13
                Else
                    Me.ForeColor = &H4000&
                End If
                Me.CurrentX = xm_alPosCol(lX) + (CHARACTER_WIDTH - TextWidth(Chr$(xm_alPreviousChar(lX)))) \ 2
                Me.CurrentY = xm_alPosRow(xm_alCurrentRow(lX) - 1)
                Me.Print Chr$(xm_alPreviousChar(lX))
            End If
            
            'If previous character is an empty space
            If xm_alPreviousChar(lX) = 0 Then
            
                'If it's never had a chance to draw characters before
                If xm_alColumnInitd(lX) = 0 Then
                
                    'Roll dice to see if it should be changed to non space
                    lDice = Rnd * 100 + 1
                    If lDice < NON_SPACE_CHANCE Then
                        xm_alColumnInitd(lX) = 1
                        xm_alPreviousChar(lX) = Rnd * 93 + 33
                        xm_alCurrentRow(lX) = -1
                    End If
                End If
                
            Else 'Previous character isn't an empty space
            
                'Roll dice to see if it should be changed to space
                lDice = Rnd * 100 + 1
                If lDice < SPACE_CHANCE Then
                    'Change to space
                    xm_alPreviousChar(lX) = 0
                    
                Else 'It's not changing to space
                
                    'Draw it, bright color
                    lDice = Rnd * 93 + 33
                    'Me.ForeColor = &H9000&
                    Me.ForeColor = &HFCFFFC
                    Me.CurrentX = xm_alPosCol(lX) + (CHARACTER_WIDTH - Me.TextWidth(Chr$(lDice))) \ 2
                    Me.CurrentY = xm_alPosRow(xm_alCurrentRow(lX))
                    Me.Print Chr$(lDice)
                    
                    'Save the drawn char value
                    xm_alPreviousChar(lX) = lDice
                End If
            End If
            
            'Increment the row index for this column
            xm_alCurrentRow(lX) = xm_alCurrentRow(lX) + 1
            
            'If we have reached the last row
            If xm_alCurrentRow(lX) = xm_lCountRow Then
                'Reset it
                xm_alCurrentRow(lX) = 0
                xm_alPreviousChar(lX) = -1
                xm_alColumnInitd(lX) = 0
            End If
            
        End If
    Next lX
    
    'Restart timer
    tmrMatrix.Enabled = True
End Sub
