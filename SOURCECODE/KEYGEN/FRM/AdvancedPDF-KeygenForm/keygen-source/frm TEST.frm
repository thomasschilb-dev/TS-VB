VERSION 5.00
Begin VB.Form frmTEST 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFEEFF&
   BorderStyle     =   0  'None
   Caption         =   "Advance PDF Without Acrobet Reader"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm TEST.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm TEST.frx":57E2
   ScaleHeight     =   4410
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblSTATUS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF00FF&
      Height          =   225
      Left            =   600
      TabIndex        =   1
      Top             =   4080
      Width           =   4740
   End
   Begin VB.Label lblCLOSE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Left            =   5760
      TabIndex        =   0
      ToolTipText     =   "Close"
      Top             =   60
      Width           =   120
   End
End
Attribute VB_Name = "frmTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objPDF As PDFClass

Private Sub AddPageNumber(objPDF As PDFClass, pageNumber As Integer)
    Dim sPageInfo As String
    Dim fontSize As Double
    Dim margin As Double
    
    fontSize = 10       'Size of font to use
    margin = 40         'Size of margin (left, right, bottom)
    
    ' Set what we want to print for page info
    sPageInfo = "Page " & pageNumber
    
    ' Should save these settings and change them back for more robust code
    objPDF.PDFSetTextColor = vbBlack
    objPDF.PDFSetAlignement = ALIGN_RIGHT
    objPDF.PDFSetFont FONT_ARIAL, CInt(fontSize), FONT_NORMAL
    objPDF.PDFSetFill = False
    
    ' Uncomment the below line if you want to see how our formating works
    'objPDF.PDFSetBorder = BORDER_ALL
    
    ' Draw the page number at the bottom of the page to the right
    objPDF.PDFCell sPageInfo, _
        margin, _
        objPDF.PDFGetPageHeight - margin - fontSize, _
        objPDF.PDFGetPageWidth - (margin * 2), _
        fontSize
End Sub

Private Sub Form_Load()
Line (0, 0)-(Me.ScaleWidth, 0), vbBlack
Line (0, 0)-(0, Me.ScaleHeight), vbBlack
Line (0, Me.ScaleHeight - 15)-(Me.ScaleWidth, Me.ScaleHeight - 15), vbBlack
Line (Me.ScaleWidth - 15, 0)-(Me.ScaleWidth - 15, Me.ScaleHeight), vbBlack
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    DragWindow Me
End If
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblCLOSE.FontBold = False
End Sub

Private Sub lblADVANCE_Click()
    ' Create a simple PDF file using the mjwPDF class
    Set objPDF = New PDFClass
    
    ' Show Status
    lblSTATUS.Caption = "Creating PDF File, Please Wait ..."
    Screen.MousePointer = 11
    
    ' Set the PDF title and filename
    objPDF.PDFTitle = "Test PDF Document"
    objPDF.PDFFileName = App.Path & "\test.pdf"
    
    ' We must tell the class where the PDF fonts are located
    objPDF.PDFLoadFont = App.Path & "\Fonts"
    
    ' Set the file properties
    objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
    objPDF.PDFFormatPage = FORMAT_A4
    objPDF.PDFOrientation = ORIENT_PORTRAIT
    objPDF.PDFSetUnit = UNIT_PT
    
    ' Lets us set see the bookmark pane when we view the PDF
    objPDF.PDFUseOutlines = True
    
    ' View the PDF file after we create it
    objPDF.PDFView = True
    
    ' Begin our PDF document
    objPDF.PDFBeginDoc
        ' Lets add a heading
        objPDF.PDFSetFont FONT_ARIAL, 15, FONT_BOLD
        objPDF.PDFSetDrawColor = vbRed
        objPDF.PDFSetTextColor = vbWhite
        objPDF.PDFSetAlignement = ALIGN_CENTER
        objPDF.PDFSetBorder = BORDER_ALL
        objPDF.PDFSetFill = True
        objPDF.PDFCell "A centered heading", 15, 15, _
            objPDF.PDFGetPageWidth - 30, 40
            
        ' Lets draw a dashed red square
        objPDF.PDFSetLineColor = vbRed
        objPDF.PDFSetFill = True
        objPDF.PDFSetLineStyle = pPDF_DASHDOT
        objPDF.PDFSetLineWidth = 1
        objPDF.PDFSetDrawMode = DRAW_NORMAL
        objPDF.PDFDrawPolygon Array(300, 150, 400, 150, 400, 250, 300, 250)
        
        ' Lets draw an elipse
        objPDF.PDFSetDrawColor = vbYellow
        objPDF.PDFSetLineColor = vbBlack
        objPDF.PDFSetLineStyle = pPDF_DASHDOT
        objPDF.PDFSetLineWidth = 1.25
        objPDF.PDFSetDrawMode = DRAW_DRAWBORDER
        objPDF.PDFDrawEllipse 300, 150, 75, 25
        
        AddPageNumber objPDF, 1
        
        'Lets add a bookmark to the start of page 1
        objPDF.PDFSetBookmark "A. Page 1", 0, 0
        
        'Now a bookmark half way down page 1
        objPDF.PDFSetBookmark "A1. Page 1 Halfway down", 1, 300
        
        'Now one at the end page 1
        objPDF.PDFSetBookmark "A2. End of Page 1", 1, 500
        
        'Another one a little further down and shows nesting
        objPDF.PDFSetBookmark "A2-Sub1.", 2, 800
        
        objPDF.PDFEndPage
        
        'Start page 2
        objPDF.PDFNewPage
        
        'Lets add an image to page 2
        objPDF.PDFImage App.Path & "\Back.jpg", _
            15, 15, 400, 300, "mailto:magic-world@email.com"
            
        'Lets add a bookmark to the start of page 2
        objPDF.PDFSetBookmark "Page 2", 0, 0
        
        'Now a bookmark just below the logo
        objPDF.PDFSetBookmark "Page 2 just below logo", 1, 75
        
        AddPageNumber objPDF, 2
    
    ' End our PDF document (this will save it to the filename)
    objPDF.PDFEndDoc

    Set objPDF = Nothing

    ' Show Status
    lblSTATUS.Caption = "PDF Created Successfully."
    Screen.MousePointer = 0
End Sub

Private Sub lblADVANCE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub lblCLOSE_Click()
Unload Me
End
End Sub

Private Sub lblCLOSE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCLOSE.FontBold = True
End Sub

Private Sub lblSIMPLE_Click()
    ' Create a simple PDF file using the PDF Class
    Set objPDF = New PDFClass
    
    ' Show Status
    lblSTATUS.Caption = "Creating PDF File, Please Wait ..."
    Screen.MousePointer = 11
    
    ' Set the PDF title and filename
    objPDF.PDFTitle = "Test PDF Document"
    objPDF.PDFFileName = App.Path & "\test.pdf"
    
    ' We must tell the class where the PDF fonts are located
    objPDF.PDFLoadFont = App.Path & "\Fonts"
    
    ' View the PDF file after we create it
    objPDF.PDFView = True
    
    ' Begin our PDF document
    objPDF.PDFBeginDoc
    
    ' Set the font name, size, and style
    objPDF.PDFSetFont FONT_ARIAL, 15, FONT_BOLD
        
    ' Set the text color
    objPDF.PDFSetTextColor = vbBlue
    
    ' Set the text we want to print
    objPDF.PDFTextOut "Hello, World! From PDF Class"
    
    ' End our PDF document (this will save it to the filename)
    objPDF.PDFEndDoc
    
    Set objPDF = Nothing
    
    ' Show Status
    lblSTATUS.Caption = "PDF Created Successfully."
    Screen.MousePointer = 0
End Sub

Private Sub lblSIMPLE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub
