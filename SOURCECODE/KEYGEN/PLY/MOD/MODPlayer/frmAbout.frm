VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2715
   Icon            =   "frmAbout.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      Caption         =   "Copyright (c) 2002 Mark Lu"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1920
   End
   Begin VB.Label lblWebsite 
      AutoSize        =   -1  'True
      Caption         =   "Website: http://marklu.cjb.net"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2130
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "Email: marklu1990@hotmail.com"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      Caption         =   "Author: Mark Lu"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1140
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "MOD, IT, XM, and S3M Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2160
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
