VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISoundInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get SamplesPerSecond() As Long

End Property

Public Property Get Channels() As Integer

End Property

Public Property Get BitsPerSample() As Integer

End Property

Public Property Get BitsPerSecond() As Long

End Property

Public Property Get Position() As Long

End Property

Public Property Get Duration() As Long

End Property

Public Property Get Tags() As StreamTags

End Property
