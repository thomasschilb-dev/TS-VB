VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISoundStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum SND_RESULT
    SND_ERR_SUCCESS
    SND_ERR_INVALID_SOURCE
    SND_ERR_INVALID_OUTPUT
    SND_ERR_INTERNAL
    SND_ERR_OUT_OF_RANGE
    SND_ERR_END_OF_STREAM
    SND_ERR_INVALID_TAG
    SND_ERR_INVALID_PARAM
    SND_ERR_TOO_BIG
    SND_ERR_NEED_MORE
    SND_ERR_UNKNOWN
End Enum

Public Enum SND_SEEK_MODE
    SND_SEEK_PERCENT
    SND_SEEK_SECONDS
End Enum

Public Property Get StreamInfo( _
) As ISoundInfo

End Property

Public Property Get EndOfStream( _
) As Boolean

End Property

Public Property Get Description( _
) As String

End Property

Public Function Extensions( _
) As String()

End Function

Public Function StreamOpen( _
    ByVal SOURCE As String _
) As SND_RESULT

End Function

Public Function StreamClose( _
) As SND_RESULT

End Function

Public Function StreamSeek( _
    ByVal value As Long, _
    ByVal seek_mode As SND_SEEK_MODE _
) As SND_RESULT

End Function

Public Function StreamRead( _
    ByVal buffer_ptr As Long, _
    ByVal buffer_len As Long, _
    ByRef buffer_read As Long _
) As SND_RESULT

End Function
