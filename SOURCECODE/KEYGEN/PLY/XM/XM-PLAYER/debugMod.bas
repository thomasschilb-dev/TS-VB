Attribute VB_Name = "debugMod"
Dim DebugString As String
Public cd As New CDialog

Public Sub addlog(temp As String)
       DebugString = DebugString & vbNewLine & temp
End Sub

Public Sub savelog()
       Dim fil As Integer
       fil = FreeFile
       
       Open App.Path & "\log.txt" For Output As #fil
            Print #fil, DebugString
       Close #fil
End Sub

Public Sub openlog()
On Error Resume Next 'if the log is empty :)
       Dim fil As Integer
       Dim tmpstr As String
       fil = FreeFile

       Open App.Path & "\log.txt" For Input As #fil
            Line Input #fil, tmpstr ' preventing that an empty line is added every time this sub is called
            Do
              Line Input #fil, tmpstr
              DebugString = DebugString & vbNewLine & tmpstr
            Loop Until EOF(fil)
       Close #fil
End Sub

