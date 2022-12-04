Attribute VB_Name = "CinVB"

Type word 'Dies simuliert ein C word
    AByte As Byte
    BByte As Byte
End Type

'Diese Funktion liest und korrigiert
'(words) die aus einem ModFile gelsen wurden
Function MakeModWord(w As word) As Long 'For the Samples
Dim A As Long
Dim B As Long
A = w.AByte
B = w.BByte
MakeModWord = ((A * 256) + B) * 2
End Function



