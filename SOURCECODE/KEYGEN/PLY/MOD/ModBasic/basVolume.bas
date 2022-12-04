Attribute VB_Name = "basVolume"
'basVolume (c) Olaf R. 2002

Option Explicit

' Transformation volume value into DirectSound value.
'
' MOD-Volume [0...63]  ->  DSound-Volume [-10000...0]
'

Public Function Vol_linear(ByVal bVol As Byte) As Long
'linear transformation
  Vol_linear = (CLng(bVol) * 10000) / 63 - 10000
End Function

Public Function Vol_log10(ByVal bVol As Byte) As Long
'logarithmic transformation (base 10)
  Vol_log10 = Int((Log10(bVol + 1) * 10000) / 1.80618) - 10000
End Function

Public Static Function Log10(ByVal X As Single) As Single
  Log10 = Log(X) / 2.30258509299405  'LOG_10
End Function

