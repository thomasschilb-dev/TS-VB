Attribute VB_Name = "ShiftLeft"
Public Static Function SHL(ByVal Value As Long, ByVal ShiftCount As Long) As Long
    ' by Jost Schwider, jost@schwider.de, 20010928
    Dim Pow2(0 To 31) As Long
    Dim i As Long
    Dim mask As Long
    
    Select Case ShiftCount
        Case 1 To 31
            
            'Ggf. Initialisieren:
            If i = 0 Then
                Pow2(0) = 1
                For i = 1 To 30
                    Pow2(i) = 2 * Pow2(i - 1)
                Next i
            End If
            
            'Los gehts:
            mask = Pow2(31 - ShiftCount)
            If Value And mask Then
                SHL = (Value And (mask - 1)) * Pow2(ShiftCount) Or &H80000000
            Else
                SHL = (Value And (mask - 1)) * Pow2(ShiftCount)
            End If
            
        Case 0
            
            SHL = Value
            
    End Select
End Function


