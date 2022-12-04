Attribute VB_Name = "ShiftRight"
Public Function SHR(ByVal Value As Long, ByVal ShiftCount As Long) As Long
    ' by Jost Schwider, jost@schwider.de, 20011010
    Select Case ShiftCount
        Case 0&:  SHR = Value
        Case 1&:  SHR = (Value And &HFFFFFFFE) \ &H2&
        Case 2&:  SHR = (Value And &HFFFFFFFC) \ &H4&
        Case 3&:  SHR = (Value And &HFFFFFFF8) \ &H8&
        Case 4&:  SHR = (Value And &HFFFFFFF0) \ &H10&
        Case 5&:  SHR = (Value And &HFFFFFFE0) \ &H20&
        Case 6&:  SHR = (Value And &HFFFFFFC0) \ &H40&
        Case 7&:  SHR = (Value And &HFFFFFF80) \ &H80&
        Case 8&:  SHR = (Value And &HFFFFFF00) \ &H100&
        Case 9&:  SHR = (Value And &HFFFFFE00) \ &H200&
        Case 10&: SHR = (Value And &HFFFFFC00) \ &H400&
        Case 11&: SHR = (Value And &HFFFFF800) \ &H800&
        Case 12&: SHR = (Value And &HFFFFF000) \ &H1000&
        Case 13&: SHR = (Value And &HFFFFE000) \ &H2000&
        Case 14&: SHR = (Value And &HFFFFC000) \ &H4000&
        Case 15&: SHR = (Value And &HFFFF8000) \ &H8000&
        Case 16&: SHR = (Value And &HFFFF0000) \ &H10000
        Case 17&: SHR = (Value And &HFFFE0000) \ &H20000
        Case 18&: SHR = (Value And &HFFFC0000) \ &H40000
        Case 19&: SHR = (Value And &HFFF80000) \ &H80000
        Case 20&: SHR = (Value And &HFFF00000) \ &H100000
        Case 21&: SHR = (Value And &HFFE00000) \ &H200000
        Case 22&: SHR = (Value And &HFFC00000) \ &H400000
        Case 23&: SHR = (Value And &HFF800000) \ &H800000
        Case 24&: SHR = (Value And &HFF000000) \ &H1000000
        Case 25&: SHR = (Value And &HFE000000) \ &H2000000
        Case 26&: SHR = (Value And &HFC000000) \ &H4000000
        Case 27&: SHR = (Value And &HF8000000) \ &H8000000
        Case 28&: SHR = (Value And &HF0000000) \ &H10000000
        Case 29&: SHR = (Value And &HE0000000) \ &H20000000
        Case 30&: SHR = (Value And &HC0000000) \ &H40000000
        Case 31&: SHR = CBool(Value And &H80000000)
    End Select
End Function



