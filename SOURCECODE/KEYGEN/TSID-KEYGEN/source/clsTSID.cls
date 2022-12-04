VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTSID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class is a very minor adaption of the work of Mark van Cuijk
' http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=41067&lngWId=1

Private lngTrack As Long
Private arrLongConversion(4) As Long
Private arrSplit64(63) As Byte

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647

Private Const S11 = 7
Private Const S12 = 12
Private Const S13 = 17
Private Const S14 = 22
Private Const S21 = 5
Private Const S22 = 9
Private Const S23 = 14
Private Const S24 = 20
Private Const S31 = 4
Private Const S32 = 11
Private Const S33 = 16
Private Const S34 = 23
Private Const S41 = 6
Private Const S42 = 10
Private Const S43 = 15
Private Const S44 = 21

Private Function TSIDRound(strRound As String, a As Long, b As Long, C As Long, d As Long, X As Long, S As Long, ac As Long) As Long
    Select Case strRound
        Case Is = "FF"
            a = TSIDLongAdd4(a, (b And C) Or (Not (b) And d), X, ac)
            a = TSIDRotate(a, S)
            a = TSIDLongAdd(a, b)
        Case Is = "GG"
            a = TSIDLongAdd4(a, (b And d) Or (C And Not (d)), X, ac)
            a = TSIDRotate(a, S)
            a = TSIDLongAdd(a, b)
        Case Is = "HH"
            a = TSIDLongAdd4(a, b Xor C Xor d, X, ac)
            a = TSIDRotate(a, S)
            a = TSIDLongAdd(a, b)
        Case Is = "II"
            a = TSIDLongAdd4(a, C Xor (b Or Not (d)), X, ac)
            a = TSIDRotate(a, S)
            a = TSIDLongAdd(a, b)
    End Select
End Function

Private Function TSIDRotate(lngValue As Long, lngBits As Long) As Long
Dim lngSign As Long
Dim lngI As Long
    lngBits = (lngBits Mod 32)
    
    If lngBits = 0 Then TSIDRotate = lngValue: Exit Function
    
    For lngI = 1 To lngBits
        lngSign = lngValue And &HC0000000
        lngValue = (lngValue And &H3FFFFFFF) * 2
        lngValue = lngValue Or ((lngSign < 0) And 1) Or (CBool(lngSign And &H40000000) And &H80000000)
    Next
    
    TSIDRotate = lngValue
End Function

Private Function TRID() As String

    Dim sngNum As Single, lngnum As Long
    Dim strResult As String
   
    sngNum = Rnd(2147483648#)
    strResult = CStr(sngNum)
    
    strResult = Replace(strResult, "0.", "")
    strResult = Replace(strResult, ".", "")
    strResult = Replace(strResult, "E-", "")
    
    TRID = strResult

End Function

Private Function TSID64Split(lngLength As Long, bytBuffer() As Byte) As String

    Dim lngBytesTotal As Long, lngBytesToAdd As Long
    Dim intLoop As Integer, intLoop2 As Integer, lngTrace As Long
    Dim intInnerLoop As Integer, intLoop3 As Integer
    
    lngBytesTotal = lngTrack Mod 64
    lngBytesToAdd = 64 - lngBytesTotal
    lngTrack = (lngTrack + lngLength)
    
    If lngLength >= lngBytesToAdd Then
        For intLoop = 0 To lngBytesToAdd - 1
            arrSplit64(lngBytesTotal + intLoop) = bytBuffer(intLoop)
        Next intLoop
        
        TSIDConversion arrSplit64
        
        lngTrace = (lngLength) Mod 64

        For intLoop2 = lngBytesToAdd To lngLength - intLoop - lngTrace Step 64
            For intInnerLoop = 0 To 63
                arrSplit64(intInnerLoop) = bytBuffer(intLoop2 + intInnerLoop)
            Next intInnerLoop
            
            TSIDConversion arrSplit64
        
        Next intLoop2
        
        lngBytesTotal = 0
    Else
    
      intLoop2 = 0
    
    End If
    
    For intLoop3 = 0 To lngLength - intLoop2 - 1
        
        arrSplit64(lngBytesTotal + intLoop3) = bytBuffer(intLoop2 + intLoop3)
    
    Next intLoop3
     
End Function

Private Function TSIDStringArray(strInput As String) As Byte()
    
    Dim intLoop As Integer
    Dim bytBuffer() As Byte
    ReDim bytBuffer(Len(strInput))
    
    For intLoop = 0 To Len(strInput) - 1
        bytBuffer(intLoop) = Asc(Mid(strInput, intLoop + 1, 1))
    Next intLoop
    
    TSIDStringArray = bytBuffer
    
End Function

Private Sub TSIDConversion(bytBuffer() As Byte)

    Dim X(16) As Long, a As Long
    Dim b As Long, C As Long
    Dim d As Long
    
    a = arrLongConversion(1)
    b = arrLongConversion(2)
    C = arrLongConversion(3)
    d = arrLongConversion(4)
    
    TSIDDecode 64, X, bytBuffer
    
    TSIDRound "FF", a, b, C, d, X(0), S11, -680876936
    TSIDRound "FF", d, a, b, C, X(1), S12, -389564586
    TSIDRound "FF", C, d, a, b, X(2), S13, 606105819
    TSIDRound "FF", b, C, d, a, X(3), S14, -1044525330
    TSIDRound "FF", a, b, C, d, X(4), S11, -176418897
    TSIDRound "FF", d, a, b, C, X(5), S12, 1200080426
    TSIDRound "FF", C, d, a, b, X(6), S13, -1473231341
    TSIDRound "FF", b, C, d, a, X(7), S14, -45705983
    TSIDRound "FF", a, b, C, d, X(8), S11, 1770035416
    TSIDRound "FF", d, a, b, C, X(9), S12, -1958414417
    TSIDRound "FF", C, d, a, b, X(10), S13, -42063
    TSIDRound "FF", b, C, d, a, X(11), S14, -1990404162
    TSIDRound "FF", a, b, C, d, X(12), S11, 1804603682
    TSIDRound "FF", d, a, b, C, X(13), S12, -40341101
    TSIDRound "FF", C, d, a, b, X(14), S13, -1502002290
    TSIDRound "FF", b, C, d, a, X(15), S14, 1236535329

    TSIDRound "GG", a, b, C, d, X(1), S21, -165796510
    TSIDRound "GG", d, a, b, C, X(6), S22, -1069501632
    TSIDRound "GG", C, d, a, b, X(11), S23, 643717713
    TSIDRound "GG", b, C, d, a, X(0), S24, -373897302
    TSIDRound "GG", a, b, C, d, X(5), S21, -701558691
    TSIDRound "GG", d, a, b, C, X(10), S22, 38016083
    TSIDRound "GG", C, d, a, b, X(15), S23, -660478335
    TSIDRound "GG", b, C, d, a, X(4), S24, -405537848
    TSIDRound "GG", a, b, C, d, X(9), S21, 568446438
    TSIDRound "GG", d, a, b, C, X(14), S22, -1019803690
    TSIDRound "GG", C, d, a, b, X(3), S23, -187363961
    TSIDRound "GG", b, C, d, a, X(8), S24, 1163531501
    TSIDRound "GG", a, b, C, d, X(13), S21, -1444681467
    TSIDRound "GG", d, a, b, C, X(2), S22, -51403784
    TSIDRound "GG", C, d, a, b, X(7), S23, 1735328473
    TSIDRound "GG", b, C, d, a, X(12), S24, -1926607734
  
    TSIDRound "HH", a, b, C, d, X(5), S31, -378558
    TSIDRound "HH", d, a, b, C, X(8), S32, -2022574463
    TSIDRound "HH", C, d, a, b, X(11), S33, 1839030562
    TSIDRound "HH", b, C, d, a, X(14), S34, -35309556
    TSIDRound "HH", a, b, C, d, X(1), S31, -1530992060
    TSIDRound "HH", d, a, b, C, X(4), S32, 1272893353
    TSIDRound "HH", C, d, a, b, X(7), S33, -155497632
    TSIDRound "HH", b, C, d, a, X(10), S34, -1094730640
    TSIDRound "HH", a, b, C, d, X(13), S31, 681279174
    TSIDRound "HH", d, a, b, C, X(0), S32, -358537222
    TSIDRound "HH", C, d, a, b, X(3), S33, -722521979
    TSIDRound "HH", b, C, d, a, X(6), S34, 76029189
    TSIDRound "HH", a, b, C, d, X(9), S31, -640364487
    TSIDRound "HH", d, a, b, C, X(12), S32, -421815835
    TSIDRound "HH", C, d, a, b, X(15), S33, 530742520
    TSIDRound "HH", b, C, d, a, X(2), S34, -995338651
 
    TSIDRound "II", a, b, C, d, X(0), S41, -198630844
    TSIDRound "II", d, a, b, C, X(7), S42, 1126891415
    TSIDRound "II", C, d, a, b, X(14), S43, -1416354905
    TSIDRound "II", b, C, d, a, X(5), S44, -57434055
    TSIDRound "II", a, b, C, d, X(12), S41, 1700485571
    TSIDRound "II", d, a, b, C, X(3), S42, -1894986606
    TSIDRound "II", C, d, a, b, X(10), S43, -1051523
    TSIDRound "II", b, C, d, a, X(1), S44, -2054922799
    TSIDRound "II", a, b, C, d, X(8), S41, 1873313359
    TSIDRound "II", d, a, b, C, X(15), S42, -30611744
    TSIDRound "II", C, d, a, b, X(6), S43, -1560198380
    TSIDRound "II", b, C, d, a, X(13), S44, 1309151649
    TSIDRound "II", a, b, C, d, X(4), S41, -145523070
    TSIDRound "II", d, a, b, C, X(11), S42, -1120210379
    TSIDRound "II", C, d, a, b, X(2), S43, 718787259
    TSIDRound "II", b, C, d, a, X(9), S44, -343485551
    
    arrLongConversion(1) = TSIDLongAdd(arrLongConversion(1), a)
    arrLongConversion(2) = TSIDLongAdd(arrLongConversion(2), b)
    arrLongConversion(3) = TSIDLongAdd(arrLongConversion(3), C)
    arrLongConversion(4) = TSIDLongAdd(arrLongConversion(4), d)
    
End Sub

Private Function TSIDLongAdd(lngVal1 As Long, lngVal2 As Long) As Long
    
    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (lngVal1 And &HFFFF&) + (lngVal2 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((lngVal1 And &HFFFF0000) \ 65536) + ((lngVal2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    
    TSIDLongAdd = TSIDLongConversion((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))

End Function

Private Function TSIDLongAdd4(lngVal1 As Long, lngVal2 As Long, lngVal3 As Long, lngVal4 As Long) As Long
    
    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (lngVal1 And &HFFFF&) + (lngVal2 And &HFFFF&) + (lngVal3 And &HFFFF&) + (lngVal4 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((lngVal1 And &HFFFF0000) \ 65536) + ((lngVal2 And &HFFFF0000) \ 65536) + ((lngVal3 And &HFFFF0000) \ 65536) + ((lngVal4 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    TSIDLongAdd4 = TSIDLongConversion((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))

End Function

Private Sub TSIDDecode(intLength As Integer, lngOutBuffer() As Long, bytInBuffer() As Byte)
    
    Dim intDblIndex As Integer
    Dim intByteIndex As Integer
    Dim dblSum As Double
    
    intDblIndex = 0
    
    For intByteIndex = 0 To intLength - 1 Step 4
        
        dblSum = bytInBuffer(intByteIndex) + bytInBuffer(intByteIndex + 1) * 256# + bytInBuffer(intByteIndex + 2) * 65536# + bytInBuffer(intByteIndex + 3) * 16777216#
        lngOutBuffer(intDblIndex) = TSIDLongConversion(dblSum)
        intDblIndex = (intDblIndex + 1)
    
    Next intByteIndex

End Sub

Private Function TSIDLongConversion(dblValue As Double) As Long
    
    If dblValue < 0 Or dblValue >= OFFSET_4 Then Error 6
        
    If dblValue <= MAXINT_4 Then
        TSIDLongConversion = dblValue
    Else
        TSIDLongConversion = dblValue - OFFSET_4
    End If
        
End Function

Private Sub TSIDFinish()
Dim dblBits As Double
Dim arrPadding(72) As Byte
Dim lngBytesBuffered As Long
    
    arrPadding(0) = &H80
    dblBits = lngTrack * 8
    
    lngBytesBuffered = lngTrack Mod 64
    
    If lngBytesBuffered <= 56 Then
        TSID64Split (56 - lngBytesBuffered), arrPadding
    Else
        TSID64Split (120 - lngTrack), arrPadding
    End If
    
    
    arrPadding(0) = TSIDLongConversion(dblBits) And &HFF&
    arrPadding(1) = TSIDLongConversion(dblBits) \ 256 And &HFF&
    arrPadding(2) = TSIDLongConversion(dblBits) \ 65536 And &HFF&
    arrPadding(3) = TSIDLongConversion(dblBits) \ 16777216 And &HFF&
    arrPadding(4) = 0
    arrPadding(5) = 0
    arrPadding(6) = 0
    arrPadding(7) = 0
    
    TSID64Split 8, arrPadding
End Sub

Private Function TSIDStringChange(lngnum As Long) As String
Dim bytA As Byte
Dim bytB As Byte
Dim bytC As Byte
Dim bytD As Byte
     bytA = lngnum And &HFF&
     If bytA < 16 Then
         TSIDStringChange = "0" & Hex(bytA)
     Else
         TSIDStringChange = Hex(bytA)
     End If
            
     bytB = (lngnum And &HFF00&) \ 256
     If bytB < 16 Then
         TSIDStringChange = TSIDStringChange & "0" & Hex(bytB)
     Else
         TSIDStringChange = TSIDStringChange & Hex(bytB)
     End If
     
     bytC = (lngnum And &HFF0000) \ 65536
     If bytC < 16 Then
         TSIDStringChange = TSIDStringChange & "0" & Hex(bytC)
     Else
         TSIDStringChange = TSIDStringChange & Hex(bytC)
     End If
    
     If lngnum < 0 Then
         bytD = ((lngnum And &H7F000000) \ 16777216) Or &H80&
     Else
         bytD = (lngnum And &HFF000000) \ 16777216
     End If
     
     If bytD < 16 Then
         TSIDStringChange = TSIDStringChange & "0" & Hex(bytD)
     Else
         TSIDStringChange = TSIDStringChange & Hex(bytD)
     End If
End Function

Private Function TSIDValue() As String
    TSIDValue = LCase(TSIDStringChange(arrLongConversion(1)) & TSIDStringChange(arrLongConversion(2)) & TSIDStringChange(arrLongConversion(3)) & TSIDStringChange(arrLongConversion(4)))
End Function

Public Function CalculateTSID(strMessage As String) As String
Dim bytBuffer() As Byte
    bytBuffer = TSIDStringArray(strMessage)
    TSIDStart
    TSID64Split Len(strMessage), bytBuffer
    TSIDFinish
    CalculateTSID = TSIDValue
End Function

Private Sub TSIDStart()
    lngTrack = 0
    arrLongConversion(1) = TSIDLongConversion(1732584193#)
    arrLongConversion(2) = TSIDLongConversion(4023233417#)
    arrLongConversion(3) = TSIDLongConversion(2562383102#)
    arrLongConversion(4) = TSIDLongConversion(271733878#)
End Sub

