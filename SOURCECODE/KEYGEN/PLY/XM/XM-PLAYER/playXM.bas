Attribute VB_Name = "playXM"
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private perfFreq                As Currency
Private timeStart               As Currency
Private TimeEnd                 As Currency
Private TimeElapsed             As Currency
Private EndTime                 As Currency
Private TempTime                As Currency
Private parms                   As String
Private skiptime                As Single

Public channelsort()            As Integer
Public StopSong                 As Boolean
Public warning16                As Boolean

Dim patt                        As Integer
Dim dochange                    As Boolean
Dim CurrentPattern              As Integer
       Dim Tick As Integer
       Dim Row  As Integer

Public Sub play()
       StopSong = False
       
       ' reset all values. should probably do this somewhere else, but what the hell :)
     '  CurrentPattern = 0
      ' patt = xm.Order(0)
       Row = 0
       Tick = 0
       
       Do
       
       Wait
       
       Tick = Tick + 1
       If Tick >= xm.tempo Then
          UpdateRow Row
          
          If dochange = True Then
             CurrentPattern = CurrentPattern + 1
             Row = -1
             dochange = False
          End If
          
          Tick = 0
          Row = Row + 1
          
          
          If Row >= pattern(xm.Order(CurrentPattern)).Rows Then CurrentPattern = CurrentPattern + 1: Row = 0
          patt = xm.Order(CurrentPattern)
       Else
         'update effects
          
       End If
       
       If Not Form2.length.value = CurrentPattern Then If CurrentPattern <= Form2.length.Max Then Form2.length.value = CurrentPattern
       Loop Until StopSong = True
End Sub

Private Sub UpdateRow(Row As Integer)
        On Error Resume Next
        Dim i As Integer
        Dim sample As Integer
        
        For i = 0 To xm.channels - 1
            If CurrentPattern > xm.patterns Then CurrentPattern = 0: patt = xm.Order(0)
            With pattern(patt).pattern(Row, i)
                 
                 If Not .instrument = 0 And Not .note = 97 Then sample = ih2(.instrument).sample(.note)
                 
                 If .note = 97 Then channelBuf(i).Stop
                 
                 If .note > 0 And .note < 97 And Not sample = 0 And sh(sample).length > 0 Then
                    If Not channelsort(i) = sample And Not sample = 0 Then
                       copysample i, sample
                    End If
                    
                    channelsort(i) = sample
                 
                    
                    frequency = GetFreq(.note, sample)
                 
                    If .note > 0 Then channelBuf(i).SetCurrentPosition 0: channelBuf(i).Stop
               
                    If Not frequency > DSBFREQUENCY_MAX Then channelBuf(i).SetFrequency frequency
                
                    channelBuf(i).SetVolume -1000
                     
                    If Form2.DaLoop.value = 1 Then
                       If sh(sample).loopend = 0 Then channelBuf(i).play DSBPLAY_DEFAULT Else channelBuf(i).play DSBPLAY_LOOPING
                    Else
                       channelBuf(i).play DSBPLAY_DEFAULT
                    End If
                 
                 End If
                 
                 If .effect = 13 Then
                     dochange = True
                 End If
                 
                 If .effect = 15 Then  ' should give this piece-of-code a sub of his own someday. maybe when he gets older :)
                     Debug.Print "tempo change " & .parameter
                     If Not .parameter > 31 Then
                        xm.tempo = .parameter
                     Else
                        xm.BPM = .parameter
                     End If
                 End If
                    
            End With
            
        Next
End Sub

Private Function SetSkipTime() As Single
        QueryPerformanceFrequency perfFreq
        
        SetSkipTime = (perfFreq / (xm.BPM * 2 / 5)) - 0.5 ' the 0.5 shouldn't be there, but if I leave it out everything will go too slow
        Form2.Caption = (perfFreq / (xm.BPM * 2 / 5))
        QueryPerformanceCounter TimeEnd
        TempTime = TimeEnd
End Function

Private Sub Wait()
        skiptime = SetSkipTime
        Do Until TimeEnd >= TempTime + skiptime
           QueryPerformanceCounter TimeEnd
           DoEvents
        Loop
End Sub


Public Sub SetPosition(pattern As Integer)
       CurrentPattern = pattern
       'add a function that searches for speed changes is nesecery. (direction backwards)
End Sub

Private Sub copysample(channel As Integer, index As Integer)
       ' On Error GoTo err
        dsbd.lBufferBytes = sh(index).length
        dsbd.lFlags = DSBCAPS_CTRLFREQUENCY + DSBCAPS_CTRLPAN + DSBCAPS_CTRLVOLUME + DSBCAPS_STATIC + DSBCAPS_GLOBALFOCUS + DSBCAPS_CTRLPOSITIONNOTIFY
       
        With dsbd.fxFormat
            .nFormatTag = WAVE_FORMAT_PCM
            'if sh(ins).type < 20 then
            .nChannels = 1
        
            .lSamplesPerSec = 44100
             
            .nBitsPerSample = 8
            
            .nBlockAlign = (.nChannels * .nBitsPerSample) / 8
            .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
            .lExtra = 0
            .nSize = 0
        End With
    
        Set channelBuf(channel) = ds.CreateSoundBuffer(dsbd)
    
        channelBuf(channel).WriteBuffer 0, sh(index).length, sh(index).sample(0), DSBLOCK_DEFAULT
End Sub

Private Function GetFreq(note As Byte, sample As Integer) As Single
         If xm.flags = 1 Then 'linear freq table
            period = 7680 - (note + sh(sample).rel) * 64 - sh(sample).fine
            GetFreq = 8363 * 2 ^ ((4608 - period) / 768)
         Else                 'amiga freq table.   must create the amiga lookup table.............. someday :)
            period = 10 * 12 * 16 * 4 - ((note + sh(sample).rel)) * 16 * 4 - sh(sample).fine
            GetFreq = (8363 * 2 ^ ((6 * 12 * 16 * 4 - period) / (12 * 16 * 4)))
         End If
End Function



Public Sub play2()
       StopSong = False
       
       
       
       Tick = Tick + 1
       If Tick >= xm.tempo Then
          UpdateRow Row
          
          If dochange = True Then
             CurrentPattern = CurrentPattern + 1
             Row = -1
             dochange = False
          End If
          
          Tick = 0
          Row = Row + 1
          
          
          If Row >= pattern(xm.Order(CurrentPattern)).Rows Then CurrentPattern = CurrentPattern + 1: Row = 0
          patt = xm.Order(CurrentPattern)
       Else
         'update effects
          
       End If
       
       If Not Form2.length.value = CurrentPattern Then If CurrentPattern <= Form2.length.Max Then Form2.length.value = CurrentPattern
End Sub


