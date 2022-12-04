Attribute VB_Name = "loadxm"
Option Explicit

Type xmheader
     ex                        As String * 17        ' stores the string Extended Module
     name                      As String * 20        ' stores the name of the song
     unknown                   As Byte               ' dunno what this is or does, thou it seems as it always is 26 ???????????
     trackername               As String * 20        ' usually stores the name of the tracker used to create the xm file
     version                   As Integer            ' self-explaining
     size                      As Long               ' stores the size of the header
     length                    As Integer            ' the length of the song (in patterns).
     restartpos                As Integer            ' restart position
     channels                  As Integer            ' number of channels (2,4,6,8,10,...,32)
     patterns                  As Integer            ' number of patterns (max 256)
     instruments               As Integer            ' number of instruments (max 128)
     flags                     As Integer            ' flags 0 = amiga freq table, 1 = linear freq table
     tempo                     As Integer            ' tempo
     BPM                       As Integer            ' BPM (Beats Per Minute)
     Order(255)                As Byte               ' pattern order (0,1,5,2,34....)
End Type

Type noteprop
     note                      As Byte               ' stores notes from 0-96. 97 is note off
     instrument                As Byte               ' stores the instrument used
     volume                    As Byte               ' volume
     effect                    As Integer            ' effect type      (arppegio,porta....)
     parameter                 As Byte               ' effect parameter (value)
End Type

Type instrumentheader
     size                      As Long               ' stores the size of the header
     name                      As String * 22        ' hmmm, name
     type                      As Byte               ' should always be 0
     samples                   As Integer            ' number of samples in instrument
End Type

Type instrumentheader2
     size                      As Long               ' stores the size of the header
     sample(95)                As Byte               ' table refering to samples
     volenv(23)                As Integer            ' volume envelopes, stored as x,y .... x,y
     panenv(23)                As Integer            ' panning envelopes, stored as x,y .... x,y
     volpoints                 As Byte               ' number of volume points stored in instrument
     panpoints                 As Byte               ' number of panning points stored in instrument
     volsustain                As Byte               ' volume sustain point
     volloopstart              As Byte               ' volume loop start point
     volloopend                As Byte               ' volume loop end point
     pansustain                As Byte               ' panning sustain point
     panloopstart              As Byte               ' panning loop start point
     panloopend                As Byte               ' panning loop end point
     voltype                   As Byte               ' volume type: 0 = on ; 1 = sustain ; 2 = loop
     pantype                   As Byte               ' panning type: 0 = on ; 1 = sustain ; 2 = loop
     vibtype                   As Byte               ' vibrato type
     vibsweep                  As Byte               ' vibrato sweep
     vibdepth                  As Byte               ' vibrato depth
     vibrate                   As Byte               ' vibrato rate
     volfade                   As Integer            ' volume fadeout
End Type

Type sampleheader
     length                    As Long               ' sample length
     loopstart                 As Long               ' loop start
     loopend                   As Long               ' loop end
     volume                    As Byte               ' volume
     fine                      As Integer            ' finetune  signed byte
     type                      As Byte               ' type 0 = no loop ; 1 = forward loop ; 2 = ping-pong loop ; 16 = 16-bit sampledata
     panning                   As Byte               ' panning
     rel                       As Integer            ' signed byte
     reserved                  As Byte               ' reserved
     name                      As String * 22        ' name
     sample()                  As Byte               ' holds 16-bit sample data (all 8-bit samples are converted into 16-bit)
     bidi()                    As Integer            ' bidi loop (BiDirectional), also known as ping-pong loop
End Type

Type patt
     size                      As Long               ' always 9
     type                      As Byte               ' always 0 (don't think that xm files contain any other types of patterns)
     Rows                      As Integer            ' number of rows in this pattern
     patternsize               As Integer            ' the size of the pattern
     pattern()                 As noteprop           ' stored in pattern(row,col).?
End Type

Public xm                      As xmheader           ' declare the XM header
Public ih()                    As instrumentheader   ' declare the instrument header nr 1
Public ih2()                   As instrumentheader2  ' declare the instrument header nr 2
Public sh()                    As sampleheader       ' declare the sample header
Public pattern()               As patt               ' pattern array pattern(pattern_number).pattern(row,col).?

Public commentary              As String             ' hmm commentary what can that be?

Dim ffile                      As Integer            ' for use with freefile

Private Const WhereError       As String = "LoadXm:" ' string that tells me where the error is

Dim temp()                     As Byte               ' temporary space for loading samples

Private Function OpenByte() As Byte
        Dim byt As Byte
        Get #ffile, , byt
        OpenByte = byt
End Function

Private Sub OpenForByt(arr() As Byte, amount As Integer)
        Dim i As Integer
    
        ReDim arr(amount)
        
        For i = 1 To amount
            arr(i) = OpenByte
        Next
End Sub

Private Function OpenLong() As Long
        Dim lng As Long
        Get #ffile, , lng
        OpenLong = lng
End Function

Private Function OpenInteger() As Integer
        Dim ints As Integer
        Get #ffile, , ints
        OpenInteger = ints
End Function

'######################################################################################################################################################
'# this function is meant for converting unsigned bytes (bytes that has been stored as signed in other programs)
'# to signed bytes (integers, since vb doesn't handle signed bytes)
'######################################################################################################################################################
Private Function sign(value As Integer) As Integer
        If value > 127 Then
           sign = (value Mod 128) - 128
        Else
           sign = value
        End If
End Function

'######################################################################################################################################################
'# this function is the same as above only revesed
'######################################################################################################################################################
Private Function unsign(value As Long) As Byte
        unsign = (value + 128) And 255
End Function

'######################################################################################################################################################
'# If first < 128 then the note isn't packed and is to be read as usuall.
'# If first >= 128 then the note is packed
'#-----------------------------------------------------------------------------------------------------------------------------------------------------
'# by checking the last   bit of "first"   we know if there is any packing scheme    (first and 128)
'# By checking the first  bit of "first"   we know if there is any note stored       (first and 1)
'# By checking the second bit of "first"   we know if there is any instrument stored (first and 2)
'# By checking the third  bit of "first"   we know if there is any volume stored     (first and 4)
'# By checking the fourth bit of "first"   we know if there is any effect stored     (first and 8)
'# By checking the fifth  bit of "first"   we know if there is any parameter stored  (first and 16)
'#-----------------------------------------------------------------------------------------------------------------------------------------------------
'# if first = 128 then you don't even need to check any of the other values, since they will all be = 0
'######################################################################################################################################################
Private Sub Xm_readnote(n As noteprop)
        On Error GoTo newerr
        Dim first As Byte
        
        first = OpenByte
         
        If first = 128 Then Exit Sub ' shouldn't be necesery, but should provide a little boost in loading time
        
        If (first And 128) Then
            If (first And 1) Then n.note = OpenByte
            If (first And 2) Then n.instrument = OpenByte
            If (first And 4) Then n.volume = OpenByte
            If (first And 8) Then n.effect = OpenByte
            If (first And 16) Then n.parameter = OpenByte
        Else
            n.note = first
            n.instrument = OpenByte
            n.volume = OpenByte
            n.effect = OpenByte
            n.parameter = OpenByte
        End If
        Exit Sub
newerr:
        addlog "###################################"
        addlog WhereError & "error reading note"
        addlog "###################################"
End Sub

'######################################################################################################################################################
'# this sub loads up the xm file.
'# if "patternsize" = 0 then the pattern is empty and therefore not stored. This means that it has to be created manually.
'# DO Not try to load a pattern that doesn't exist.
'# the Xm_readnote function is described above this sub
'######################################################################################################################################################
Public Sub load_file(FileName As String)
       Dim i          As Integer
       Dim x          As Byte
       Dim y          As Integer ' I made this an integer instead of a byte just to aviod overflow errors
 
       ffile = FreeFile
       
       warning16 = False
       
       addlog "Loading song: " & FileName
       
       If Len(FileName) = 0 Then Exit Sub
       
       Open FileName For Binary As #ffile
            Get #ffile, , xm
            
            ReDim pattern(xm.patterns)
            
            For i = 0 To (xm.patterns - 1)
                With pattern(i)
                     Get #ffile, , .size
                     Get #ffile, , .type
                     Get #ffile, , .Rows
                     Get #ffile, , .patternsize
                     
                     If .patternsize = 0 Then
                        ReDim Preserve pattern(i).pattern(64, xm.channels)
                        .Rows = 64
                        GoTo EmptyPattern
                     End If
                
                     ReDim Preserve pattern(i).pattern(.Rows, xm.channels)
             
                     On Error GoTo err ' if an error should occur then the file is probably damaged
             
                     For y = 1 To .Rows
                     For x = 1 To xm.channels
                         Xm_readnote pattern(i).pattern(y - 1, x - 1)
                     Next
                     Next
                     
                End With
EmptyPattern:
            Next
            On Error GoTo 0

             
            Call loadinstruments
            Call LoadCommentary
        
            ReDim channelBuf(xm.channels)
            ReDim channelsort(xm.channels)
            
            Dim location As Byte

            For i = 255 To 1 Step -1
                If Not xm.Order(i) = 0 And Not xm.Order(i) = 255 Then
                   location = i
                   Exit For
                End If
            Next
            
            If xm.patterns < location Then xm.patterns = location 'some songs doesn't seem to store the correct length, hence this function.

       Close #ffile
       Exit Sub
    
err:
       Close #ffile
       addlog "###################################"
       addlog WhereError & "error loading patterns" & "  " & err.Description & "  " & err.Number
       addlog "###################################"
End Sub

Private Sub LoadCommentary()
        commentary = ""
            
        Dim name  As String * 1
            
        Do
          Get #ffile, , name
          commentary = commentary & name
        Loop Until EOF(ffile)
End Sub

'######################################################################################################################################################
'# calculating the reserved data is necesery, since it varies depending on wich file version it is.
'# as you probalby notice I chosed to open every data segment manually instead of using "get #ffile, , ih"
'# this is becouse I did some debuging of the input. (It has been removed since i got everything to load as it should)
'# besides I think you got more control and can check very easy if something doesn't load or behave as it should
'#
'# Note this is something that is unclear to me
'# should 16-bit data be loaded into a integer, due to the fact that an integer is an 16-bit varible?
'# I did try this and it produced a really good sound in overall, but the samples somehow was chopped of in the middle?????
'# Or could you load up 16-bit data in bytes aswell
'######################################################################################################################################################
Private Sub loadinstruments()
        Dim reservedsize As Long
        Dim i            As Integer
        Dim reserved()   As Byte
        Dim tmp          As Integer
        Dim Tmp2         As Integer
        Dim n            As Integer
        Dim sampcount    As Integer
        
        ReDim ih(xm.instruments)
        ReDim ih2(xm.instruments)
        ReDim sh(0)
        
        For i = 1 To xm.instruments
        
        reservedsize = Seek(ffile)
        
        With ih(i)
            .size = OpenLong
             Get #ffile, , .name
            .type = OpenByte
            .samples = OpenInteger
        End With
        
        reservedsize = reservedsize + ih(i).size
        
        If ih(i).size > 29 Then
           If ih(i).samples > 0 And ih(i).samples < 96 Then
              With ih2(i)
                  .size = OpenLong
                   Get #ffile, , .sample
                   Get #ffile, , .volenv
                   Get #ffile, , .panenv
                  .volpoints = OpenByte
                  .panpoints = OpenByte
                  .volsustain = OpenByte
                  .volloopstart = OpenByte
                  .volloopend = OpenByte
                  .pansustain = OpenByte
                  .panloopstart = OpenByte
                  .panloopend = OpenByte
                  .voltype = OpenByte
                  .pantype = OpenByte
                  .vibtype = OpenByte
                  .vibsweep = OpenByte
                  .vibdepth = OpenByte
                  .vibrate = OpenByte
                  .volfade = OpenInteger
                   
                   OpenForByt reserved, reservedsize - Seek(ffile)
                   
                   sampcount = sampcount + ih(i).samples
              
                   For n = 0 To UBound(ih2(i).sample)
                       If Not ih2(i).sample(n) + sampcount - ih(i).samples + 1 > 255 Then ih2(i).sample(n) = ih2(i).sample(n) + sampcount - ih(i).samples + 1
                   Next
              End With
           Else
              
              OpenForByt reserved, reservedsize - Seek(ffile)
              
           End If
          
           For n = 1 To ih(i).samples
               tmp = tmp + 1
               ReDim Preserve sh(tmp)
               With sh(tmp)
                   .length = OpenLong
                   .loopstart = OpenLong
                   .loopend = OpenLong
                   .volume = OpenByte
                   .fine = sign(OpenByte) / 2
                   .type = OpenByte
                   .panning = OpenByte
                   .rel = sign(OpenByte)
                   .reserved = OpenByte
                    Get #ffile, , .name
               End With
           Next
              
           For n = 1 To ih(i).samples
               Tmp2 = Tmp2 + 1
               With sh(Tmp2)
                    If .length > 0 And .length < 6000000 Then
                        ReDim temp(.length - 1)
                        Get #ffile, , temp()
                        ConvertSamples Tmp2
                    End If
               End With
           Next
           End If
        Next
        
err:

Exit Sub
End Sub

'######################################################################################################################################################
'# my intentions with this sub is to create 16-bit stereo samples of all samples that are loaded into memmory.
'# this makes it alot easier to handle the play function. since I don't have to worry about what type of format the
'# incomming signal is. there's only one problem with this. I unfortunatly don't know how to do it (yet).
'# 8-bit samples are working though.
'#
'# All samples are stored as(s) signed delta values, I need to convert them into unsigned values
'# all I need to do now is to convert them into 16-bit stereo samples
'# the reason I want it to be "16-bit, stereo" is that a update of the format has been done making it support stereo.
'# otherwise I could have settled with "16-bit, mono"
'######################################################################################################################################################
Private Sub ConvertSamples(index As Integer)
        Dim i As Long
        Dim old As Long
        Dim Anew As Long
        Dim t As Long
        On Error Resume Next
        '32767
        old = 0
        Anew = 0
        
        ReDim sh(index).sample(sh(index).length)
        
        For i = 0 To sh(index).length - 1
            Anew = temp(i) + old
            sh(index).sample(i) = unsign(CLng(Anew))
            old = Anew
        Next
        

        If sh(index).type > 4 Then warning16 = True
End Sub
