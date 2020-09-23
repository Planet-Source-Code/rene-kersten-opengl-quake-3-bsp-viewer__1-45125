Attribute VB_Name = "Sound"

Option Explicit

'DirectX8: The master controlling object
Private DX As DirectX8
'DirectSound8: Looks after all of the sound playback interfaces
Private DS As DirectSound8
'DSenum: object for enumeration
Private DSEnum As DirectSoundEnum8
'bLoaded: Status flag insuring that we have actually loaded our sound properly...
Private bLoaded As Boolean

Private DMPerformance As DirectMusicPerformance8
Private DMLoader As DirectMusicLoader8

Private numMidis As Long
Private Midis() As DirectMusicSegment8

Private numSounds As Long
Private sounds() As DirectSoundSecondaryBuffer8

Public Sub initSound()
    Dim params As DMUS_AUDIOPARAMS
    Dim i As Long
    
    If bLoaded Then
        For i = 0 To numSounds
            Set sounds(i) = Nothing
        Next i
        Set DS = Nothing
        Set DX = Nothing
    End If
    bLoaded = True
    
    Set DX = New DirectX8
    Set DSEnum = DX.GetDSEnum
    Set DS = DX.DirectSoundCreate(DSEnum.GetGuid(1))
    
    Set DMPerformance = DX.DirectMusicPerformanceCreate
    Set DMLoader = DX.DirectMusicLoaderCreate
    
    DS.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL
    
    DMPerformance.InitAudio Form1.hWnd, DMUS_AUDIOF_ALL, params, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128
    DMPerformance.SetMasterAutoDownload True
    
    numSounds = -1
    numMidis = -1
    
    DMLoader.SetSearchDirectory App.path & "\"
End Sub

Public Sub quitSound()
    Dim i
    If bLoaded Then
        For i = 0 To numSounds
            Set sounds(i) = Nothing
        Next i
        Set DS = Nothing
        
        For i = 0 To numMidis
            Set Midis(i) = Nothing
        Next i
        Set DMLoader = Nothing
        DMPerformance.CloseDown
        Set DMPerformance = Nothing
        
        Set DX = Nothing
    End If
    bLoaded = False
End Sub

Public Function precacheSound(filename As String) As Long
    Dim tmpDesc As DSBUFFERDESC
    tmpDesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY
    
    numSounds = numSounds + 1
    ReDim Preserve sounds(numSounds)
    precacheSound = numSounds
    
    Set sounds(numSounds) = DS.CreateSoundBufferFromFile(filename, tmpDesc)
End Function

Public Function precacheMidi(filename As String) As Long
    numMidis = numMidis + 1
    ReDim Preserve Midis(numMidis)
    precacheMidi = numMidis
    
    Set Midis(numMidis) = DMLoader.LoadSegment(filename)
    Midis(numMidis).SetStandardMidiFile
End Function

Public Sub playSound(soundNum As Long, shouldLoop As Boolean, Optional position As Long = 0)
    If Not position = 0 Then sounds(soundNum).SetCurrentPosition position
    If shouldLoop Then
        sounds(soundNum).Play DSBPLAY_LOOPING
    Else
        sounds(soundNum).Play DSBPLAY_DEFAULT
    End If
End Sub

Public Sub playMidi(midiNum As Long, shouldLoop As Boolean)
    If shouldLoop Then
        Midis(midiNum).SetRepeats -1
    Else: Midis(midiNum).SetRepeats 0
    End If
    
    DMPerformance.PlaySegmentEx Midis(midiNum), DMUS_SEGF_DEFAULT, 0
End Sub

Public Sub stopSound(soundNum As Long)
    sounds(soundNum).Stop
    sounds(soundNum).SetCurrentPosition 0
End Sub

Public Sub stopMidi(midiNum As Long)
    DMPerformance.StopEx Midis(midiNum), 0, DMUS_SEGF_DEFAULT
End Sub
