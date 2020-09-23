VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SoundCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'DirectSound sections
Private DX As New DirectX7
Private DS As DirectSound
Private DSBuffer As DirectSoundBuffer

'DirectMusic sections
Private DMPerformance As DirectMusicPerformance
Private DMSegment As DirectMusicSegment
Private DMSegmentState As DirectMusicSegmentState
Private DMLoader As DirectMusicLoader

Public aHwnd As Long


Public Sub CloseDM()
    Set DMSegment = Nothing
    Set DMSegmentState = Nothing
    
    If Not (DMPerformance Is Nothing) Then DMPerformance.CloseDown
    Set DMPerformance = Nothing
    Set DMLoader = Nothing
End Sub

Public Sub CloseDS()
    Set DSBuffer = Nothing
    Set DS = Nothing
    Set DX = Nothing
End Sub

Public Function InitDM() As Boolean
    On Error GoTo failed
    Set DMLoader = DX.DirectMusicLoaderCreate()
    Set DMPerformance = DX.DirectMusicPerformanceCreate()
    
    DMPerformance.Init Nothing, aHwnd
    DMPerformance.SetPort -1, 80
    DMPerformance.SetMasterAutoDownload (True)
    
    InitDM = True
    Exit Function
failed:
    MsgBox "Failed to Init DirectMusic", vbOKOnly, "DirectMusic"
    InitDM = False
End Function

Public Function InitDS() As Boolean
    On Error GoTo failed
    Set DS = DX.DirectSoundCreate("")
    DS.SetCooperativeLevel aHwnd, DSSCL_PRIORITY
    
    InitDS = True
    Exit Function
failed:
    MsgBox "Failed To Init DirectSound", vbOKOnly, "DirectSound"
    InitDS = False
End Function

Public Function LoadMusic(StrFileName As String) As Boolean
    On Error GoTo failed
    
    Set DMSegment = DMLoader.LoadSegment(StrFileName)
    DMSegment.SetTempoEnable (True)
    DMSegment.SetStandardMidiFile
    DMSegment.SetRepeats 10
    
    LoadMusic = True
    Exit Function
failed:
    'MsgBox "Error Loading Music", vbOKOnly, "Loading Music"
    LoadMusic = False
End Function

Public Function LoadWave(StrFileName As String) As Boolean
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX

    On Error GoTo failed
    
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    Set DSBuffer = DS.CreateSoundBufferFromFile(StrFileName, bufferDesc, waveFormat)
    
    LoadWave = True
    Exit Function
failed:
    MsgBox "Unable to Load the Wave File", vbOKOnly, "Load Wave"
    LoadWave = False
End Function

Public Sub PlayMusic()
    Set DMSegmentState = DMPerformance.PlaySegment(DMSegment, 0, 0)
End Sub

Public Sub PlayWave()
    DSBuffer.Play DSBPLAY_DEFAULT
End Sub

Public Sub StopMusic()
    Call DMPerformance.Stop(Nothing, Nothing, 0, 0)
End Sub

