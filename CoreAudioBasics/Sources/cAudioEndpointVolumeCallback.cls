VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAudioEndpointVolumeCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAudioEndpointVolumeCallback

Private Sub IAudioEndpointVolumeCallback_OnNotify(pNotify As AUDIO_VOLUME_NOTIFICATION_DATA)
Debug.Print "Got OnNotify call; mute=" & pNotify.bMuted
Form1.AppendMsg "EndpointVolCallback: mute=" & pNotify.bMuted
End Sub
