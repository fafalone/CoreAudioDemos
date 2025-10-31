VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAudioSessionEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Implements IAudioSessionEvents


Private Sub IAudioSessionEvents_OnDisplayNameChanged(ByVal lpszNewDisplayName As LongPtr, EventContext As UUID)

End Sub
Private Sub IAudioSessionEvents_OnIconPathChanged(ByVal lpszNewIconPath As LongPtr, EventContext As UUID)

End Sub
Private Sub IAudioSessionEvents_OnSimpleVolumeChanged(ByVal NewVolume As Single, ByVal NewMute As BOOL, EventContext As UUID)

End Sub
Private Sub IAudioSessionEvents_OnChannelVolumeChanged(ByVal ChannelCount As Long, ByVal NewChannelVolumeArray As Longptr, ByVal ChangedChannel As Long, EventContext As UUID)

End Sub
Private Sub IAudioSessionEvents_OnGroupingParamChanged(NewGroupingParam As UUID, EventContext As UUID)

End Sub
Private Sub IAudioSessionEvents_OnStateChanged(ByVal NewState As AudioSessionState)

End Sub
Private Sub IAudioSessionEvents_OnSessionDisconnected(ByVal DisconnectReason As AudioSessionDisconnectReason)

End Sub
