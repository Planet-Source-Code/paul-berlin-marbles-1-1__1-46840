VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Timer Class
'---------------------------
'Created by Paul Berlin 2002-2003
'
Private lStartTime As Long
Private lStopTime As Long

Option Explicit

Public Sub StartTime()
  lStartTime = timeGetTime
End Sub

Public Function GetTime() As Single
  GetTime = (timeGetTime - lStartTime) / 1000
End Function

Public Function GetTimeMs() As Long
  GetTimeMs = timeGetTime - lStartTime
End Function

Public Sub SetStop(ByVal sSec As Single)
  lStopTime = timeGetTime + (sSec * 1000)
  lStartTime = timeGetTime
End Sub

Public Function HasStopped() As Boolean
  If timeGetTime >= lStopTime Then HasStopped = True
End Function

Public Sub Hold(ByVal sSec As Single)
  Dim lStart As Long
  
  lStart = timeGetTime
  Do
    DoEvents
  Loop Until timeGetTime - lStart >= (sSec * 1000)
End Sub
