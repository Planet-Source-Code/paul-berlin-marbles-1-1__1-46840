VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Keyboard Class
'------------------------------
'Created by Paul Berlin 2002-2003
'

Option Explicit

Private Type KeyData
  KeyTimer As Long
  KeyInterval As Long
End Type

Public Suspended As Boolean
Private Key(255) As KeyData
Private EndHoldTimer As Long

Private DirectInputDevice As DirectInputDevice8

Private Sub Class_Terminate()
  DirectInputDevice.Unacquire
  Set DirectInputDevice = Nothing
End Sub

Public Sub Create(ByVal hWnd As Long)
  Set DirectInputDevice = DirectInput.CreateDevice("GUID_SysKeyboard")
  DirectInputDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
  DirectInputDevice.SetCooperativeLevel hWnd, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND
  Aquire
End Sub

Public Function KeyDown(ByVal KeyCode As CONST_DIKEYFLAGS) As Boolean
  On Error Resume Next
  
  If Not Suspended And timeGetTime >= EndHoldTimer Then
    Dim KeyState As DIKEYBOARDSTATE
    Call DirectInputDevice.GetDeviceStateKeyboard(KeyState)
    If KeyState.Key(KeyCode) And Key(KeyCode).KeyTimer - timeGetTime <= 0 Then
      Key(KeyCode).KeyTimer = timeGetTime + Key(KeyCode).KeyInterval
      KeyDown = True
    End If
  End If

ErrH:
  Aquire
End Function

Public Sub SetTimer(ByVal sec As Single)
  Dim x As Long
  sec = sec * 1000
  For x = 0 To 255
    Key(x).KeyInterval = sec
  Next
End Sub

Public Sub SetTimerEx(ByVal KeyCode As CONST_DIKEYFLAGS, ByVal sec As Single)
  Key(KeyCode).KeyInterval = sec * 1000
End Sub

Public Sub HoldTimer(ByVal sec As Single)
  EndHoldTimer = timeGetTime + (sec * 1000)
End Sub

Private Sub Aquire()
  On Error GoTo ErrH
  DirectInputDevice.Acquire
  Exit Sub
ErrH:
End Sub
