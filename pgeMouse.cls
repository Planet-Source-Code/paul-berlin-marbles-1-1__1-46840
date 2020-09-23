VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Mouse Class
'---------------------------
'Created by Paul Berlin 2002-2003
'
Option Explicit

Public g_cursorx As Long
Public g_cursory As Long
Public g_cursorz As Long
Public g_Sensitivity As Single
Public button1 As Boolean
Public button2 As Boolean
Public button3 As Boolean
Public button4 As Boolean

Private Const BufferSize = 20

Private DirectInputDevice As DirectInputDevice8

Private Sub Class_Terminate()
  On Error Resume Next
  DirectInputDevice.Unacquire
  Set DirectInputDevice = Nothing
End Sub

Public Sub Create(ByRef frm As Form)
  g_cursorx = frm.ScaleWidth / 2
  g_cursory = frm.ScaleHeight / 2
  g_Sensitivity = 1.5
  
  Set DirectInputDevice = DirectInput.CreateDevice("GUID_SysMouse")
  DirectInputDevice.SetCommonDataFormat DIFORMAT_MOUSE
  DirectInputDevice.SetCooperativeLevel frm.hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE
  
  ' Set the buffer size
  Dim diProp As DIPROPLONG
  diProp.lHow = DIPH_DEVICE
  diProp.lObj = 0
  diProp.lData = BufferSize
  
  Call DirectInputDevice.SetProperty("DIPROP_BUFFERSIZE", diProp)
  
  Aquire
  
End Sub

Public Sub Aquire()
  On Error GoTo CANNOTACQUIRE
  
  DirectInputDevice.Acquire
  On Error GoTo 0
  Exit Sub
  
CANNOTACQUIRE:
  Exit Sub
End Sub

Public Sub Poll(ByRef frm As Form)
  Dim diDeviceData(1 To BufferSize) As DIDEVICEOBJECTDATA
  Dim NumItems As Long
  Dim i As Integer
  
  On Error GoTo ErrH
  DirectInputDevice.Poll
  NumItems = DirectInputDevice.GetDeviceData(diDeviceData, 0)
  On Error GoTo 0
  For i = 1 To NumItems
    Select Case diDeviceData(i).lOfs
      Case DIMOFS_X
        g_cursorx = g_cursorx + diDeviceData(i).lData * g_Sensitivity
      Case DIMOFS_Y
        g_cursory = g_cursory + diDeviceData(i).lData * g_Sensitivity
      Case DIMOFS_Z
        g_cursorz = g_cursorz + diDeviceData(i).lData
      Case DIMOFS_BUTTON0
        If diDeviceData(i).lData And &H80 Then
          button1 = True
        Else
          button1 = False
        End If
      Case DIMOFS_BUTTON1
        If diDeviceData(i).lData And &H80 Then
          button2 = True
        Else
          button2 = False
        End If
      Case DIMOFS_BUTTON2
        If diDeviceData(i).lData And &H80 Then
          button3 = True
        Else
          button3 = False
        End If
      Case DIMOFS_BUTTON3
        If diDeviceData(i).lData And &H80 Then
          button4 = True
        Else
          button4 = False
        End If
    End Select
  Next i

  If g_cursorx < 0 Then g_cursorx = 0
  If g_cursorx >= frm.ScaleWidth Then g_cursorx = frm.ScaleWidth - 1
  If g_cursory < 0 Then g_cursory = 0
  If g_cursory >= frm.ScaleHeight Then g_cursory = frm.ScaleHeight - 1

  Exit Sub
ErrH:
  If (Err.Number = DIERR_INPUTLOST) Or (Err.Number = DIERR_NOTACQUIRED) Then
    Aquire
  End If
End Sub

Public Sub UnAquire()
  DirectInputDevice.Unacquire
End Sub
