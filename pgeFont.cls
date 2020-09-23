VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Font Class
'--------------------------
'Created by Paul Berlin 2002-2003
'
'
Option Explicit

Private g_Font As D3DXFont
Private Info As LOGFONT

Public Function GetInfo() As LOGFONT
  GetInfo = Info
End Function

Public Sub Create(fnt As IFont)
  Set g_Font = Direct3DX.CreateFont(Direct3DDevice, fnt.hFont)
  g_Font.GetLogFont Info ' return info about the font
End Sub

Public Sub DrawText(ByVal sText As String, rt As RECT, lColor As Long, Optional format As CONST_DTFLAGS = 0)
  Direct3DX.DrawText g_Font, lColor, sText, rt, format
End Sub

