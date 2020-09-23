Attribute VB_Name = "modMain"
Option Explicit

Public Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Public Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long

Public pEngine As New pgeMain
Public pKeyboard As New pgeKeyboard
Public pTextures As New pgeTexturePool
Public pSound As New pgeSound
Public pMouse As New pgeMouse

Public FontArial As New pgeFont
Public MainFont As New pgeFont
Public LedFont As New pgeFont

Type tSettings
  SfxVolume As Byte
  MusicVolume As Byte
  MouseSpeed As Single
End Type

Type tHighScore
  lScore As Long
  sName As String
End Type

Type tPlayer
  lScore As Long
  lTime As Long
  lDisplayTime As Long
  lBombs As Long
End Type

Type tGrd
  lType As Integer 'Marble type. 0 = empty
  lY As Single 'Y coordinate of marble
  lX As Single 'X coordinate of marble
  lDead As Long
  lFlag(3) As Long 'Used when dying
  bSpecial As Byte 'Special number. 0 = no special
End Type

Public CurrentMusic As Integer
Public Const MaxMusic As Integer = 14

Public LatestHigh As Long
Public bFps As Boolean 'Show fps on/off
Public lGrid(7, 8) As tGrd 'Playing field grid
Public Player As tPlayer 'Player status
Public Settings As tSettings 'Program settings
Public High(9) As tHighScore 'Highscore

Public Function FileExist(ByVal FileName As String) As Boolean
  FileExist = Not (Dir(FileName) = "")
End Function
