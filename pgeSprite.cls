VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Sprite Class
'----------------------------
'Created by Paul Berlin 2002-2003
'
Option Explicit

' SPRITE CHANGEABLE PARAMETERS
Private SprCA As Single 'Current Alpha color component
Private SprCR As Single 'Current Red color component
Private SprCG As Single 'Current Green color component
Private SprCB As Single 'Current Blue color component
Private SprPos As D3DVECTOR2 'Current Position of sprite
Private SprRot As Single 'Current rotation of sprite
Private SprRotC As D3DVECTOR2 'The center of rotation on sprite
Private SprScale As D3DVECTOR2 'Current scale of sprite
Public AniCurrentFrame As Long 'Current frame number
Public Active As Boolean 'Is sprite active? Only active sprites are drawn

'FIXED ANIMATION INFO
Private AniTexture As String 'Name of texture
Private AniWidth As Long 'Width of each sprite
Private AniHeight As Long 'Height of each sprite
Private AniCols As Long '# Columns of sprites
Private AniRows As Long '# Rows of sprites
Private AniStartFrame As Long 'First frame of animation
Private AniEndFrame As Long 'Last frame of animation
Private AniNumFrames As Long '# of frames

'ANIMATION FLAGS
Public bLoop As Boolean 'LOOPOFF flag
Public bAnimStop As Boolean 'DISABLEWHENDONE flag
Public bCenterScale As Boolean 'CENTERSCALE flag
Public bReverse As Boolean 'Reverse animation
Public bAnimated As Boolean
Public bYMirror As Boolean
Public bXMirror As Boolean

'AUTOMATED MOVEMENTS
Private ARotAngle As Single 'Angle to rotate sprite each interval
Private ARotInterval As Long 'Time in ms between each rotation

Private AScale As D3DVECTOR2 'Scale goal
Private AScaleDelta As D3DVECTOR2 'scale change per 10 ms

Private AMove As D3DVECTOR2 'Delta each interval
Private AMoveInterval As Long 'Time in ms between each movement

Private APathGoal As D3DVECTOR2 'Destination coordinates
Private APathDelta As D3DVECTOR2 'Delta each interval

Private AColRDelta As Single 'Red color delta each interval
Private AColGDelta As Single 'Green color delta each interval
Private AColBDelta As Single 'Blue color delta each interval
Private AColADelta As Single 'Alpha color delta each interval
Private AColRGoal As Byte 'Red color goal
Private AColGGoal As Byte 'Green color goal
Private AColBGoal As Byte 'Blue color goal
Private AColAGoal As Byte 'Alpha color goal
Private AColDisable As Boolean 'If true, disable sprite when reached fade stop

'FRAME INFO
Private Type tValue
  lValue As Long
  bDelta As Boolean 'If true lValue is an delta value
End Type

Private Type tFrame
  lDelay As Long 'Frame delay in milliseconds
  lR As tValue 'red value
  lG As tValue 'green value
  lB As tValue 'blue value
  lA As tValue 'alpha value
End Type
  
Private Frame() As tFrame
  
'RENDER VALUES
Private DestRect As RECT
Private SourceRect As RECT

Friend Sub CreateFromTexture(ByVal sTexture As String, Optional ByVal StartFrame As Long = 1, Optional ByVal EndFrame As Long = 1, Optional ByVal FrameInterval As Long = 100)
    AniTexture = sTexture
    
    AniWidth = TPool(AniTexture).SpriteWidth
    AniHeight = TPool(AniTexture).SpriteHeight
    AniCols = TPool(AniTexture).Columns
    AniRows = TPool(AniTexture).Rows
    AniStartFrame = StartFrame
    AniEndFrame = EndFrame
    AniNumFrames = EndFrame - StartFrame + 1
    AniCurrentFrame = StartFrame
    bLoop = True
    
    If AniNumFrames > 1 Then bAnimated = True
    
    ReDim Frame(AniNumFrames - 1)
    For StartFrame = 0 To AniNumFrames - 1
      Frame(StartFrame).lDelay = FrameInterval
      Frame(StartFrame).lA.bDelta = True
      Frame(StartFrame).lR.bDelta = True
      Frame(StartFrame).lG.bDelta = True
      Frame(StartFrame).lB.bDelta = True
    Next
    
    bCenterScale = True
    
    SprRotC.x = AniWidth / 2: SprRotC.y = AniHeight / 2
    SprPos.x = SprRotC.x: SprPos.y = SprRotC.y
    SprRot = 0
    SprCA = 255: SprCR = 255: SprCG = 255: SprCB = 255
    SprScale.x = 1: SprScale.y = 1
    
    Active = True
End Sub

Public Sub Render()
  
  If Active Then
        
    Dim iCol As Long, iRow As Long
    Dim sVec As D3DVECTOR2
    Dim rCent As D3DVECTOR2
    Dim offX As Single, offY As Single
    Static fTime As Long 'frame Timer
    Static mTime As Long 'move Timer
    Static rTime As Long 'roatate Timer
    Static cTime As Long 'color Timer
    Static sTime As Long 'scale Timer
    Static pTime As Long 'path Timer
    Dim tSng As Single, tLng As Long
    Dim iFrame As Long
    
    'make sure timers are on
    If mTime = 0 Then mTime = timeGetTime
    If fTime = 0 Then fTime = timeGetTime
    If rTime = 0 Then rTime = timeGetTime
        
    '#### AUTOMOVE
    If timeGetTime - mTime >= AMoveInterval And AMoveInterval > 0 Then
      If FrameSkip Then 'Get missed movements & add them
        tSng = (timeGetTime - mTime) / AMoveInterval
      Else
        tSng = 1
      End If
      mTime = timeGetTime
      
      With SprPos
        .x = .x + (AMove.x * tSng)
        .y = .y + (AMove.y * tSng)
      End With
    Else
      If AMoveInterval = 0 Then mTime = timeGetTime
    End If
    
    '#### AUTOPATH
    If (APathDelta.x <> 0 Or APathDelta.y <> 0) Then
      If FrameSkip And pTime <> 0 Then 'Get missed moves & add them
        tSng = (timeGetTime - pTime) / 10
      Else
        tSng = 1
      End If
      pTime = timeGetTime
      
      If APathDelta.x <> 0 Then
        With SprPos
          .x = .x + (APathDelta.x * tSng)
          If APathDelta.x < 0 Then
            If .x <= APathGoal.x Then
              .x = APathGoal.x
              APathDelta.x = 0
            End If
          Else
            If .x >= APathGoal.x Then
              .x = APathGoal.x
              APathDelta.x = 0
            End If
          End If
        End With
      End If
      If APathDelta.y <> 0 Then
        With SprPos
          .y = .y + (APathDelta.y * tSng)
          If APathDelta.y < 0 Then
            If .y <= APathGoal.y Then
              .y = APathGoal.y
              APathDelta.y = 0
            End If
          Else
            If .y >= APathGoal.y Then
              .y = APathGoal.y
              APathDelta.y = 0
            End If
          End If
        End With
      End If
      
    Else
      pTime = 0
    End If
    '####
    
    'DESTINIATION RECTANGLE, USED FOR INTERSECTING SPRITES
    With DestRect
      If bCenterScale Then
        .Left = SprPos.x - AniWidth / 2 * SprScale.x
        .Top = SprPos.y - AniHeight / 2 * SprScale.y
        .Right = SprPos.x + AniWidth / 2 * SprScale.x
        .bottom = SprPos.y + AniHeight / 2 * SprScale.y
      Else
        .Left = SprPos.x
        .Top = SprPos.y
        .Right = SprPos.x + AniWidth * SprScale.x
        .bottom = SprPos.y + AniHeight * SprScale.y
      End If
    End With
    
    'If the sprite isnt on screen, skip it!
    'If (DestRect.Right > Target.Left - ScrollX And DestRect.Left < Target.Right - ScrollX) Or (DestRect.Left = 0 And DestRect.Right = 0) Then
    '  If (DestRect.bottom > Target.Top - ScrollY And DestRect.Top < Target.bottom - ScrollY) Or (DestRect.Top = 0 And DestRect.bottom = 0) Then
    
        '#### AUTOSCALING
        If (AScaleDelta.x <> 0 Or AScaleDelta.y <> 0) Then
          If FrameSkip And sTime <> 0 Then 'Get missed scales & add them
            tSng = (timeGetTime - sTime) / 10
          Else
            tSng = 1
          End If
          sTime = timeGetTime
          
          'Debug.Print tSng
          
          If AScaleDelta.x <> 0 Then
            With SprScale
              .x = .x + (AScaleDelta.x * tSng)
              If AScaleDelta.x < 0 Then
                If .x <= AScale.x Then
                  .x = AScale.x
                  AScaleDelta.x = 0
                End If
              Else
                If .x >= AScale.x Then
                  .x = AScale.x
                  AScaleDelta.x = 0
                End If
              End If
            End With
          End If
          If AScaleDelta.y <> 0 Then
            With SprScale
              .y = .y + (AScaleDelta.y * tSng)
              If AScaleDelta.y < 0 Then
                If .y <= AScale.y Then
                  .y = AScale.y
                  AScaleDelta.y = 0
                End If
              Else
                If .y >= AScale.y Then
                  .y = AScale.y
                  AScaleDelta.y = 0
                End If
              End If
            End With
          End If
          
        Else
          sTime = 0
        End If
        '####
        
        '#### MIRROR SPRITE
        sVec = SprScale
        If bXMirror Then
          sVec.x = -sVec.x
          offX = AniWidth
        End If
        If bYMirror Then
          sVec.y = -sVec.y
          offY = AniHeight
        End If

        '#### AUTOROTATE
        If timeGetTime - rTime >= ARotInterval And ARotInterval > 0 Then
          If FrameSkip Then 'Get missed rotations & add them
            tSng = (timeGetTime - rTime) / ARotInterval
          Else
            tSng = 1
          End If
          rTime = timeGetTime
            
          SprRot = SprRot + (ARotAngle * tSng)
        Else
          If ARotInterval = 0 Then rTime = timeGetTime
        End If
        
        '#### ANIMATED SPRITES
        If bAnimated Then
       
          'increase, decrease frame if specified
          iFrame = AniCurrentFrame - AniStartFrame
          If timeGetTime - fTime >= Frame(iFrame).lDelay And Frame(iFrame).lDelay > 0 And bAnimated Then
                    
            If FrameSkip Then 'Skip frame if double time or more has elapsed
              tLng = (timeGetTime - fTime) / Frame(iFrame).lDelay
            Else
              tLng = 1
            End If
            fTime = timeGetTime
                    
            If bReverse Then
              AniCurrentFrame = AniCurrentFrame - (1 * tLng)
            Else
              AniCurrentFrame = AniCurrentFrame + (1 * tLng)
            End If
                    
            'bounds of animation:
            If AniCurrentFrame > AniEndFrame Then
              If bAnimStop Then
                Active = False
                Exit Sub
              End If
              If bLoop Then
                Do 'Wrap around frame number if it is higher than max frame
                  AniCurrentFrame = AniCurrentFrame - AniEndFrame + AniStartFrame - 1
                Loop Until AniCurrentFrame <= AniEndFrame
              Else
                AniCurrentFrame = AniEndFrame
              End If
            End If
            If AniCurrentFrame < AniStartFrame Then
              If bLoop Then
                Do 'Wrap around frame number if it is lower than min frame
                  AniCurrentFrame = AniCurrentFrame - AniStartFrame + AniEndFrame + 1
                Loop Until AniCurrentFrame >= AniStartFrame
              Else
                AniCurrentFrame = AniStartFrame
              End If
            End If
            
          End If
                    
          'Setup this frames color
          iFrame = AniCurrentFrame - AniStartFrame
          With Frame(iFrame)
            If .lA.bDelta Then
              SprCA = SprCA + .lA.lValue
            Else
              SprCA = .lA.lValue
            End If
            If .lR.bDelta Then
              SprCR = SprCR + .lR.lValue
            Else
              SprCR = .lR.lValue
            End If
            If .lG.bDelta Then
              SprCG = SprCG + .lG.lValue
            Else
              SprCG = .lG.lValue
            End If
            If .lB.bDelta Then
              SprCB = SprCB + .lB.lValue
            Else
              SprCB = .lB.lValue
            End If
          End With
                    
        End If
                
        '#### AUTOFADE
        If (AColRDelta <> 0 Or AColGDelta <> 0 Or AColBDelta <> 0 Or AColADelta <> 0) Then
          If FrameSkip And cTime <> 0 Then  'Get missed color changes & add them
            tSng = (timeGetTime - cTime) / 10
          Else
            tSng = 1
          End If
          cTime = timeGetTime
          
          If AColRDelta <> 0 Then
            SprCR = SprCR + (AColRDelta * tSng)
            If AColRDelta < 0 Then
              If SprCR <= AColRGoal Then
                SprCR = AColRGoal
                AColRDelta = 0
              End If
            Else
              If SprCR >= AColRGoal Then
                SprCR = AColRGoal
                AColRDelta = 0
              End If
            End If
          End If
          If AColGDelta <> 0 Then
            SprCG = SprCG + (AColGDelta * tSng)
            If AColGDelta < 0 Then
              If SprCG <= AColGGoal Then
                SprCG = AColGGoal
                AColGDelta = 0
              End If
            Else
              If SprCG >= AColGGoal Then
                SprCG = AColGGoal
                AColGDelta = 0
              End If
            End If
          End If
          If AColBDelta <> 0 Then
            SprCB = SprCB + (AColBDelta * tSng)
            If AColBDelta < 0 Then
              If SprCB <= AColBGoal Then
                SprCB = AColBGoal
                AColBDelta = 0
              End If
            Else
              If SprCB >= AColBGoal Then
                SprCB = AColBGoal
                AColBDelta = 0
              End If
            End If
          End If
          If AColADelta <> 0 Then
            SprCA = SprCA + (AColADelta * tSng)
            If AColADelta < 0 Then
              If SprCA <= AColAGoal Then
                SprCA = AColAGoal
                AColADelta = 0
              End If
            Else
              If SprCA >= AColAGoal Then
                SprCA = AColAGoal
                AColADelta = 0
              End If
            End If
          End If
          
          If SprCA = AColAGoal And SprCR = AColRGoal And SprCG = AColGGoal And SprCB = AColBGoal And AColDisable Then
            AColDisable = False
            Active = False
          End If
        Else
          cTime = 0
        End If
            
        'Calculate current row & col from frame number
        If (AniCurrentFrame - 1) / AniCols >= 1 Then
          iRow = Int((AniCurrentFrame - 1) / AniCols)
        Else
          iRow = 0
        End If
        iCol = (AniCurrentFrame - 1) - (iRow * AniCols)
        
        'Calculate scaled rotation center
        rCent.x = SprRotC.x * SprScale.x
        rCent.y = SprRotC.y * SprScale.y
        
        With SourceRect
          .Left = iCol * AniWidth
          .Right = .Left + AniWidth
          .Top = iRow * AniHeight
          .bottom = .Top + AniHeight
        End With
          
        If bCenterScale Then
          Sprites.Draw TPool(AniTexture).GetTexture, SourceRect, sVec, rCent, SprRot, vec2(SprPos.x + offX + ScrollX - ((DestRect.Right - DestRect.Left) / 2), SprPos.y + offY + ScrollY - ((DestRect.bottom - DestRect.Top) / 2)), RGBA(SprCR, SprCG, SprCB, SprCA)
        Else
          Sprites.Draw TPool(AniTexture).GetTexture, SourceRect, sVec, rCent, SprRot, vec2(SprPos.x + offX + ScrollX, SprPos.y + offY + ScrollY), RGBA(SprCR, SprCG, SprCB, SprCA)
        End If
      'End If
    'End If
  End If
    
End Sub

Private Sub Class_Initialize()
  ReDim Frame(0)
End Sub

Public Sub SetAutoMovement(ByVal mX As Single, ByVal mY As Single, ByVal IntervalMs As Long)
  AMove.x = mX
  AMove.y = mY
  AMoveInterval = IntervalMs
End Sub

Public Function GetAutoMovement() As D3DVECTOR2
  GetAutoMovement.x = AMove.x
  GetAutoMovement.y = AMove.y
End Function

Public Sub SetAutoRotation(ByVal RotDegree As Single, ByVal IntervalMs As Long)
  ARotAngle = D2R(RotDegree)
  ARotInterval = IntervalMs
End Sub

Public Function GetAutoRotation() As Single
  GetAutoRotation = ARotAngle
End Function

Public Sub SetFrameDelay(ByVal lFrame As Long, ByVal lDelay As Long)
  With Frame(lFrame)
    If lDelay > -1 Then .lDelay = lDelay
  End With
End Sub

Public Sub SetFrameColor(ByVal lFrame As Long, ByVal lColor As Long, ByVal lValue As Long, ByVal bDelta As Boolean)
  With Frame(lFrame)
    Select Case lColor
      Case 1
        .lR.bDelta = bDelta
        .lR.lValue = lValue
      Case 2
        .lG.bDelta = bDelta
        .lG.lValue = lValue
      Case 3
        .lB.bDelta = bDelta
        .lB.lValue = lValue
      Case 4
        .lA.bDelta = bDelta
        .lA.lValue = lValue
    End Select
  End With
End Sub

Public Sub SetAutoFade(ldestR As Integer, ldestG As Integer, ldestB As Integer, ldestA As Integer, TotalMs As Long, Optional bDisable As Boolean = False)
  'If any color value is to be remained unchanged, enter an unvalid value (<0, >255)
  If ldestR >= 0 And ldestR <= 255 Then
    AColRGoal = ldestR
    AColRDelta = (ldestR - SprCR) / (TotalMs / 10)
  Else
    AColRGoal = SprCR
  End If
  If ldestG >= 0 And ldestG <= 255 Then
    AColGGoal = ldestG
    AColGDelta = (ldestG - SprCG) / (TotalMs / 10)
  Else
    AColGGoal = SprCG
  End If
  If ldestB >= 0 And ldestB <= 255 Then
    AColBGoal = ldestB
    AColBDelta = (ldestB - SprCB) / (TotalMs / 10)
  Else
    AColBGoal = SprCB
  End If
  If ldestA >= 0 And ldestA <= 255 Then
    AColAGoal = ldestA
    AColADelta = (ldestA - SprCA) / (TotalMs / 10)
  Else
    AColAGoal = SprCA
  End If
  AColDisable = bDisable
End Sub

Public Sub SetAutoScale(destX As Single, destY As Single, TotalMs As Long)
  AScale.x = destX
  AScale.y = destY
  AScaleDelta.x = (AScale.x - SprScale.x) / (TotalMs / 10)
  AScaleDelta.y = (AScale.y - SprScale.y) / (TotalMs / 10)
End Sub

Public Sub SetScale(x As Single, y As Single)
  SprScale.x = x
  SprScale.y = y
End Sub

Public Function GetScale() As D3DVECTOR2
  GetScale.x = SprScale.x
  GetScale.y = SprScale.y
End Function

Public Sub SetCenterRot(x As Single, y As Single)
  SprRotC.x = x: SprRotC.y = y
End Sub

Public Function GetCenterRot() As D3DVECTOR2
  GetCenterRot.x = SprRotC.x
  GetCenterRot.y = SprRotC.y
End Function

Public Sub SetColor(r As Byte, g As Byte, b As Byte, A As Byte)
  SprCA = A
  SprCR = r
  SprCG = g
  SprCB = b
End Sub

Public Sub SetRotation(RotDegree As Single)
  SprRot = D2R(RotDegree)
End Sub

Public Function GetRotationD() As Single
  GetRotationD = R2D(SprRot)
End Function

Public Sub SetPosition(x As Single, y As Single)
  SprPos.x = x
  SprPos.y = y
End Sub

Public Function GetPosition() As D3DVECTOR2
  GetPosition = SprPos
End Function

Public Function GetDestRect() As RECT
  GetDestRect = DestRect
End Function

Public Function GetUpperLeftCorner() As D3DVECTOR2
  If bCenterScale Then
    GetUpperLeftCorner.x = SprPos.x - AniWidth / 2 * SprScale.x
    GetUpperLeftCorner.y = SprPos.y - AniHeight / 2 * SprScale.y
  Else
    GetUpperLeftCorner.x = SprPos.x
    GetUpperLeftCorner.y = SprPos.y
  End If
End Function

Public Function GetWidth() As Long
  GetWidth = AniWidth * SprScale.x
End Function

Public Function GetHeight() As Long
  GetHeight = AniHeight * SprScale.y
End Function

Public Sub SetAutoPath(destX As Single, destY As Single, TotalMs As Long)
  APathGoal.x = destX
  APathGoal.y = destY
  APathDelta.x = (APathGoal.x - SprPos.x) / (TotalMs / 10)
  APathDelta.y = (APathGoal.y - SprPos.y) / (TotalMs / 10)
End Sub
