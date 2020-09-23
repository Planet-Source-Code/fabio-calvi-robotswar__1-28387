VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRobot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private PosX As Single, PosY As Single, Speed As Single
Private destX As Integer, destY As Integer, DestSet As Boolean
Private MoveX As Single, MoveY As Single
Private Index As Integer
Private CurrentKeyFrame As Integer
Private CRect As RECT
Private distS As Double, distus As Double, distnem As Double
Private StartTick As Long, DX As New DirectX7
Private CTick As Long, Bultick As Long
Private ViseManjeX As Integer, ViseManjeY As Integer
Private SpriteWidth As Integer, SpriteHeight As Integer
Private explokeyframe As Byte, pflg1 As Boolean, pflg2 As Boolean

Public Sub Init(Index As Integer)
If robots(Index).type = 1 Then
   SpriteWidth = 82
   SpriteHeight = 82
   CRect.Bottom = 82
   CRect.Right = 82
Else
   SpriteWidth = 72
   SpriteHeight = 72
   CRect.Bottom = 72
   CRect.Right = 72
End If

CurrentKeyFrame = 1
Speed = 80
explokeyframe = 0
sexpl = 0
Turn 9
pflg1 = False
pflg2 = True
ReDim robots(Index).bullet(0)
End Sub
Public Sub riposi(ind As Integer)
Dim i As Integer, j As Integer
PosX = 190
PosY = 190
robots(ind).x = PosX
robots(ind).y = PosY
For i = 0 To 5
    If i <> ind Then
       If (robots(ind).x >= robots(i).x - 72 And robots(ind).x <= robots(i).x + 72 And robots(ind).y <= robots(i).y + 72 And robots(ind).y >= robots(i).y - 72) Then
redo:
          If pflg1 = True Then
             PosX = PosX + 100
             PosY = PosY - 100
             If pflg2 = True Then pflg2 = False Else pflg2 = True
             If pflg2 = False Then
                PosX = PosX - 100
                PosY = PosY + 100
              End If
          Else
             PosX = PosX + 100
             PosY = PosY + 100
          End If
          robots(ind).x = PosX
          robots(ind).y = PosY
          For j = 0 To 5
              If j <> ind Then
                 If (robots(ind).x >= robots(j).x - 72 And robots(ind).x <= robots(j).x + 72 And robots(ind).y <= robots(j).y + 72 And robots(ind).y >= robots(j).y - 72) Then
                    If pflg1 = False Then pflg1 = True Else pflg1 = False
                    GoTo redo
                 End If
              End If
          Next j
       End If
    End If
Next i

robots(ind).hit = 0
robots(ind).status = True
Turn 9

End Sub
Public Sub UpdateAI()
Dim j As Integer
      For j = 0 To 5
            GetDistance enemy(j).x, enemy(j).y
            distnem = distS
            If distnem <= 250 Then
               robots(Index).appoX = robots(Index).odestX
               robots(Index).appoY = robots(Index).odestY
               robots(Index).odestX = PosX
               robots(Index).odestY = PosY
               SetDestination Int(robots(Index).odestX), Int(robots(Index).odestY)
               If DX.TickCount >= Bultick + 150 Then
                  Bultick = DX.TickCount
                  Getdir Abs(PosX + SpriteWidth) / 2, Abs(PosY + SpriteHeight) / 2, Abs(enemy(j).x + SpriteWidth) / 2, Abs(enemy(j).y + SpriteHeight) / 2
                  Turn dir
                  If robots(Index).type = 1 Then Call PlaySound(dsGun1)
                  If robots(Index).type = 2 Then Call PlaySound(dsFire1)
                  shoot Index, dir
               End If
            Else
               If robots(Index).appoX > 0 And robots(Index).appoY > 0 Then
                  robots(Index).odestX = robots(Index).appoX
                  robots(Index).odestY = robots(Index).appoY
                  robots(Index).appoX = 0
                  robots(Index).appoY = 0
               End If
            End If
        Next

End Sub

Public Sub UpdateAnimation()
If DestSet Then
    If DX.TickCount >= StartTick + 1000 / 2 Then
        StartTick = DX.TickCount
        If Not CurrentKeyFrame = 2 Then CurrentKeyFrame = CurrentKeyFrame + 1 Else CurrentKeyFrame = 1
        CRect.Left = CurrentKeyFrame * SpriteWidth - SpriteWidth
        CRect.Right = CRect.Left + SpriteWidth
    End If
Else
    CurrentKeyFrame = 1
End If
End Sub

Public Sub Turn(Side As Integer)
CRect.Left = 0
CRect.Top = 0
CRect.Right = CRect.Left + SpriteWidth
CRect.Top = Side * SpriteHeight
CRect.Bottom = CRect.Top + SpriteHeight
End Sub

Public Sub Draw(dds As DirectDrawSurface7, destSurface As DirectDrawSurface7, Index)
If robots(Index).status = True Then destSurface.BltFast -ViewPortX + (Int(PosX) - SpriteWidth / 2), -ViewPortY + (Int(PosY) - (SpriteHeight - 10)), dds, CRect, DDBLTFAST_WAIT + DDBLTFAST_SRCCOLORKEY
soldx = PosX
soldy = PosY
robots(Index).x = PosX
robots(Index).y = PosY
End Sub
Private Function GetCoef() As Single
GetCoef = (DX.TickCount - CTick) / 1000
End Function
Public Sub UpdateMove(indice As Integer)
Dim i As Integer, flag As Byte, Coef As Single, j As Integer
Index = indice
If DestSet And CTick < DX.TickCount Then
    VerifyPosition

    Coef = GetCoef
    If Int(PosX) = Int(pos1x) And Int(PosY) = Int(pos1y) Then
       SetDestination Int(robots(Index).odestX), Int(robots(Index).odestY)
       Getdir PosX, PosY, robots(Index).odestX, robots(Index).odestY
       Turn dir
       pos1x = -1
       pos1y = -1
    Else
       PosX = PosX + MoveX * Coef
       PosY = PosY + MoveY * Coef
    End If
         
    For j = 1 To 5
       If PosX + 72 >= nodes(j).x + 50 And PosX < nodes(j).x + 180 And PosY + 72 >= nodes(j).y + 50 And PosY < nodes(j).y + 180 Then nodes(j).owner = 1
    Next
    
    For i = 0 To 5
        If i <> Index Then
          
           If (PosX >= robots(i).x - 60 And PosX < robots(i).x + 60 And PosY <= robots(i).y + 60 And PosY > robots(i).y - 60) Or TableMap(Int(PosX / 130), Int(PosY / 130)) = 1 Then

              PosX = PosX - MoveX * Coef
              PosY = PosY - MoveY * Coef
              flag = (4 * Rnd) + 1
redo:
              If flag = 1 Then
                 pos1x = PosX - 36
                 pos1y = PosY + 36
              ElseIf flag = 2 Then
                 pos1x = PosX - 36
                 pos1y = PosY - 36
              ElseIf flag = 3 Then
                 pos1x = PosX + 36
                 pos1y = PosY - 36
              Else
                 pos1x = PosX + 36
                 pos1y = PosY + 36
              End If

              If (pos1x >= robots(i).x - 60 And pos1x < robots(i).x + 60 And pos1y <= robots(i).y + 60 And pos1y > robots(i).y - 60) Or TableMap(Int(pos1x / 130), Int(pos1y / 130)) = 1 Then
                 flag = flag + 1
                 If flag > 4 Then flag = 1
                 GoTo redo
              End If
              SetDestination Int(pos1x), Int(pos1y)
              Getdir PosX, PosY, pos1x, pos1y
              Turn dir
           End If
        End If
    Next
    
    CTick = DX.TickCount
End If
End Sub

Public Sub SetDestination(x As Integer, y As Integer)
Dim CoefX As Single, CoefY As Single, SizeY As Double, SizeX As Double, SizeC As Double
destX = x
destY = y
SizeX = Abs(destX - PosX)
SizeY = Abs(destY - PosY)
SizeC = Sqr(SizeX * SizeX + SizeY * SizeY)
If SizeX = 0 Or SizeY = 0 Then Exit Sub
CoefX = SizeX / SizeC
CoefY = SizeY / SizeC

Coef = SizeY / SizeX
If destX > PosX Then
    MoveX = Speed * CoefX
    ViseManjeX = 1
ElseIf destX < PosX Then
    MoveX = (Speed * CoefX) * -1
    ViseManjeX = 0
End If
If destY > PosY Then
    MoveY = Speed * CoefY
    ViseManjeY = 1
ElseIf destY < PosY Then
    MoveY = (Speed * CoefY) * -1
    ViseManjeY = 0
End If

DestSet = True
CTick = DX.TickCount
End Sub

Public Sub GetDistance(x As Single, y As Single)
Dim DistY As Double, DistX As Double

DistX = Abs(x - PosX)
DistY = Abs(y - PosY)
distS = Sqr(DistX * DistX + DistY * DistY)

End Sub

Public Sub SetPosition(x As Integer, y As Integer)
PosX = x
PosY = y
CurrentKeyFrame = 1
StopMovement
End Sub

Public Sub GetDestination(xOut As Single, yOut As Single)
xOut = destX
yOut = destY
End Sub

Public Sub GetPosition(xOut As Single, yOut As Single)
xOut = PosX
yOut = PosY
End Sub

Private Sub VerifyPosition()
If ViseManjeX = 1 And PosX >= destX And ViseManjeY = 1 And PosY >= destY Then EmergencyStop
If ViseManjeX = 1 And PosX >= destX And ViseManjeY = 0 And PosY <= destY Then EmergencyStop
If ViseManjeX = 0 And PosX <= destX And ViseManjeY = 1 And PosY >= destY Then EmergencyStop
If ViseManjeX = 0 And PosX <= destX And ViseManjeY = 0 And PosY <= destY Then EmergencyStop
End Sub

Public Function GetTurn() As Integer
GetTurn = CRect.Top / SpriteHeight
End Function

Public Sub StopMovement()
DestSet = False
End Sub

Private Sub EmergencyStop()
GetDistance robots(Index).odestX, robots(Index).odestY
distus = distS
If distus > 10 Then
   SetDestination Int(robots(Index).odestX), Int(robots(Index).odestY)
   Getdir PosX, PosY, robots(Index).odestX, robots(Index).odestY
   Turn dir
Else
   PosX = robots(Index).odestX
   PosY = robots(Index).odestY
   DestSet = False
End If
End Sub

Private Function Collision(Coef As Single) As Boolean
Dim Result As Integer

Result = TableMap(Int((PosX + MoveX * Coef) / 130), Int((PosY + MoveY * Coef) / 130))

If Result = 1 Then
   If Int(PosX) <> Int(robots(Index).odestX) And Int(PosY) <> Int(robots(Index).odestY) Then
      SetDestination Int(robots(Index).odestX), Int(robots(Index).odestY)
      Getdir PosX, PosY, robots(Index).odestX, robots(Index).odestY
      Turn dir
   End If
   Collision = True
Else
   Collision = False
End If

End Function

Public Sub Destroy()
Set DX = Nothing
End Sub
Public Sub Getdir(orgx As Single, orgy As Single, tarx As Single, tary As Single)
Dim rad As Double, Angle As Byte, xdiff As Single, ydiff As Single
xdiff = Abs(orgx - tarx)
ydiff = Abs(orgy - tary)
If ydiff <= 0 Then ydiff = 1
rad = Atn(xdiff / ydiff)
Angle = rad * (180 / pi)
   If tarx > orgx Then
      If tary < orgy Then
         If Angle < 15 Then
            dir = 0
         ElseIf Angle >= 15 And Angle < 30 Then dir = 1
         ElseIf Angle >= 30 And Angle < 45 Then dir = 2
         ElseIf Angle >= 45 And Angle < 60 Then dir = 3
         ElseIf Angle >= 60 And Angle < 75 Then dir = 4
         Else
            dir = 5
         End If
       Else
         If Angle > 75 Then
            dir = 6
         ElseIf Angle > 60 And Angle <= 75 Then dir = 7
         ElseIf Angle > 45 And Angle <= 60 Then dir = 8
         ElseIf Angle > 30 And Angle <= 45 Then dir = 9
         ElseIf Angle > 15 And Angle <= 30 Then dir = 10
         Else
            dir = 11
         End If
       End If
   Else
       If orgy < tary Then
          If Angle < 15 Then
             dir = 12
          ElseIf Angle >= 15 And Angle < 30 Then dir = 13
          ElseIf Angle >= 30 And Angle < 45 Then dir = 14
          ElseIf Angle >= 45 And Angle < 60 Then dir = 15
          ElseIf Angle >= 60 And Angle < 75 Then dir = 16
          Else
             dir = 17
          End If
       Else
          If Angle > 75 Then
             dir = 18
          ElseIf Angle > 60 And Angle <= 75 Then dir = 19
          ElseIf Angle > 45 And Angle <= 60 Then dir = 20
          ElseIf Angle > 30 And Angle <= 45 Then dir = 21
          ElseIf Angle > 15 And Angle <= 30 Then dir = 22
          Else
             dir = 23
          End If
       End If
   End If
End Sub
Private Sub shoot(Index As Integer, dir As Integer)
Dim i As Integer

i = UBound(robots(Index).bullet)

If dir = 0 Then
      robots(Index).bullet(i).lx = robots(Index).x - 27
      robots(Index).bullet(i).ly = robots(Index).y - 15
      robots(Index).bullet(i).rx = robots(Index).x + 26
      robots(Index).bullet(i).ry = robots(Index).y - 15
      robots(Index).bullet(i).cx = robots(Index).x - 3
      robots(Index).bullet(i).cy = robots(Index).y + 3
End If
If dir = 1 Then
      robots(Index).bullet(i).lx = robots(Index).x + 17
      robots(Index).bullet(i).ly = robots(Index).y - 7
      robots(Index).bullet(i).rx = robots(Index).x - 34
      robots(Index).bullet(i).ry = robots(Index).y - 20
      robots(Index).bullet(i).cx = robots(Index).x - 12
      robots(Index).bullet(i).cy = robots(Index).y + 4
End If
If dir = 2 Then
      robots(Index).bullet(i).lx = robots(Index).x + 10
      robots(Index).bullet(i).ly = robots(Index).y + 1
      robots(Index).bullet(i).rx = robots(Index).x - 36
      robots(Index).bullet(i).ry = robots(Index).y - 26
      robots(Index).bullet(i).cx = robots(Index).x - 18
      robots(Index).bullet(i).cy = robots(Index).y + 1
End If
If dir = 3 Then
      robots(Index).bullet(i).lx = robots(Index).x - 16
      robots(Index).bullet(i).ly = robots(Index).y + 19
      robots(Index).bullet(i).rx = robots(Index).x - 53
      robots(Index).bullet(i).ry = robots(Index).y - 18
      robots(Index).bullet(i).cx = robots(Index).x - 39
      robots(Index).bullet(i).cy = robots(Index).y + 11
End If
If dir = 4 Then
      robots(Index).bullet(i).lx = robots(Index).x - 8
      robots(Index).bullet(i).ly = robots(Index).y + 5
      robots(Index).bullet(i).rx = robots(Index).x - 33
      robots(Index).bullet(i).ry = robots(Index).y - 41
      robots(Index).bullet(i).cx = robots(Index).x - 29
      robots(Index).bullet(i).cy = robots(Index).y - 10
End If
If dir = 5 Then
      robots(Index).bullet(i).lx = robots(Index).x - 14
      robots(Index).bullet(i).ly = robots(Index).y + 2
      robots(Index).bullet(i).rx = robots(Index).x - 27
      robots(Index).bullet(i).ry = robots(Index).y - 52
      robots(Index).bullet(i).cx = robots(Index).x - 31
      robots(Index).bullet(i).cy = robots(Index).y - 19
End If
If dir = 6 Then
      robots(Index).bullet(i).lx = robots(Index).x - 19
      robots(Index).bullet(i).ly = robots(Index).y - 60
      robots(Index).bullet(i).rx = robots(Index).x - 18
      robots(Index).bullet(i).ry = robots(Index).y - 5
      robots(Index).bullet(i).cx = robots(Index).x - 34
      robots(Index).bullet(i).cy = robots(Index).y - 29
End If
If dir = 7 Then
   robots(Index).bullet(i).lx = robots(Index).x - 14
   robots(Index).bullet(i).ly = robots(Index).y - 67
   robots(Index).bullet(i).rx = robots(Index).x - 28
   robots(Index).bullet(i).ry = robots(Index).y - 14
   robots(Index).bullet(i).cx = robots(Index).x - 36
   robots(Index).bullet(i).cy = robots(Index).y - 38
End If
If dir = 8 Then
   robots(Index).bullet(i).lx = robots(Index).x - 7
   robots(Index).bullet(i).ly = robots(Index).y - 70
   robots(Index).bullet(i).rx = robots(Index).x - 34
   robots(Index).bullet(i).ry = robots(Index).y - 22
   robots(Index).bullet(i).cx = robots(Index).x - 32
   robots(Index).bullet(i).cy = robots(Index).y - 46
End If
If dir = 9 Then
   robots(Index).bullet(i).lx = robots(Index).x - 15
   robots(Index).bullet(i).ly = robots(Index).y - 84
   robots(Index).bullet(i).rx = robots(Index).x - 53
   robots(Index).bullet(i).ry = robots(Index).y - 46
   robots(Index).bullet(i).cx = robots(Index).x - 43
   robots(Index).bullet(i).cy = robots(Index).y - 67
End If
If dir = 10 Then
   robots(Index).bullet(i).lx = robots(Index).x + 9
   robots(Index).bullet(i).ly = robots(Index).y - 66
   robots(Index).bullet(i).rx = robots(Index).x - 38
   robots(Index).bullet(i).ry = robots(Index).y - 39
   robots(Index).bullet(i).cx = robots(Index).x - 20
   robots(Index).bullet(i).cy = robots(Index).y - 56
End If
If dir = 11 Then
   robots(Index).bullet(i).lx = robots(Index).x + 19
   robots(Index).bullet(i).ly = robots(Index).y - 60
   robots(Index).bullet(i).rx = robots(Index).x - 35
   robots(Index).bullet(i).ry = robots(Index).y - 47
   robots(Index).bullet(i).cx = robots(Index).x - 11
   robots(Index).bullet(i).cy = robots(Index).y - 58
End If
If dir = 12 Then
   robots(Index).bullet(i).lx = robots(Index).x + 27
   robots(Index).bullet(i).ly = robots(Index).y - 50
   robots(Index).bullet(i).rx = robots(Index).x - 27
   robots(Index).bullet(i).ry = robots(Index).y - 50
   robots(Index).bullet(i).cx = robots(Index).x - 3
   robots(Index).bullet(i).cy = robots(Index).y - 61
End If
If dir = 13 Then
   robots(Index).bullet(i).lx = robots(Index).x + 31
   robots(Index).bullet(i).ly = robots(Index).y - 45
   robots(Index).bullet(i).rx = robots(Index).x - 22
   robots(Index).bullet(i).ry = robots(Index).y - 60
   robots(Index).bullet(i).cx = robots(Index).x + 5
   robots(Index).bullet(i).cy = robots(Index).y - 62
End If
If dir = 14 Then
   robots(Index).bullet(i).lx = robots(Index).x + 35
   robots(Index).bullet(i).ly = robots(Index).y - 39
   robots(Index).bullet(i).rx = robots(Index).x - 11
   robots(Index).bullet(i).ry = robots(Index).y - 66
   robots(Index).bullet(i).cx = robots(Index).x + 14
   robots(Index).bullet(i).cy = robots(Index).y - 60
End If
If dir = 15 Then
   robots(Index).bullet(i).lx = robots(Index).x + 13
   robots(Index).bullet(i).ly = robots(Index).y - 83
   robots(Index).bullet(i).rx = robots(Index).x + 51
   robots(Index).bullet(i).ry = robots(Index).y - 45
   robots(Index).bullet(i).cx = robots(Index).x + 34
   robots(Index).bullet(i).cy = robots(Index).y - 68
End If
If dir = 16 Then
   robots(Index).bullet(i).lx = robots(Index).x + 6
   robots(Index).bullet(i).ly = robots(Index).y - 72
   robots(Index).bullet(i).rx = robots(Index).x + 30
   robots(Index).bullet(i).ry = robots(Index).y - 23
   robots(Index).bullet(i).cx = robots(Index).x + 24
   robots(Index).bullet(i).cy = robots(Index).y - 47
End If
If dir = 17 Then
   robots(Index).bullet(i).lx = robots(Index).x + 10
   robots(Index).bullet(i).ly = robots(Index).y - 68
   robots(Index).bullet(i).rx = robots(Index).x + 24
   robots(Index).bullet(i).ry = robots(Index).y - 14
   robots(Index).bullet(i).cx = robots(Index).x + 27
   robots(Index).bullet(i).cy = robots(Index).y - 39
End If
If dir = 18 Then
   robots(Index).bullet(i).lx = robots(Index).x + 16
   robots(Index).bullet(i).ly = robots(Index).y - 60
   robots(Index).bullet(i).rx = robots(Index).x + 16
   robots(Index).bullet(i).ry = robots(Index).y - 5
   robots(Index).bullet(i).cx = robots(Index).x + 29
   robots(Index).bullet(i).cy = robots(Index).y - 29
End If
If dir = 19 Then
   robots(Index).bullet(i).lx = robots(Index).x + 23
   robots(Index).bullet(i).ly = robots(Index).y - 51
   robots(Index).bullet(i).rx = robots(Index).x + 11
   robots(Index).bullet(i).ry = robots(Index).y + 1
   robots(Index).bullet(i).cx = robots(Index).x + 30
   robots(Index).bullet(i).cy = robots(Index).y - 19
End If
If dir = 20 Then
   robots(Index).bullet(i).lx = robots(Index).x + 30
   robots(Index).bullet(i).ly = robots(Index).y - 41
   robots(Index).bullet(i).rx = robots(Index).x + 4
   robots(Index).bullet(i).ry = robots(Index).y + 6
   robots(Index).bullet(i).cx = robots(Index).x + 26
   robots(Index).bullet(i).cy = robots(Index).y - 11
End If
If dir = 21 Then
   robots(Index).bullet(i).lx = robots(Index).x + 50
   robots(Index).bullet(i).ly = robots(Index).y - 18
   robots(Index).bullet(i).rx = robots(Index).x + 11
   robots(Index).bullet(i).ry = robots(Index).y + 21
   robots(Index).bullet(i).cx = robots(Index).x + 36
   robots(Index).bullet(i).cy = robots(Index).y + 9
End If
If dir = 22 Then
   robots(Index).bullet(i).lx = robots(Index).x + 35
   robots(Index).bullet(i).ly = robots(Index).y - 27
   robots(Index).bullet(i).rx = robots(Index).x - 13
   robots(Index).bullet(i).ry = robots(Index).y '+ 1
   robots(Index).bullet(i).cx = robots(Index).x + 14
   robots(Index).bullet(i).cy = robots(Index).y - 2
End If
If dir = 23 Then
   robots(Index).bullet(i).lx = robots(Index).x + 31
   robots(Index).bullet(i).ly = robots(Index).y - 19
   robots(Index).bullet(i).rx = robots(Index).x - 22
   robots(Index).bullet(i).ry = robots(Index).y - 5
   robots(Index).bullet(i).cx = robots(Index).x + 6
   robots(Index).bullet(i).cy = robots(Index).y - 1
End If
If robots(Index).type = 1 Then
   robots(Index).bullet(i).llive = True
   robots(Index).bullet(i).rlive = True
Else
   robots(Index).bullet(i).clive = True
End If
robots(Index).bullet(i).buldir = dir
robots(Index).bullet(i).Speed = 50

End Sub
Public Sub Explosion(dds As DirectDrawSurface7, destSurface As DirectDrawSurface7, x As Integer, y As Integer, Index As Integer)
Dim CRect As RECT, explowidth As Integer, exploheight As Integer
explowidth = 124
exploheight = 128
CRect.Bottom = 128
CRect.Right = 0
CRect.Left = explokeyframe * explowidth - explowidth
CRect.Right = CRect.Left + explowidth
DestSet = False
robots(Index).odestX = -1
robots(Index).odestY = -1
destSurface.BltFast x - explowidth / 2, y - exploheight / 2, dds, CRect, DDBLTFAST_WAIT + DDBLTFAST_SRCCOLORKEY
explokeyframe = explokeyframe + 1
If (sexpl = 0 And explokeyframe > 17) Or (sexpl = 1 And explokeyframe > 15) Then
   If sexpl = 0 Then sexpl = 1 Else sexpl = 0
   explokeyframe = 0
   riposi Index
End If
End Sub

