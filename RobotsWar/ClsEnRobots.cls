VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsEnRobots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private PosX As Single, PosY As Single, Speed As Single
Private destX As Integer, destY As Integer, DestSet As Boolean
Private MoveX As Single, MoveY As Single
Private Index As Integer
Private distE As Double, distsu As Double, distnem As Double
Private CurrentKeyFrame As Integer
Private CRect As RECT
Private StartTick As Long, DX As New DirectX7
Private CTick As Long, Bultick As Long
Private ViseManjeX As Integer, ViseManjeY As Integer
Private SpriteWidth As Integer, SpriteHeight As Integer
Private explokeyframe As Byte, pflg1 As Boolean, pflg2 As Boolean

Public Sub Init(Index As Integer)
If enemy(Index).type = 1 Then
   SpriteWidth = 82
   SpriteHeight = 82
   CRect.Bottom = 82
   CRect.Right = 82
Else
   SpriteWidth = 80
   SpriteHeight = 80
   CRect.Bottom = 80
   CRect.Right = 80
End If
CurrentKeyFrame = 1
Speed = 80
enemy(Index).appoX = 0
enemy(Index).appoY = 0
explokeyframe = 0
eexpl = 0
Turn 21
pflg1 = False
pflg2 = True

ReDim enemy(Index).bullet(0)

Pick1 Index

End Sub
Public Sub riposi(ind As Integer)
Dim i As Integer, j As Integer, k As Integer
PosX = 2380
PosY = 2380
enemy(ind).x = PosX
enemy(ind).y = PosY
For i = 0 To 5
    If i <> ind Then
       If (enemy(ind).x >= enemy(i).x - 82 And enemy(ind).x <= enemy(i).x + 82 And enemy(ind).y <= enemy(i).y + 82 And enemy(ind).y >= enemy(i).y - 82) Then
redo:
          If pflg1 = True Then
             PosX = PosX - 100
             PosY = PosY + 100
             If pflg2 = True Then pflg2 = False Else pflg2 = True
             If pflg2 = False Then
                PosX = PosX + 100
                PosY = PosY - 100
             End If
          Else
             PosX = PosX - 100
             PosY = PosY - 100
          End If
          enemy(ind).x = PosX
          enemy(ind).y = PosY
          For j = 0 To 5
              If j <> ind Then
                 If (enemy(ind).x >= enemy(j).x - 82 And enemy(ind).x <= enemy(j).x + 82 And enemy(ind).y <= enemy(j).y + 82 And enemy(ind).y >= enemy(j).y - 82) Then
                    If pflg1 = False Then pflg1 = True Else pflg1 = False
                    GoTo redo
                 End If
              End If
          Next j
       End If
    End If
Next i

enemy(ind).hit = 0
enemy(ind).status = True
enemy(ind).appoX = 0
enemy(ind).appoY = 0

If my > 0 Then
   enemy(ind).odestX = nodes(nodedist(my).ind).x
   enemy(ind).odestY = nodes(nodedist(my).ind).y
Else
   enemy(ind).odestX = Int(nodes(nodedist(1).ind).x)
   enemy(ind).odestY = Int(nodes(nodedist(1).ind).y)
   For k = 1 To 5
       If nodes(nodedist(k).ind).owner = 0 Then
          enemy(ind).odestX = Int(nodes(nodedist(k).ind).x)
          enemy(ind).odestY = Int(nodes(nodedist(k).ind).y)
          Exit For
       End If
   Next
End If
SetDestination Int(enemy(ind).odestX), Int(enemy(ind).odestY)
Getdir enemy(ind).x, enemy(ind).y, enemy(ind).odestX, enemy(ind).odestY
Turn dir

End Sub
Public Sub UpdateAI()
Dim j As Integer
      For j = 0 To 5
            GetDistance robots(j).x, robots(j).y
            distnem = distE
            If distnem <= 250 Then
               enemy(Index).appoX = enemy(Index).odestX
               enemy(Index).appoY = enemy(Index).odestY
               enemy(Index).odestX = PosX
               enemy(Index).odestY = PosY
               SetDestination Int(enemy(Index).odestX), Int(enemy(Index).odestY)
               If DX.TickCount >= Bultick + 150 Then
                  Bultick = DX.TickCount
                  Getdir Abs(PosX + SpriteWidth) / 2, Abs(PosY + SpriteHeight) / 2, Abs(robots(j).x + SpriteWidth) / 2, Abs(robots(j).y + SpriteHeight) / 2
                  Turn dir
                  If enemy(Index).type = 1 Then Call PlaySound(dsGun2)
                  If enemy(Index).type = 2 Then Call PlaySound(dsFire2)
                  shoot Index, dir
               End If
            Else
               If enemy(Index).appoX > 0 And enemy(Index).appoY > 0 Then
                  enemy(Index).odestX = enemy(Index).appoX
                  enemy(Index).odestY = enemy(Index).appoY
                  enemy(Index).appoX = 0
                  enemy(Index).appoY = 0
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
If enemy(Index).status = True Then destSurface.BltFast -ViewPortX + (Int(PosX) - SpriteWidth / 2), -ViewPortY + (Int(PosY) - (SpriteHeight - 10)), dds, CRect, DDBLTFAST_WAIT + DDBLTFAST_SRCCOLORKEY
eoldx = PosX
eoldy = PosY
enemy(Index).x = PosX
enemy(Index).y = PosY
End Sub
Private Function GetCoef() As Single
GetCoef = (DX.TickCount - CTick) / 1000
End Function
Public Sub UpdateMove(indice As Integer)
Dim i As Integer, flag As Byte, Coef As Single, j As Integer
Index = indice

If CTick < DX.TickCount Then
    VerifyPosition
    If DestSet = False And flgup = True Then DestSet = True
    If DestSet = True Then
    Coef = GetCoef
    If Int(PosX) = Int(pos2x) And Int(PosY) = Int(pos2y) Then
       SetDestination Int(enemy(Index).odestX), Int(enemy(Index).odestY)
       Getdir PosX, PosY, enemy(Index).odestX, enemy(Index).odestY
       Turn dir
       pos2x = -1
       pos2y = -1
    ElseIf Not Collision(Coef) = True Then
       PosX = PosX + MoveX * Coef
       PosY = PosY + MoveY * Coef
    End If
         
    If flgup = True Then
       Pick2
       enemy(Index).odestX = nodes(my).x + 60
       enemy(Index).odestY = nodes(my).y + 60
    End If
         
         
    For j = 1 To 5
       If PosX + 72 >= nodes(j).x + 50 And PosX < nodes(j).x + 180 And PosY + 72 >= nodes(j).y + 50 And PosY < nodes(j).y + 180 Then nodes(j).owner = 2
    Next
    

    For i = 0 To 5
        If i <> Index Then
           If PosX >= enemy(i).x - 82 And PosX < enemy(i).x + 82 And PosY <= enemy(i).y + 82 And PosY > enemy(i).y - 82 Then
              PosX = PosX - MoveX * Coef
              PosY = PosY - MoveY * Coef
              flag = (4 * Rnd) + 1
redo:
              If flag = 1 Then
                 pos2x = PosX - 80
                 pos2y = PosY + 80
              ElseIf flag = 2 Then
                 pos2x = PosX + 80
                 pos2y = PosY - 80
              ElseIf flag = 3 Then
                 pos2x = PosX - 80
                 pos2y = PosY - 80
              Else
                 pos2x = PosX + 80
                 pos2y = PosY + 80
              End If

              If pos2x >= enemy(i).x - 82 And pos2x < enemy(i).x + 82 And pos2y <= enemy(i).y + 82 And pos2y > enemy(i).y - 82 Then
                 flag = flag + 1
                 If flag > 4 Then flag = 1
                 GoTo redo
              End If
              SetDestination Int(pos2x), Int(pos2y)
              Getdir PosX, PosY, pos2x, pos2y
              Turn dir
           End If
        End If
    
   
    Next
    End If
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
distE = Sqr(DistX * DistX + DistY * DistY)

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
GetDistance enemy(Index).odestX, enemy(Index).odestY
distsu = distE
If distsu > 10 Then
   SetDestination Int(enemy(Index).odestX), Int(enemy(Index).odestY)
   Getdir PosX, PosY, enemy(Index).odestX, enemy(Index).odestY
   Turn dir
Else
   PosX = enemy(Index).odestX
   PosY = enemy(Index).odestY
   DestSet = False
End If
End Sub

Private Function Collision(Coef As Single) As Boolean
Dim Result As Integer

Result = TableMap(Int((PosX + MoveX * Coef) / 130), Int((PosY + MoveY * Coef) / 130))

If Result = 1 Then
   If Int(PosX) <> Int(enemy(Index).odestX) And Int(PosY) <> Int(enemy(Index).odestY) Then
      SetDestination Int(enemy(Index).odestX), Int(enemy(Index).odestY)
      Getdir PosX, PosY, enemy(Index).odestX, enemy(Index).odestY
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
Private Sub Pick1(Index As Integer)
Dim i As Byte, j As Byte, tempi As Byte, tempd As Double, rand As Byte

'Get distance from the nodes
For i = 1 To 5
    GetDistance nodes(i).x, nodes(i).y
    nodedist(i).ind = i
    nodedist(i).dist = distE
Next

'Sort the distance from the nodes
For j = 1 To 4
    For i = 1 To 4
        If nodedist(i).dist > nodedist(i + 1).dist Then
           tempi = nodedist(i + 1).ind
           tempd = nodedist(i + 1).dist
           nodedist(i + 1).ind = nodedist(i).ind
           nodedist(i + 1).dist = nodedist(i).dist
           nodedist(i).ind = tempi
           nodedist(i).dist = tempd
        End If
    Next i
Next j

rand = Int(Rnd * 3) + 1
'Go to the nearest (min distance) node
If rand = 1 Then
   If (Index = 0 Or Index = 3) And nodes(nodedist(1).ind).owner = 0 Then
      enemy(Index).odestX = nodes(nodedist(1).ind).x
      enemy(Index).odestY = nodes(nodedist(1).ind).y
      SetDestination Int(nodes(nodedist(1).ind).x), Int(nodes(nodedist(1).ind).y)
      If (Index = 0 Or Index = 3) Then Getdir enemy(Index).x, enemy(Index).y, nodes(nodedist(1).ind).x, nodes(nodedist(1).ind).y
   ElseIf (Index = 1 Or Index = 4) And nodes(nodedist(2).ind).owner = 0 Then
      enemy(Index).odestX = nodes(nodedist(2).ind).x
      enemy(Index).odestY = nodes(nodedist(2).ind).y
      SetDestination Int(nodes(nodedist(2).ind).x), Int(nodes(nodedist(2).ind).y)
      If (Index = 1 Or Index = 4) Then Getdir enemy(Index).x, enemy(Index).y, nodes(nodedist(2).ind).x, nodes(nodedist(2).ind).y
   ElseIf (Index = 2 Or Index = 5) And nodes(nodedist(3).ind).owner = 0 Then
      enemy(Index).odestX = nodes(nodedist(3).ind).x
      enemy(Index).odestY = nodes(nodedist(3).ind).y
      SetDestination Int(nodes(nodedist(3).ind).x), Int(nodes(nodedist(3).ind).y)
      If (Index = 2 Or Index = 5) Then Getdir enemy(Index).x, enemy(Index).y, nodes(nodedist(3).ind).x, nodes(nodedist(3).ind).y
   End If
ElseIf rand = 2 Then
   If Index <> 2 And Index <> 5 And nodes(nodedist(1).ind).owner = 0 Then
      enemy(Index).odestX = nodes(nodedist(1).ind).x
      enemy(Index).odestY = nodes(nodedist(1).ind).y
      SetDestination Int(nodes(nodedist(1).ind).x), Int(nodes(nodedist(1).ind).y)
      If (Index = 0 Or Index = 3) Then Getdir enemy(0).x, enemy(0).y, nodes(nodedist(1).ind).x, nodes(nodedist(1).ind).y
      If (Index = 1 Or Index = 4) Then Getdir enemy(1).x, enemy(1).y, nodes(nodedist(1).ind).x, nodes(nodedist(1).ind).y
   ElseIf (Index = 2 Or Index = 5) And nodes(nodedist(2).ind).owner = 0 Then
      enemy(Index).odestX = nodes(nodedist(2).ind).x
      enemy(Index).odestY = nodes(nodedist(2).ind).y
      SetDestination Int(nodes(nodedist(2).ind).x), Int(nodes(nodedist(2).ind).y)
      If (Index = 2 Or Index = 5) Then Getdir enemy(2).x, enemy(2).y, nodes(nodedist(2).ind).x, nodes(nodedist(2).ind).y
   End If
Else
   If nodes(nodedist(1).ind).owner = 0 Then
      enemy(Index).odestX = nodes(nodedist(1).ind).x
      enemy(Index).odestY = nodes(nodedist(1).ind).y
      SetDestination Int(nodes(nodedist(1).ind).x), Int(nodes(nodedist(1).ind).y)
      Getdir enemy(Index).x, enemy(Index).y, nodes(nodedist(1).ind).x, nodes(nodedist(1).ind).y
   End If
End If
Turn dir
DestSet = True
End Sub
Private Sub Pick2()
Dim i As Byte, j As Byte, tempi As Byte, tempd As Double, rand As Byte

'Get distance from the nodes
For i = 1 To 5
    GetDistance nodes(i).x, nodes(i).y
    nodedist(i).ind = i
    nodedist(i).dist = distE
Next

'Sort the distance from the nodes
For j = 1 To 4
    For i = 1 To 4
        If nodedist(i).dist > nodedist(i + 1).dist Then
           tempi = nodedist(i + 1).ind
           tempd = nodedist(i + 1).dist
           nodedist(i + 1).ind = nodedist(i).ind
           nodedist(i + 1).dist = nodedist(i).dist
           nodedist(i).ind = tempi
           nodedist(i).dist = tempd
        End If
    Next i
Next j

For i = 1 To 5
    If nodes(nodedist(i).ind).owner = 1 Then
       my = nodedist(i).ind
       Exit For
    End If
Next

End Sub
Private Sub shoot(Index As Integer, dir As Integer)
Dim i As Integer

i = UBound(enemy(Index).bullet)

If dir = 0 Then
   enemy(Index).bullet(i).lx = enemy(Index).x - 30
   enemy(Index).bullet(i).ly = enemy(Index).y - 15
   enemy(Index).bullet(i).rx = enemy(Index).x + 28
   enemy(Index).bullet(i).ry = enemy(Index).y - 15
   enemy(Index).bullet(i).cx = enemy(Index).x - 3
   enemy(Index).bullet(i).cy = enemy(Index).y + 15
End If
If dir = 1 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 20
   enemy(Index).bullet(i).ly = enemy(Index).y - 3
   enemy(Index).bullet(i).rx = enemy(Index).x - 35
   enemy(Index).bullet(i).ry = enemy(Index).y - 18
   enemy(Index).bullet(i).cx = enemy(Index).x - 15
   enemy(Index).bullet(i).cy = enemy(Index).y + 15
End If
If dir = 2 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 14
   enemy(Index).bullet(i).ly = enemy(Index).y + 2
   enemy(Index).bullet(i).rx = enemy(Index).x - 41
   enemy(Index).bullet(i).ry = enemy(Index).y - 29
   enemy(Index).bullet(i).cx = enemy(Index).x - 26
   enemy(Index).bullet(i).cy = enemy(Index).y + 10
End If
If dir = 3 Then
   enemy(Index).bullet(i).lx = enemy(Index).x - 13
   enemy(Index).bullet(i).ly = enemy(Index).y + 20
   enemy(Index).bullet(i).rx = enemy(Index).x - 55
   enemy(Index).bullet(i).ry = enemy(Index).y - 20
   enemy(Index).bullet(i).cx = enemy(Index).x - 51
   enemy(Index).bullet(i).cy = enemy(Index).y + 16
End If
If dir = 4 Then
   enemy(Index).bullet(i).lx = enemy(Index).x - 5
   enemy(Index).bullet(i).ly = enemy(Index).y + 8
   enemy(Index).bullet(i).rx = enemy(Index).x - 35
   enemy(Index).bullet(i).ry = enemy(Index).y - 45
   enemy(Index).bullet(i).cx = enemy(Index).x - 44
   enemy(Index).bullet(i).cy = enemy(Index).y - 8
End If
If dir = 5 Then
   enemy(Index).bullet(i).lx = enemy(Index).x - 14
   enemy(Index).bullet(i).ly = enemy(Index).y + 3
   enemy(Index).bullet(i).rx = enemy(Index).x - 29
   enemy(Index).bullet(i).ry = enemy(Index).y - 55
   enemy(Index).bullet(i).cx = enemy(Index).x - 50
   enemy(Index).bullet(i).cy = enemy(Index).y - 20
End If
If dir = 6 Then
   enemy(Index).bullet(i).lx = enemy(Index).x - 20
   enemy(Index).bullet(i).ly = enemy(Index).y - 62
   enemy(Index).bullet(i).rx = enemy(Index).x - 20
   enemy(Index).bullet(i).ry = enemy(Index).y - 5
   enemy(Index).bullet(i).cx = enemy(Index).x - 52
   enemy(Index).bullet(i).cy = enemy(Index).y - 34
End If
If dir = 7 Then
   enemy(Index).bullet(i).lx = enemy(Index).x - 14
   enemy(Index).bullet(i).ly = enemy(Index).y - 68
   enemy(Index).bullet(i).rx = enemy(Index).x - 30
   enemy(Index).bullet(i).ry = enemy(Index).y - 13
   enemy(Index).bullet(i).cx = enemy(Index).x - 53
   enemy(Index).bullet(i).cy = enemy(Index).y - 46
End If
If dir = 8 Then
   enemy(Index).bullet(i).lx = enemy(Index).x - 7
   enemy(Index).bullet(i).ly = enemy(Index).y - 72
   enemy(Index).bullet(i).rx = enemy(Index).x - 36
   enemy(Index).bullet(i).ry = enemy(Index).y - 20
   enemy(Index).bullet(i).cx = enemy(Index).x - 46
   enemy(Index).bullet(i).cy = enemy(Index).y - 57
End If
If dir = 9 Then
   enemy(Index).bullet(i).lx = enemy(Index).x - 14
   enemy(Index).bullet(i).ly = enemy(Index).y - 86
   enemy(Index).bullet(i).rx = enemy(Index).x - 55
   enemy(Index).bullet(i).ry = enemy(Index).y - 45
   enemy(Index).bullet(i).cx = enemy(Index).x - 53
   enemy(Index).bullet(i).cy = enemy(Index).y - 82
End If
If dir = 10 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 11
   enemy(Index).bullet(i).ly = enemy(Index).y - 66
   enemy(Index).bullet(i).rx = enemy(Index).x - 37
   enemy(Index).bullet(i).ry = enemy(Index).y - 37
   enemy(Index).bullet(i).cx = enemy(Index).x - 26
   enemy(Index).bullet(i).cy = enemy(Index).y - 74
End If
If dir = 11 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 21
   enemy(Index).bullet(i).ly = enemy(Index).y - 61
   enemy(Index).bullet(i).rx = enemy(Index).x - 37
   enemy(Index).bullet(i).ry = enemy(Index).y - 46
   enemy(Index).bullet(i).cx = enemy(Index).x - 15
   enemy(Index).bullet(i).cy = enemy(Index).y - 80
End If
If dir = 12 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 30
   enemy(Index).bullet(i).ly = enemy(Index).y - 50
   enemy(Index).bullet(i).rx = enemy(Index).x - 30
   enemy(Index).bullet(i).ry = enemy(Index).y - 50
   enemy(Index).bullet(i).cx = enemy(Index).x - 4
   enemy(Index).bullet(i).cy = enemy(Index).y - 82
End If
If dir = 13 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 32
   enemy(Index).bullet(i).ly = enemy(Index).y - 45
   enemy(Index).bullet(i).rx = enemy(Index).x - 25
   enemy(Index).bullet(i).ry = enemy(Index).y - 60
   enemy(Index).bullet(i).cx = enemy(Index).x + 9
   enemy(Index).bullet(i).cy = enemy(Index).y - 82
End If
If dir = 14 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 37
   enemy(Index).bullet(i).ly = enemy(Index).y - 38
   enemy(Index).bullet(i).rx = enemy(Index).x - 15
   enemy(Index).bullet(i).ry = enemy(Index).y - 67
   enemy(Index).bullet(i).cx = enemy(Index).x + 20
   enemy(Index).bullet(i).cy = enemy(Index).y - 78
End If
If dir = 15 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 10
   enemy(Index).bullet(i).ly = enemy(Index).y - 85
   enemy(Index).bullet(i).rx = enemy(Index).x + 52
   enemy(Index).bullet(i).ry = enemy(Index).y - 44
   enemy(Index).bullet(i).cx = enemy(Index).x + 45
   enemy(Index).bullet(i).cy = enemy(Index).y - 83
End If
If dir = 16 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 5
   enemy(Index).bullet(i).ly = enemy(Index).y - 72
   enemy(Index).bullet(i).rx = enemy(Index).x + 34
   enemy(Index).bullet(i).ry = enemy(Index).y - 21
   enemy(Index).bullet(i).cx = enemy(Index).x + 37
   enemy(Index).bullet(i).cy = enemy(Index).y - 57
End If
If dir = 17 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 11
   enemy(Index).bullet(i).ly = enemy(Index).y - 68
   enemy(Index).bullet(i).rx = enemy(Index).x + 26
   enemy(Index).bullet(i).ry = enemy(Index).y - 13
   enemy(Index).bullet(i).cx = enemy(Index).x + 44
   enemy(Index).bullet(i).cy = enemy(Index).y - 46
End If
If dir = 18 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 17
   enemy(Index).bullet(i).ly = enemy(Index).y - 62
   enemy(Index).bullet(i).rx = enemy(Index).x + 17
   enemy(Index).bullet(i).ry = enemy(Index).y - 3
   enemy(Index).bullet(i).cx = enemy(Index).x + 45
   enemy(Index).bullet(i).cy = enemy(Index).y - 34
End If
If dir = 19 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 26
   enemy(Index).bullet(i).ly = enemy(Index).y - 52
   enemy(Index).bullet(i).rx = enemy(Index).x + 12
   enemy(Index).bullet(i).ry = enemy(Index).y + 4
   enemy(Index).bullet(i).cx = enemy(Index).x + 46
   enemy(Index).bullet(i).cy = enemy(Index).y - 19
End If
If dir = 20 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 32
   enemy(Index).bullet(i).ly = enemy(Index).y - 43
   enemy(Index).bullet(i).rx = enemy(Index).x + 4
   enemy(Index).bullet(i).ry = enemy(Index).y + 8
   enemy(Index).bullet(i).cx = enemy(Index).x + 40
   enemy(Index).bullet(i).cy = enemy(Index).y - 8
End If
If dir = 21 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 53
   enemy(Index).bullet(i).ly = enemy(Index).y - 20
   enemy(Index).bullet(i).rx = enemy(Index).x + 10
   enemy(Index).bullet(i).ry = enemy(Index).y + 21
   enemy(Index).bullet(i).cx = enemy(Index).x + 45
   enemy(Index).bullet(i).cy = enemy(Index).y + 14
End If
If dir = 22 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 37
   enemy(Index).bullet(i).ly = enemy(Index).y - 29
   enemy(Index).bullet(i).rx = enemy(Index).x - 15
   enemy(Index).bullet(i).ry = enemy(Index).y + 1
   enemy(Index).bullet(i).cx = enemy(Index).x + 20
   enemy(Index).bullet(i).cy = enemy(Index).y + 10
End If
If dir = 23 Then
   enemy(Index).bullet(i).lx = enemy(Index).x + 32
   enemy(Index).bullet(i).ly = enemy(Index).y - 19
   enemy(Index).bullet(i).rx = enemy(Index).x - 24
   enemy(Index).bullet(i).ry = enemy(Index).y - 5
   enemy(Index).bullet(i).cx = enemy(Index).x + 8
   enemy(Index).bullet(i).cy = enemy(Index).y + 13
End If
If enemy(Index).type = 1 Then
   enemy(Index).bullet(i).llive = True
   enemy(Index).bullet(i).rlive = True
Else
   enemy(Index).bullet(i).clive = True
End If
enemy(Index).bullet(i).buldir = dir
enemy(Index).bullet(i).Speed = 50

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
enemy(Index).odestX = -1
enemy(Index).odestY = -1
destSurface.BltFast x - explowidth / 2, y - exploheight / 2, dds, CRect, DDBLTFAST_WAIT + DDBLTFAST_SRCCOLORKEY
explokeyframe = explokeyframe + 1
If (eexpl = 0 And explokeyframe > 17) Or (eexpl = 1 And explokeyframe > 15) Then
   If eexpl = 0 Then eexpl = 1 Else eexpl = 0
   explokeyframe = 0
   riposi Index
End If
End Sub

