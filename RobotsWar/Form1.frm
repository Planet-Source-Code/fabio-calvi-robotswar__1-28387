VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   FillColor       =   &H00FF00FF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerGame 
      Interval        =   1000
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer TMR_nuvens 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer TMR 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements DirectXEvent

'DirectDraw Object
Dim DX As DirectX7
Dim dd As DirectDraw7
'Dim di As IDirectInput2A
Public objDXEvent As DirectXEvent
Public objDI As DirectInput
Public objDIDev As DirectInputDevice

'Display capabilities
Dim DXCaps As DDSCAPS2

'Description of DirectDraw Surface
Dim DXDFront As DDSURFACEDESC2

'Front and Back Surface, for double Buffering
Dim DXSFront As DirectDrawSurface7
Dim DXSBack As DirectDrawSurface7

'Bitmap Resources sources, defined as DX-Surfaces
Dim DXSRobot1 As DirectDrawSurface7
Dim DXSRobot2 As DirectDrawSurface7
Dim DXSEnemy1 As DirectDrawSurface7
Dim DXSEnemy2 As DirectDrawSurface7
Dim DXSBase00 As DirectDrawSurface7
Dim DXSBase01 As DirectDrawSurface7
Dim DXSBase02 As DirectDrawSurface7
Dim DXSBackTiled As DirectDrawSurface7
Dim DXSBackTiled2 As DirectDrawSurface7
Dim DXSMapTiles() As DirectDrawSurface7
Dim DXSControlBar As DirectDrawSurface7
Dim DXSFont As DirectDrawSurface7
Dim DXSVLine As DirectDrawSurface7
Dim DXSHLine As DirectDrawSurface7
Dim DXSMouse As DirectDrawSurface7
Dim DXSSmMap As DirectDrawSurface7
Dim DXSSmScreen As DirectDrawSurface7
Dim DXSSmSpos As DirectDrawSurface7
Dim DXSSmEpos As DirectDrawSurface7
Dim DXSBlock As DirectDrawSurface7
Dim DXSNode1 As DirectDrawSurface7
Dim DXSNode2 As DirectDrawSurface7
Dim DXSLive As DirectDrawSurface7
Dim DXSLive1 As DirectDrawSurface7
Dim DXSExplo1 As DirectDrawSurface7
Dim DXSExplo2 As DirectDrawSurface7
Dim DXSBulletS As DirectDrawSurface7
Dim DXSBulletR As DirectDrawSurface7
Dim DXSSBullet As DirectDrawSurface7
Dim DXSRBullet As DirectDrawSurface7
Dim DXSWin As DirectDrawSurface7
Dim DXSLose As DirectDrawSurface7
Dim BufferMouse As DirectDrawSurface7

'Global Program Variables
Dim Message1 As MessageDisplay
Dim MAPWidth, MAPHeight
Dim MouseClick As Boolean
Dim MouseP As POINTAPI
Dim SMScreenX, SMScreenY
Dim Mousex As Integer, Mousey As Integer
Dim SelectSquare As Boolean
Dim SmallMapKliked As Boolean
Dim DemoStarted As Boolean
Dim DemoRunning As Boolean
Dim indind() As Integer, multisel As Boolean, sel As Boolean
Dim RS As RECT
Dim BufferMouseRS As RECT
Dim ddck As DDCOLORKEY
Dim DirectionX, DirectionY
Dim ScrollMapHorizontal As Boolean
Dim ScrollMapVertical As Boolean
Dim DeskTopHWND As Long
Dim DeskTopHDC As Long
Dim mouseHDC As Long
Dim lTemp As Long
Dim encount As Byte, mycount As Byte
Dim mb As Byte, explox As Integer, exploy As Integer
Dim gametimemin As Single, gametimesec As Single, appotime As Single
Dim enekills As Integer, roblosses As Integer, robkills As Integer, enelosses As Integer
Dim PicBits(1 To 338000) As Byte, PicInfo As BITMAP, Cnt As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
 
Private Sub Form_Click()
Message1.DisplayedTime = 1
MouseClick = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Dim i, i2, TiledWidth As Long, TiledHeight As Long
Dim file1, strtemp As String, strtemp1, strtemp2

'The Intro form with options is not implemented yet, sorry
'So, set this values as you prefer:
'Fx Sounds
'Max(100%) -> SoundVolume = 0
'High      -> SoundVolume = -500
'Medium    -> SoundVolume = -1000
'Low       -> SoundVolume = -2000
'Null(0%)  -> SoundVolume = -10000
SoundVolume = -750
Init_DSound
'Music
'Max(100%) -> MusicVolume = 100
'High      -> MusicVolume = 75
'Medium    -> MusicVolume = 50
'Low       -> MusicVolume = 25
'Null(0%)  -> MusicVolume = 0
MusicVolume = 65
Call modDXMusic.Initialize_Music
Call modDXMusic.Load_Music(0)
Call modDXMusic.SetMusic(MusicVolume)
Call modDXMusic.PlayMusic

Randomize
   
'Retrieve the desktop window handle
DeskTopHWND = GetDesktopWindow
   
'Find out it's HDC
DeskTopHDC = GetWindowDC(hWnd)
 
'Assign a free file number
file1 = FreeFile
    
'Open the map file
Open App.path & "\map1.txt" For Input As #file1
     
'Read width and height the first 2 lines
Input #file1, strtemp1, strtemp2
TiledWidth = Val(strtemp1)
TiledHeight = Val(strtemp2)
  
'Redimension of the MapTable array according to the file
ReDim TableMap(TiledWidth - 1, TiledHeight - 1)
     
'Counter for while loop
i = 0
     
'Load the map array with the file values
 Do While Not EOF(file1)
    Input #file1, strtemp
    For i2 = 0 To Len(strtemp) - 1
        TableMap(i2, i) = Val(Mid(strtemp, i2 + 1, 1))
    Next
    i = i + 1
Loop
  
'Set the real width and heigth according to the file and Tile
MAPWidth = i2 * TILEWIDTH
MAPHeight = i * TILEHEIGHT
    
'Set left and top of the small map to be diplayed
SMScreenX = 10
SMScreenY = 390
     
'This should be number of different tile in your map
'and you should only load tiles utilized in the map
ReDim DXSMapTiles(2)
     
'Set the message to be displayed when mouse is clicked
BufferMouseRS.Right = 23
BufferMouseRS.Bottom = 29
BufferMouseRS.Left = Mousex
BufferMouseRS.Top = Mousey
     
Message1.DisplayTime = 25
Message1.Position.Bottom = 15
Message1.Position.Top = 0
Message1.Position.Left = 0
    
mb = 0
ind = -1
ReDim indind(0)
    
'Building random nodes positions
nodes(1).x = Rnd * (2340 - 131) + 131
nodes(1).y = Rnd * (520 - 131) + 131
nodes(1).owner = 0
nodes(2).x = Rnd * (2340 - 131) + 131
nodes(2).y = Rnd * (1040 - 521) + 521
nodes(2).owner = 0
nodes(3).x = Rnd * (2340 - 131) + 131
nodes(3).y = Rnd * (1560 - 1041) + 1041
nodes(3).owner = 0
nodes(4).x = Rnd * (2340 - 131) + 131
nodes(4).y = Rnd * (2080 - 1561) + 1561
nodes(4).owner = 0
nodes(5).x = Rnd * (2340 - 131) + 131
nodes(5).y = Rnd * (2340 - 2081) + 2081
nodes(5).owner = 0
gametimemin = 0
gametimesec = 0
appotime = 0
en = 0
my = 0
enscore = 0
myscore = 0
encount = 0
mycount = 0
flgup = False
enekills = 0
robkills = 0
enelosses = 0
roblosses = 0
End Sub

Sub DDInitiation()
'Initialize DirectX-Object; set display mode
Set DX = New DirectX7
Set dd = DX.DirectDrawCreate("")
     
'dx DirectDrawCreate ByVal 0&, dx, Nothing
dd.SetCooperativeLevel Me.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
dd.SetDisplayMode 640, 480, 32, 0, 0
  
'Initialize front buffer description
With DXDFront
'     .dwSize = Len(DXDFront)
     .lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
     .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX 'Or DDSCAPS_SYSTEMMEMORY
     .lBackBufferCount = 1
End With
     
'Create front buffer from structure
Set DXSFront = dd.CreateSurface(DXDFront)
   
'Create back buffer from front buffer
DXCaps.lCaps = DDSCAPS_BACKBUFFER
   
'Attach both surface together
 Set DXSBack = DXSFront.GetAttachedSurface(DXCaps)

'Initialize our cursor
' procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SysMenuProc)

Mousex = 320
Mousey = 240

'g_Sensitivity = 1.2
  
'Create DirectInput and set up the mouse
' Set objDI = dx.DirectInputCreate
' Set objDIDev = objDI.CreateDevice("guid_SysMouse")
' Call objDIDev.SetCommonDataFormat(DIFORMAT_MOUSE)
' Call objDIDev.SetCooperativeLevel(hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE)
 
'Set the buffer size
' Dim diProp As DIPROPLONG
' diProp.lHow = DIPH_DEVICE
' diProp.lObj = 0
' diProp.lData = BufferSize
' diProp.lSize = Len(diProp)
' Call objDIDev.SetProperty("DIPROP_BUFFERSIZE", diProp)

'Ask for notifications
' EventHandle = dx.CreateEvent(Me)
' Call objDIDev.SetEventNotification(EventHandle)

'Acquire the mouse
' AcquireMouse
     
End Sub

Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)
Dim diDeviceData(1 To BufferSize) As DIDEVICEOBJECTDATA
Dim NumItems As Integer, i As Integer, x As Integer
Static OldSequence As Long
  
'Get data
On Error GoTo INPUTLOST
NumItems = objDIDev.GetDeviceData(diDeviceData, 0)
On Error GoTo 0
  
'If Button = 1 Then SelectSquare = True
  
'Process data
For i = 1 To NumItems
    Select Case diDeviceData(i).lOfs
       Case DIMOFS_X
            Mousex = Mousex + diDeviceData(i).lData * g_Sensitivity
'            If OldSequence <> diDeviceData(i).lSequence Then
'               OldSequence = diDeviceData(i).lSequence
'            Else
'               OldSequence = 0
'            End If
         
       Case DIMOFS_Y
            Mousey = Mousey + diDeviceData(i).lData * g_Sensitivity
'            If OldSequence <> diDeviceData(i).lSequence Then
'               OldSequence = diDeviceData(i).lSequence
'            Else
'               OldSequence = 0
'            End If
      
       Case DIMOFS_BUTTON0
            If diDeviceData(i).lData And &H80 Then
               'Keep record for Line function
               CurrentX = Mousex
               CurrentY = Mousey
               ClickOrigineX = CurrentX
               ClickOrigineY = CurrentY
               'Define square for the little map if clicked on
               SelectSquare = True
               If CurrentX > 10 And CurrentX < 88 And CurrentY > 390 And CurrentY < 468 Then SmallMapKliked = True
            Else
               SelectSquare = False
               MouseClick = True
'              Drawing = False
            End If
         
       Case DIMOFS_BUTTON1
            If diDeviceData(i).lData = 0 Then
'               Popup
            End If
   End Select
Next i
 
'Green select square
Select Case Mousex
   Case Is <= 0: Mousex = 0: DirectionX = -1: ScrollMapHorizontal = True
   Case Is >= 639: Mousex = 639: DirectionX = 1: ScrollMapHorizontal = True
   Case Else: ScrollMapHorizontal = False
End Select
Select Case Mousey
   Case Is <= 0: Mousey = 0: DirectionY = -1: ScrollMapVertical = True
   Case Is >= 479: Mousey = 479: DirectionY = 1: ScrollMapVertical = True
   Case Else: ScrollMapVertical = False
End Select
 
Exit Sub
  
INPUTLOST:
'Windows stole the mouse from us. DIERR_INPUTLOST is raised if the user switched to
'another app, but DIERR_NOTACQUIRED is raised if the Windows key was pressed.
If (Err.Number = DIERR_INPUTLOST) Or (Err.Number = DIERR_NOTACQUIRED) Then
    SetSystemCursor
    Exit Sub
End If
    
End Sub
Sub AcquireMouse()

  Dim CursorPoint As POINTAPI
  
  'Move private cursor to system cursor; get position before Windows loses cursor
  Call GetCursorPos(CursorPoint)
  Call ScreenToClient(hWnd, CursorPoint)
  
  On Error GoTo CANNOTACQUIRE
  objDIDev.Acquire
  Mousex = CursorPoint.x
  Mousey = CursorPoint.y
  
  On Error GoTo 0
  Exit Sub

CANNOTACQUIRE:
  Exit Sub
End Sub

Public Sub SetSystemCursor()
Dim point As POINTAPI
 
'Get the system cursor into the same position as the private cursor and stop drawing
point.x = Mousex
point.y = Mousey
ClientToScreen hWnd, point
SetCursorPos point.x, point.y

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'Quit the game
Select Case KeyCode
    Case 27
        DemoRunning = False
End Select

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'Save the mouse click origin
CurrentX = x
CurrentY = y
ClickOrigineX = CurrentX
ClickOrigineY = CurrentY
SelectSquare = True

'Define square for the little map if clicked on
If CurrentX > 10 And CurrentX < 88 And CurrentY > 390 And CurrentY < 468 Then
   SmallMapKliked = True
   ScrollMapHorizontal = True
   ScrollMapVertical = True
   SelectSquare = False
Else
   'Left mouse button pressed
   If Button = 2 Then
      mb = 2
   Else
     'Right mouse button pressed
      mb = 1
      ind = -1
      If Not Shift = 0 Then multisel = True Else multisel = False
      sel = False
      ReDim indind(0)
   End If
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer

If Not SmallMapKliked Then
   Select Case x
      Case 0: Mousex = 0: DirectionX = -1: ScrollMapHorizontal = True
      Case 639: Mousex = 639: DirectionX = 1: ScrollMapHorizontal = True
      Case Else: ScrollMapHorizontal = False
   End Select
   Select Case y
      Case 0: Mousey = 0: DirectionY = -1: ScrollMapVertical = True
      Case 479: Mousey = 479: DirectionY = 1: ScrollMapVertical = True
      Case Else: ScrollMapVertical = False
   End Select
End If
 
'If CurrentX > 10 And CurrentX < MAPWidth / 100 + 10 And CurrentY > 390 And CurrentY < MAPHeight / 100 + 390 Then SmallMapKliked = True
Mousex = x
Mousey = y

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Byte, i2 As Byte, m As Byte

SelectSquare = False
SmallMapKliked = False

If CurrentX > 10 And CurrentX < 88 And CurrentY > 390 And CurrentY < 468 Then
   MouseClick = True
Else
   'Left mouse button allows selection
   If mb = 1 Then
      ind = -1
      'Single robot selected
      If multisel = False Then
         For i = 0 To 5
             Robot(i).GetPosition xx, yy
             If x + ViewPortX >= Int(xx - 36) And x + ViewPortX < Int(xx + 36) And y + ViewPortY >= Int(yy - 72) And y + ViewPortY < Int(yy) Then ind = i
         Next
      Else
         'Multi robots selected
         For i = 0 To 5
             Robot(i).GetPosition xx, yy
             If (ClickOrigineX + ViewPortX <= Int(xx - 36) And Mousex + ViewPortX > Int(xx + 36) And ClickOrigineY + ViewPortY < Int(yy + 36) And Mousey + ViewPortY >= Int(yy - 36)) Or _
                (ClickOrigineX + ViewPortX >= Int(xx - 36) And Mousex + ViewPortX < Int(xx + 36) And ClickOrigineY + ViewPortY < Int(yy + 36) And Mousey + ViewPortY >= Int(yy - 36)) Or _
                (ClickOrigineX + ViewPortX <= Int(xx - 36) And Mousex + ViewPortX > Int(xx + 36) And ClickOrigineY + ViewPortY > Int(yy + 36) And Mousey + ViewPortY <= Int(yy - 36)) Or _
                (ClickOrigineX + ViewPortX >= Int(xx - 36) And Mousex + ViewPortX < Int(xx + 36) And ClickOrigineY + ViewPortY > Int(yy + 36) And Mousey + ViewPortY <= Int(yy - 36)) Then
                  ReDim Preserve indind(UBound(indind) + 1)
                  indind(UBound(indind)) = i
                  sel = True
                  multisel = False
             End If
         Next
      End If
   End If
   
   'Right mouse button allows destination
   If mb = 2 Then
      'Multi robots
      If sel = True Then
         For i2 = 1 To UBound(indind)
             Robot(indind(i2)).GetPosition xx, yy
             Robot(indind(i2)).SetDestination ViewPortX + Mousex, ViewPortY + Mousey
             Robot(indind(i2)).Getdir xx, yy, ViewPortX + Mousex, ViewPortY + Mousey
             Robot(indind(i2)).Turn dir
             robots(indind(i2)).odestX = ViewPortX + Mousex
             robots(indind(i2)).odestY = ViewPortY + Mousey
         Next i2
         multisel = False
      'Single robot
      ElseIf ind >= 0 Then
         Robot(ind).GetPosition xx, yy
         Robot(ind).SetDestination ViewPortX + Mousex, ViewPortY + Mousey
         Robot(ind).Getdir xx, yy, ViewPortX + Mousex, ViewPortY + Mousey
         Robot(ind).Turn dir
         robots(ind).odestX = ViewPortX + Mousex
         robots(ind).odestY = ViewPortY + Mousey
      End If
      m = Int(3 * Rnd) + 1
      Select Case m
        Case 1
           Message1.MessageText = "on your commands"
        Case 2
            Message1.MessageText = "yes commander"
        Case 3
           Message1.MessageText = "at your orders"
       End Select
   End If

End If
End Sub

Private Sub Form_Paint()
Dim i As Integer

If DemoStarted Then Exit Sub
DemoStarted = True
    
'DirectDraw initialization
DDInitiation
    
'Initialize surfaces
Set DXSRobot1 = LoadBitmapIntoDXS(dd, App.path + "\robot1.bmp", 2, 2, 1)
Set DXSRobot2 = LoadBitmapIntoDXS(dd, App.path + "\robot.bmp", 2, 2, 1)
Set DXSEnemy1 = LoadBitmapIntoDXS(dd, App.path + "\robot0.bmp", 2, 2, 1)
Set DXSEnemy2 = LoadBitmapIntoDXS(dd, App.path + "\robot3.bmp", 2, 2, 1)
Set DXSBackTiled = LoadBitmapIntoDXS(dd, App.path + "\tile1.bmp", 6, 5, 1)
Set DXSMapTiles(0) = LoadBitmapIntoDXS(dd, App.path + "\tile1.bmp", 1, 1, 1)
Set DXSMapTiles(1) = LoadBitmapIntoDXS(dd, App.path + "\tile2.bmp", 1, 1, 1)
Set DXSBase00 = LoadBitmapIntoDXS(dd, App.path + "\Base0.bmp", 1, 1, 1)
Set DXSBase01 = LoadBitmapIntoDXS(dd, App.path + "\Base1.bmp", 1, 1, 1)
Set DXSBase02 = LoadBitmapIntoDXS(dd, App.path + "\Base2.bmp", 1, 1, 1)
Set DXSControlBar = LoadBitmapIntoDXS(dd, App.path + "\barra.bmp", 1, 1, 1)
Set DXSFont = LoadBitmapIntoDXS(dd, App.path + "\myfont.bmp", 1, 1, 1)
Set DXSVLine = LoadBitmapIntoDXS(dd, App.path + "\Vline.bmp", 640, 1, 1)
Set DXSHLine = LoadBitmapIntoDXS(dd, App.path + "\Vline.bmp", 1, 480, 1)
Set DXSMouse = LoadBitmapIntoDXS(dd, App.path + "\mouse.bmp", 1, 1, 1)
Set BufferMouse = LoadBitmapIntoDXS(dd, App.path + "\mouse.bmp", 1, 1, 1)
Set DXSSmMap = LoadBitmapIntoDXS(dd, App.path + "\SmallMap.bmp", 1, 1, 1)
Set DXSSmScreen = LoadBitmapIntoDXS(dd, App.path + "\smscreen.bmp", 1, 1, 1)
Set DXSSmSpos = LoadBitmapIntoDXS(dd, App.path + "\smsspos.bmp", 1, 1, 1)
Set DXSSmEpos = LoadBitmapIntoDXS(dd, App.path + "\smsepos.bmp", 1, 1, 1)
Set DXSBlock = LoadBitmapIntoDXS(dd, App.path + "\smblock.bmp", 1, 1, 1)
Set DXSNode1 = LoadBitmapIntoDXS(dd, App.path + "\smnode1.bmp", 1, 1, 1)
Set DXSNode2 = LoadBitmapIntoDXS(dd, App.path + "\smnode2.bmp", 1, 1, 1)
Set DXSExplo1 = LoadBitmapIntoDXS(dd, App.path + "\explosion1.bmp", 2, 2, 1)
Set DXSExplo2 = LoadBitmapIntoDXS(dd, App.path + "\explosion2.bmp", 2, 2, 1)
Set DXSLive = LoadBitmapIntoDXS(dd, App.path + "\live.bmp", 1, 1, 1)
Set DXSLive1 = LoadBitmapIntoDXS(dd, App.path + "\live1.bmp", 1, 1, 1)
Set DXSBulletS = LoadBitmapIntoDXS(dd, App.path + "\bullets.bmp", 1, 1, 1)
Set DXSBulletR = LoadBitmapIntoDXS(dd, App.path + "\bulletr.bmp", 1, 1, 1)
Set DXSSBullet = LoadBitmapIntoDXS(dd, App.path + "\Sbullet.bmp", 1, 1, 1)
Set DXSRBullet = LoadBitmapIntoDXS(dd, App.path + "\Rbullet.bmp", 1, 1, 1)
Set DXSWin = LoadBitmapIntoDXS(dd, App.path + "\win.bmp", 3, 5, 1)
Set DXSLose = LoadBitmapIntoDXS(dd, App.path + "\lose.bmp", 3, 5, 1)
    
ShowCursor False
ScrollMapHorizontal = False
ScrollMapVertical = False
ViewPortY = 0
ViewPortX = 0
    
ReDim robots(5)
ReDim Robot(5)
    
Robot(0).SetPosition 180, 180
robots(0).type = 1
Robot(1).SetPosition 280, 280
robots(1).type = 1
Robot(2).SetPosition 380, 380
robots(2).type = 1
Robot(3).SetPosition 180, 380
robots(3).type = 2
Robot(4).SetPosition 280, 180
robots(4).type = 2
Robot(5).SetPosition 380, 280
robots(5).type = 2
For i = 0 To 5
    Robot(i).Init i
    robots(i).status = True
Next
   
ReDim enemy(5)
ReDim Enemies(5)

Enemies(0).SetPosition 2280, 2280
enemy(0).x = 2280: enemy(0).y = 2280
enemy(0).type = 1
Enemies(1).SetPosition 2150, 2150
enemy(1).x = 2150: enemy(1).y = 2150
enemy(1).type = 1
Enemies(2).SetPosition 1990, 1990
enemy(2).x = 1990: enemy(2).y = 1990
enemy(3).type = 1
Enemies(3).SetPosition 1990, 2280
enemy(3).x = 1990: enemy(3).y = 2280
enemy(4).type = 2
Enemies(4).SetPosition 2150, 1990
enemy(4).x = 2150: enemy(4).y = 1990
enemy(4).type = 2
Enemies(5).SetPosition 2280, 2150
enemy(5).x = 2280: enemy(5).y = 2150
enemy(5).type = 2
For i = 0 To 5
    Enemies(i).Init i
    enemy(i).status = True
Next

bullet_delay = 250

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
    
'Flip from DX-Surface to standard GDI
dd.FlipToGDISurface

' Restore old resolution and depth
dd.RestoreDisplayMode
    
' Return control to windows
dd.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL

'Terminate Music
modDXMusic.EndMusic
    
'Terminate DirectSound
Call Terminate_DSound

'Clear all DX Objects
Set DXSRobot1 = Nothing
Set DXSRobot2 = Nothing
For i = 0 To 5
    Robot(i).Destroy
Next
Set DXSEnemy1 = Nothing
Set DXSEnemy2 = Nothing
For i = 0 To 5
    Enemies(i).Destroy
Next
Set DXSBackTiled = Nothing
Set DXSMapTiles(0) = Nothing
Set DXSMapTiles(1) = Nothing
Set DXSBase00 = Nothing
Set DXSBase01 = Nothing
Set DXSBase02 = Nothing
Set DXSControlBar = Nothing
Set DXSFont = Nothing
Set DXSVLine = Nothing
Set DXSHLine = Nothing
Set DXSMouse = Nothing
Set DXSSmMap = Nothing
Set DXSSmScreen = Nothing
Set DXSSmSpos = Nothing
Set DXSSmEpos = Nothing
Set DXSBack = Nothing
Set DXSFront = Nothing
Set DXSBlock = Nothing
Set DXSNode1 = Nothing
Set DXSNode2 = Nothing
Set DXSExplo1 = Nothing
Set DXSExplo2 = Nothing
Set DXSLive = Nothing
Set DXSLive1 = Nothing
Set DXSBulletS = Nothing
Set DXSBulletR = Nothing
Set DXSSBullet = Nothing
Set DXSRBullet = Nothing
Set DXSWin = Nothing
Set DXSLose = Nothing
Set dd = Nothing
'If procOld <> 0 Then Call SetWindowLong(hWnd, GWL_WNDPROC, procOld)
'If EventHandle <> 0 Then dx.DestroyEvent EventHandle
 Set DX = Nothing
 
 ShowCursor 1
'lTemp = ReleaseDC(DeskTopHWND, DeskTopHDC)

End Sub

Private Sub TimerGame_Timer()
Dim i As Integer
gametimesec = gametimesec + 1
If gametimesec = appotime + 10 Then
   appotime = gametimesec
   encount = 0
   mycount = 0
   For i = 1 To 5
       If nodes(i).owner = 1 Then mycount = mycount + 1
       If nodes(i).owner = 2 Then encount = encount + 1
   Next

   If mycount = 1 Then myscore = myscore + 4
   If mycount = 2 Then myscore = myscore + 12
   If mycount = 3 Then myscore = myscore + 24
   If mycount = 4 Then myscore = myscore + 40
   If mycount = 5 Then myscore = myscore + 60
   If encount = 1 Then enscore = enscore + 4
   If encount = 2 Then enscore = enscore + 12
   If encount = 3 Then enscore = enscore + 24
   If encount = 4 Then enscore = enscore + 40
   If encount = 5 Then enscore = enscore + 60
End If
If myscore > enscore Then flgup = True Else flgup = False

If gametimesec > 59 Then
   gametimesec = 0
   appotime = 0
   gametimemin = gametimemin + 1
End If
If gametimemin = 10 And gametimesec = 1 Then
   RS.Right = 200
   RS.Bottom = 100
   If myscore > enscore Then
      DXSBack.BltFast 220, 170, DXSWin, RS, DDBLTFAST_SRCCOLORKEY
   Else
      DXSBack.BltFast 220, 170, DXSLose, RS, DDBLTFAST_SRCCOLORKEY
   End If
   DemoRunning = False
End If

End Sub

Private Sub TMR_Timer()
 'Wait until all the objects are created from the paint event
 If Not DemoStarted Then Exit Sub
 Dim i As Integer, i2 As Integer
 Dim divtempX, divtempY
 Dim dxdc As Long, divrestY As Integer, divrestX As Integer
 Dim ScrollMapHorizontal2 As Boolean
 
 TMR.Enabled = False
 DemoRunning = True
 
 While DemoRunning

    'Scroll map check for event
    If ScrollMapHorizontal And Not SelectSquare Then ViewPortX = GetNewViewPortPosition(DirectionX, ViewPortX, 640, MAPWidth)
    If ScrollMapVertical And Not SelectSquare Then ViewPortY = GetNewViewPortPosition(DirectionY, ViewPortY, 480, MAPHeight)
  
    'Scrolling for small map
    If SmallMapKliked Then
      If Mousex > 10 And Mousex < 88 And Mousey > 390 And Mousey < 468 Then
        ViewPortX = Int(((Mousex - 10) * 33) / SCROLLSPEED) * SCROLLSPEED
        ViewPortY = Int(((Mousey - 390) * 33) / SCROLLSPEED) * SCROLLSPEED
        If Mousex > 60 And Mousex < 88 Then ViewPortX = Int(((Mousex - 28) * 33) / SCROLLSPEED) * SCROLLSPEED
        If Mousey > 458 And Mousey < 468 Then ViewPortY = Int(((Mousey - 401) * 33) / SCROLLSPEED) * SCROLLSPEED
      Else
        SmallMapKliked = False
      End If
    End If

    'Paint map tiles
    divrestX = ViewPortX Mod 130
    divrestY = ViewPortY Mod 130
    RS.Top = divrestY
    RS.Left = divrestX
    RS.Right = 130
    RS.Bottom = 130
    divtempX = 0
    divtempY = 0
    For i = 0 To 4
       For i2 = 0 To 5
         DXSBackTiled.BltFast (i2 * 130 - divtempX), (i * 130 - divtempY), DXSMapTiles(TableMap(Int(ViewPortX / 130) + i2, Int(ViewPortY / 130) + i)), RS, DDBLTFAST_SRCCOLORKEY
         RS.Left = 0
         divtempX = divrestX
       Next
       RS.Left = divrestX
       RS.Top = 0
       divtempY = divrestY
       divtempX = 0
    Next

    'Final painting before fliping
    RS.Top = 0
    RS.Left = 0
    RS.Right = 640
    RS.Bottom = 480
    DXSBack.BltFast 0, 0, DXSBackTiled, RS, DDBLTFAST_SRCCOLORKEY
    
    'Painting Nodes
     For i = 1 To 5
         RS.Top = 0
         RS.Left = 0
         RS.Bottom = 130
         RS.Right = 130
         If nodes(i).owner = 1 Then
            DXSBack.BltFast -ViewPortX + nodes(i).x, -ViewPortY + nodes(i).y, DXSBase01, RS, DDBLTFAST_SRCCOLORKEY
         ElseIf nodes(i).owner = 2 Then DXSBack.BltFast -ViewPortX + nodes(i).x, -ViewPortY + nodes(i).y, DXSBase02, RS, DDBLTFAST_SRCCOLORKEY
         Else
            DXSBack.BltFast -ViewPortX + nodes(i).x, -ViewPortY + nodes(i).y, DXSBase00, RS, DDBLTFAST_SRCCOLORKEY
         End If
     Next
   
    'If mouse clicked paint message
    If MouseClick Then
        Set Message1.MessageSurface = SetDisplayMessage(dd, Message1.MessageText)
        Message1.Position.Right = 15 * Len(Message1.MessageText)
        DXSBack.BltFast ((640 + Len(Message1.MessageText)) / 2) - 97, 400, Message1.MessageSurface, Message1.Position, DDBLTFAST_SRCCOLORKEY
        Set Message1.MessageSurface = Nothing
        If Message1.DisplayedTime = Message1.DisplayTime Then MouseClick = False
        Message1.DisplayedTime = Message1.DisplayedTime + 1
    End If
    
    'Black bar at the bottom
     RS.Bottom = 115
     RS.Right = 640
     DXSBack.BltFast 0, 365, DXSControlBar, RS, DDBLTFAST_SRCCOLORKEY
  
    'Paint small map
     RS.Right = 78
     RS.Bottom = 78
     DXSBack.BltFast 10, 390, DXSSmMap, RS, DDBLTFAST_SRCCOLORKEY
    
     RS.Right = 4
     RS.Bottom = 4
     For i = 0 To 19
       For i2 = 0 To 19
         If TableMap(i, i2) = 1 Then DXSBack.BltFast SMScreenX + Int((i * 130) / 33), SMScreenY + Int((i2 * 130) / 33), DXSBlock, RS, DDBLTFAST_SRCCOLORKEY
       Next
     Next
     
     RS.Right = 4
     RS.Bottom = 4
     For i = 1 To 5
         If nodes(i).owner = 1 Then
            DXSBack.BltFast SMScreenX + Int(nodes(i).x / 33), SMScreenY + Int(nodes(i).y / 33), DXSNode1, RS, DDBLTFAST_SRCCOLORKEY
         ElseIf nodes(i).owner = 2 Then DXSBack.BltFast SMScreenX + Int(nodes(i).x / 33), SMScreenY + Int(nodes(i).y / 33), DXSNode2, RS, DDBLTFAST_SRCCOLORKEY
         Else
            DXSBack.BltFast SMScreenX + Int(nodes(i).x / 33), SMScreenY + Int(nodes(i).y / 33), DXSBlock, RS, DDBLTFAST_SRCCOLORKEY
         End If
     Next
    
    'Paint small area on the small map
     RS.Right = 20
     RS.Bottom = 14
     DXSBack.BltFast SMScreenX + Int(ViewPortX / 33), SMScreenY + Int(ViewPortY / 33), DXSSmScreen, RS, DDBLTFAST_SRCCOLORKEY
    
    'Updating and painting Robots
     For i = 0 To 5
        Robot(i).UpdateAI
        Robot(i).UpdateMove i
        Robot(i).UpdateAnimation
        If robots(i).type = 1 Then
           Robot(i).Draw DXSRobot1, DXSBack, i
        Else
           Robot(i).Draw DXSRobot2, DXSBack, i
        End If
        If robots(i).status = False Then
           If sexpl = 0 Then Robot(i).Explosion DXSExplo2, DXSBack, Int(-ViewPortX + robots(i).x), Int(-ViewPortY + robots(i).y), i
           If sexpl = 1 Then Robot(i).Explosion DXSExplo1, DXSBack, Int(-ViewPortX + robots(i).x), Int(-ViewPortY + robots(i).y), i
        End If
        'Drawing selection box
        'Multi sel
        If robots(i).status = True Then
           If sel = True Then
              Robot(indind(i + 1)).GetPosition xx, yy
              If robots(indind(i + 1)).type = 2 Then
                 RS.Bottom = 1
                 RS.Right = 73
                 DXSBack.BltFast -ViewPortX + xx - 36, -ViewPortY + yy - 65, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
                 DXSBack.BltFast -ViewPortX + xx - 36, -ViewPortY + yy + 10, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
                 RS.Right = 1
                 RS.Bottom = 75
                 DXSBack.BltFast -ViewPortX + xx - 36, -ViewPortY + yy - 65, DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
                 DXSBack.BltFast -ViewPortX + xx + 36, -ViewPortY + yy - 65, DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
              Else
                 RS.Bottom = 1
                 RS.Right = 74
                 DXSBack.BltFast -ViewPortX + xx - 37, -ViewPortY + yy - 75, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
                 DXSBack.BltFast -ViewPortX + xx - 37, -ViewPortY + yy + 10, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
                 RS.Right = 1
                 RS.Bottom = 85
                 DXSBack.BltFast -ViewPortX + xx - 37, -ViewPortY + yy - 75, DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
                 DXSBack.BltFast -ViewPortX + xx + 37, -ViewPortY + yy - 75, DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
              End If
              'Single sel
           ElseIf i = ind Then
              If robots(indind(ind)).type = 2 Then
                 RS.Bottom = 1
                 RS.Right = 73
                 DXSBack.BltFast -ViewPortX + soldx - 36, -ViewPortY + soldy - 65, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
                 DXSBack.BltFast -ViewPortX + soldx - 36, -ViewPortY + soldy + 10, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
                 RS.Right = 1
                 RS.Bottom = 75
                 DXSBack.BltFast -ViewPortX + soldx - 36, -ViewPortY + soldy - 65, DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
                 DXSBack.BltFast -ViewPortX + soldx + 36, -ViewPortY + soldy - 65, DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
              Else
                 RS.Bottom = 1
                 RS.Right = 74
                 DXSBack.BltFast -ViewPortX + soldx - 37, -ViewPortY + soldy - 75, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
                 DXSBack.BltFast -ViewPortX + soldx - 37, -ViewPortY + soldy + 10, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
                 RS.Right = 1
                 RS.Bottom = 85
                 DXSBack.BltFast -ViewPortX + soldx - 37, -ViewPortY + soldy - 75, DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
                 DXSBack.BltFast -ViewPortX + soldx + 37, -ViewPortY + soldy - 75, DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
             End If
           End If
        
           'Drawing robots life box
           RS.Bottom = 1
           RS.Right = 73
           DXSBack.BltFast -ViewPortX + robots(i).x - 36, -ViewPortY + robots(i).y + 15, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
           DXSBack.BltFast -ViewPortX + robots(i).x - 36, -ViewPortY + robots(i).y + 20, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
           RS.Right = 1
           RS.Bottom = 5
           DXSBack.BltFast -ViewPortX + robots(i).x - 36, -ViewPortY + robots(i).y + 15, DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
           DXSBack.BltFast -ViewPortX + robots(i).x + 36, -ViewPortY + robots(i).y + 15, DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
           RS.Right = 71 - robots(i).hit
           RS.Bottom = 4
           If RS.Right <= 0 Then
              If robots(i).type = 1 Then Call PlaySound(dsXpl3)
              If robots(i).type = 2 Then Call PlaySound(dsXpl4)
              Robot(i).StopMovement
              robots(i).status = False
              roblosses = roblosses + 1
              robkills = robkills + 1
           End If
           DXSBack.BltFast -ViewPortX + robots(i).x - 35, -ViewPortY + robots(i).y + 16, DXSLive, RS, DDBLTFAST_SRCCOLORKEY
           DXSBack.BltFast -ViewPortX + robots(i).x - 35, -ViewPortY + robots(i).y + 16, DXSLive, RS, DDBLTFAST_SRCCOLORKEY
         
           'Painting robot position on the small map
           RS.Right = 2
           RS.Bottom = 2
           DXSBack.BltFast SMScreenX + Int(soldx / 33), SMScreenY + Int(soldy / 33), DXSSmSpos, RS, DDBLTFAST_SRCCOLORKEY
        End If

     Next
     
    'Updating and painting Enemy Robots
     For i = 0 To 5
        Enemies(i).UpdateAI
        Enemies(i).UpdateMove i
        Enemies(i).UpdateAnimation
        If enemy(i).type = 1 Then
           Enemies(i).Draw DXSEnemy1, DXSBack, i
        Else
           Enemies(i).Draw DXSEnemy2, DXSBack, i
        End If
        If enemy(i).status = False Then
           If eexpl = 0 Then Enemies(i).Explosion DXSExplo2, DXSBack, Int(-ViewPortX + enemy(i).x), Int(-ViewPortY + enemy(i).y), i
           If eexpl = 1 Then Enemies(i).Explosion DXSExplo1, DXSBack, Int(-ViewPortX + enemy(i).x), Int(-ViewPortY + enemy(i).y), i
        End If
        If enemy(i).status = True Then
           RS.Bottom = 1
           RS.Right = 73
           DXSBack.BltFast -ViewPortX + eoldx - 36, -ViewPortY + eoldy + 10, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
           DXSBack.BltFast -ViewPortX + eoldx - 36, -ViewPortY + eoldy + 15, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
           RS.Right = 1
           RS.Bottom = 5
           DXSBack.BltFast -ViewPortX + eoldx - 36, -ViewPortY + eoldy + 10, DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
           DXSBack.BltFast -ViewPortX + eoldx + 36, -ViewPortY + eoldy + 10, DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
           RS.Right = 71 - enemy(i).hit
           RS.Bottom = 4
           If RS.Right <= 0 Then
              If enemy(i).type = 1 Then Call PlaySound(dsXpl1)
              If enemy(i).type = 2 Then Call PlaySound(dsXpl2)
              Enemies(i).StopMovement
              enemy(i).status = False
              enelosses = enelosses + 1
              enekills = enekills + 1
           End If
           DXSBack.BltFast -ViewPortX + eoldx - 35, -ViewPortY + eoldy + 11, DXSLive1, RS, DDBLTFAST_SRCCOLORKEY
           DXSBack.BltFast -ViewPortX + eoldx - 35, -ViewPortY + eoldy + 11, DXSLive1, RS, DDBLTFAST_SRCCOLORKEY
        
           'Painting enemy robots on the small map
           RS.Right = 2
           RS.Bottom = 2
           DXSBack.BltFast SMScreenX + Int(eoldx / 33), SMScreenY + Int(eoldy / 33), DXSSmEpos, RS, DDBLTFAST_SRCCOLORKEY
        End If
     Next
    
    DXSBack.SetForeColor vbWhite
    DXSBack.DrawText 220, 432, myscore, False
    DXSBack.DrawText 572, 432, enscore, False
    DXSBack.DrawText 326, 450, CStr(Format(gametimemin, "00")) & ":" & CStr(Format(gametimesec, "00")), False
    DXSBack.DrawText 220, 448, enekills, False
    DXSBack.DrawText 220, 462, roblosses, False
    DXSBack.DrawText 572, 448, robkills, False
    DXSBack.DrawText 572, 462, enelosses, False

    For i = 0 To 5
        If enemy(i).type = 1 Then
           ebullets DXSBulletS, DXSBack, i
        Else
           ebullets DXSSBullet, DXSBack, i
        End If
        If robots(i).type = 1 Then
           rbullets DXSBulletR, DXSBack, i
        Else
           rbullets DXSRBullet, DXSBack, i
        End If
    Next
    
    'Loop Music
    Call modDXMusic.LoopMusic

    DoEvents
    
    'If select square then paint the square
    If SelectSquare And Not SmallMapKliked And multisel = True Then
       RS.Bottom = 1
       RS.Right = Abs(Mousex - ClickOrigineX)
       DXSBack.BltFast IIf(Mousex < ClickOrigineX, Mousex, ClickOrigineX), ClickOrigineY, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
       DXSBack.BltFast IIf(Mousex < ClickOrigineX, Mousex, ClickOrigineX), Mousey, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
       RS.Right = 1
       RS.Bottom = Abs(Mousey - ClickOrigineY) + 1
       DXSBack.BltFast ClickOrigineX, IIf(Mousey < ClickOrigineY, Mousey, ClickOrigineY), DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
       DXSBack.BltFast Mousex, IIf(Mousey < ClickOrigineY, Mousey, ClickOrigineY), DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
    End If
    
    'Painting the mouse
    RS.Left = 0
    RS.Top = 0
    RS.Bottom = 29
    RS.Right = 23
    If Mousey > 450 Then RS.Bottom = 480 - Mousey
    If Mousex > 616 Then RS.Right = 640 - Mousex
    DXSBack.BltFast Mousex, Mousey, DXSMouse, RS, DDBLTFAST_SRCCOLORKEY
       
    'Fliping buffers chain
    On Error Resume Next
    DXSFront.Flip DXSBack, 0
    If Err.Number = DDERR_SURFACELOST Then DXSFront.restore
        
Wend

End Sub

Private Function SetDisplayMessage(DXObject As DirectDraw7, ByVal mess As String) As DirectDrawSurface7
'This function to create and return a text message with the format of a surface with a special font

Dim DXSMessTemp As DirectDrawSurface7
Dim TempDXD As DDSURFACEDESC2
Dim ddck As DDCOLORKEY
Dim RStemp As RECT
Dim i
ddck.low = 0
ddck.high = 0

'Set the surface values
With TempDXD
        '.dwSize = Len(TempDXD)
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lWidth = 15 * Len(mess)
        .lHeight = 15
End With
    
'Create DX surface
Set DXSMessTemp = DXObject.CreateSurface(TempDXD)
RStemp.Top = 0
RStemp.Bottom = 15

'Convert message to bitmap onto surface
For i = 1 To Len(mess)
    RStemp.Left = (Asc(Mid(mess, i, 1)) - 97) * 15
    RStemp.Right = RStemp.Left + 15
    DXSMessTemp.BltFast (i - 1) * 15, 0, DXSFont, RStemp, DDBLTFAST_SRCCOLORKEY
Next
DXSMessTemp.SetColorKey DDCKEY_SRCBLT, ddck
Set SetDisplayMessage = DXSMessTemp
End Function

