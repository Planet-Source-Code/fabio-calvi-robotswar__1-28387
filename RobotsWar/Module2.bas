Attribute VB_Name = "Module1"
Option Explicit
Public Const pi As Double = 3.14159265358979
'Public Const Radians As Double = (2 * pi) / 360
'Public Const MOUSEEVENTF_LEFTDOWN = &H2
'Public Const MOUSEEVENTF_LEFTUP = &H4
'Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
'Public Const MOUSEEVENTF_MIDDLEUP = &H40
'Public Const MOUSEEVENTF_RIGHTDOWN = &H8
'Public Const MOUSEEVENTF_RIGHTUP = &H10
'Public Const MOUSEEVENTF_MOVE = &H1
Public ViewPortX As Integer
Public ViewPortY As Integer
Public ClickOrigineX As Integer, ClickOrigineY As Integer
Public soldx As Integer, soldy As Integer
Public eoldx As Integer, eoldy As Integer
Public letsgo As Boolean
'MAP ARRAY
Public TableMap() As Integer
Public ind As Integer, dist As Double
Public distnode As Double, DistEnemy As Double, distus As Double
Public Robot() As New clsRobot
Public Enemies() As New ClsEnRobots
Public xx As Single, yy As Single
Public pos1x As Single, pos1y As Single
Public pos2x As Single, pos2y As Single
Public en As Integer, enscore As Integer
Public my As Integer, myscore As Integer
Public flgup As Boolean, sexpl As Byte, eexpl As Byte
Public bullet_delay As Long
Public freespot As Integer, firing As Single
Public brect As RECT, dir As Integer, buldir As Byte
Private DX As New DirectX7
'PROGRAM CONSTANT
Public Const TILEWIDTH = 130
Public Const TILEHEIGHT = 130
Public Const SCROLLSPEED = 50

'VB/WINDOW CONSTANT
Public Const LR_LOADFROMFILE = &H10
Public Const LR_CREATEDIBSECTION = &H2000
Public Const SRCCOPY = &HCC0020

'VB/WINDOW TYPE
Public Type BITMAP
        bmType          As Long
        bmWidth         As Long
        bmHeight        As Long
        bmWidthBytes    As Long
        bmPlanes        As Integer
        bmBitsPixel     As Integer
        bmBits          As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type
'*************************************
Public Type Coordinates
    x As Long
    y As Long
End Type

Public Type bullets
    lx As Single
    ly As Single
    llive As Boolean
    cx As Single
    cy As Single
    clive As Boolean
    rx As Single
    ry As Single
    rlive As Boolean
    Speed As Single
    buldir As Byte
    type As Byte
End Type

Public Type Robot
     x As Single
     y As Single
     status As Boolean
     type As Byte
     hit As Single
     odestX As Single
     odestY As Single
     appoX As Single
     appoY As Single
     bullet() As bullets
     bulletdelay As Long
End Type
Public robots() As Robot
Public enemy() As Robot

Public Type node
     x As Single
     y As Single
     owner As Byte
End Type
Public nodes(1 To 5) As node

Public Type distnode
     ind As Byte
     dist As Double
End Type
Public nodedist(1 To 5) As distnode

Public Mousex As Long
Public Mousey As Long
Public g_Sensitivity
Public Const BufferSize = 10

Public EventHandle As Long
'Public Drawing As Boolean
Public Suspended As Boolean

Public procOld As Long

' Windows API declares and constants

Public Const GWL_WNDPROC = (-4)
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212
Public Const WM_SYSCOMMAND = &H112


'PROGRAM TYPE
Public Type MessageDisplay
        MessageSurface As DirectDrawSurface7
        MessageText As String
        Position As RECT
        DisplayTime As Long
        DisplayedTime As Long
End Type

' Various Standard API Functions for manipulation of DCs and Bitmaps
'*******************************************************************
'THOSE ONES COULD BE USEFULL
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Public Function LoadBitmapIntoDXS(DXObject As DirectDraw7, ByVal BMPFile As String, ByVal NW As Integer, ByVal NH As Integer, ByVal StretchV) As DirectDrawSurface7
'Function fit any type of surface (tiled or not tiled)
    
    Dim hBitmap As Long                 ' Handle on bitmap
    Dim dBitmap As BITMAP               ' Handle on bitmap descriptor
    Dim TempDXD As DDSURFACEDESC2       ' Surface description
    Dim TempDXS As DirectDrawSurface7   ' Created surface
    Dim dcBitmap As Long                ' Handle on image
    Dim dcDXS As Long                   ' Handle on surface context
    Dim ddck As DDCOLORKEY
    Dim i, i2
    ddck.low = 0
    ddck.high = 0
    
    'Load bitmap
    hBitmap = LoadImage(ByVal 0&, BMPFile, 0, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
    'Get bitmap descriptor
    GetObject hBitmap, Len(dBitmap), dBitmap
    'Fill DX surface description
    With TempDXD
        '.dwSize = Len(TempDXD)
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lWidth = (dBitmap.bmWidth / StretchV) * NW
        .lHeight = (dBitmap.bmHeight / StretchV) * NH
    End With
    'Create DX surface
    Set TempDXS = DXObject.CreateSurface(TempDXD)
    
    'Create API memory DC
    dcBitmap = CreateCompatibleDC(ByVal 0&)
    'Select the bitmap into API memory DC
    SelectObject dcBitmap, hBitmap
    'Restore DX surface
    TempDXS.restore
    'Get DX surface API DC
    dcDXS = TempDXS.GetDC()
    
    'Blit BMP from API DC into DX DC using standard API StretchBlt
    For i = 0 To NH
       For i2 = 0 To NW
           StretchBlt dcDXS, i2 * (dBitmap.bmWidth / StretchV), i * (dBitmap.bmHeight / StretchV), (dBitmap.bmWidth / StretchV), (dBitmap.bmHeight / StretchV), dcBitmap, 0, 0, dBitmap.bmWidth, dBitmap.bmHeight, SRCCOPY
       Next
    Next
    
    'Cleanup
    TempDXS.ReleaseDC dcDXS
    DeleteDC dcBitmap
    DeleteObject hBitmap
    TempDXS.SetColorKey DDCKEY_SRCBLT, ddck
    'Return created DX surface
    Set LoadBitmapIntoDXS = TempDXS
End Function

Public Function GetNewViewPortPosition(ByVal direction, ByVal num, ByVal w1, ByVal w2) As Long
'Function to scroll the map

If direction > 0 Then
  If num + w1 < w2 - SCROLLSPEED + 1 Then
    num = num + direction * SCROLLSPEED
  End If
Else
  If num > SCROLLSPEED - 1 Then
     num = num + direction * SCROLLSPEED
  End If
End If
GetNewViewPortPosition = num
End Function

Public Sub ebullets(dds As DirectDrawSurface7, destSurface As DirectDrawSurface7, ind As Integer)
Dim i As Integer, j As Integer

If enemy(ind).type = 1 Then
   brect.Bottom = 3
   brect.Right = 3
Else
   brect.Bottom = 6
   brect.Right = 6
End If
For i = 0 To UBound(enemy(ind).bullet)
       If enemy(ind).bullet(i).buldir = 0 Then
          If enemy(ind).type = 1 Then
             If enemy(ind).bullet(i).llive = True Then
                enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx
                enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly - enemy(ind).bullet(i).Speed
             End If
             If enemy(ind).bullet(i).rlive = True Then
                enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx
                enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry - enemy(ind).bullet(i).Speed
             End If
          Else
             If enemy(ind).bullet(i).clive = True Then
                enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx
                enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy - enemy(ind).bullet(i).Speed
             End If
          End If
        End If
       If enemy(ind).bullet(i).buldir = 1 Then
          If enemy(ind).type = 1 Then
             If enemy(ind).bullet(i).llive = True Then
                enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx + enemy(ind).bullet(i).Speed * Sin(14 * pi / 180)
                enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly - enemy(ind).bullet(i).Speed * Cos(14 * pi / 180)
             End If
             If enemy(ind).bullet(i).rlive = True Then
                enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx + enemy(ind).bullet(i).Speed * Sin(14 * pi / 180)
                enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry - enemy(ind).bullet(i).Speed * Cos(14 * pi / 180)
             End If
          Else
             If enemy(ind).bullet(i).clive = True Then
                enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx + enemy(ind).bullet(i).Speed * Sin(14 * pi / 180)
                enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy - enemy(ind).bullet(i).Speed * Cos(14 * pi / 180)
             End If
          End If
        End If
        If enemy(ind).bullet(i).buldir = 2 Then
          If enemy(ind).type = 1 Then
             If enemy(ind).bullet(i).llive = True Then
                enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx + enemy(ind).bullet(i).Speed * Sin(29 * pi / 180)
                enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly - enemy(ind).bullet(i).Speed * Cos(29 * pi / 180)
             End If
             If enemy(ind).bullet(i).rlive = True Then
                enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx + enemy(ind).bullet(i).Speed * Sin(29 * pi / 180)
                enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry - enemy(ind).bullet(i).Speed * Cos(29 * pi / 180)
             End If
          Else
             If enemy(ind).bullet(i).clive = True Then
                enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx + enemy(ind).bullet(i).Speed * Sin(29 * pi / 180)
                enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy - enemy(ind).bullet(i).Speed * Cos(29 * pi / 180)
             End If
          End If
        End If
        If enemy(ind).bullet(i).buldir = 3 Then
          If enemy(ind).type = 1 Then
             If enemy(ind).bullet(i).llive = True Then
                enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx + enemy(ind).bullet(i).Speed
                enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly - enemy(ind).bullet(i).Speed
             End If
             If enemy(ind).bullet(i).rlive = True Then
                enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx + enemy(ind).bullet(i).Speed
                enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry - enemy(ind).bullet(i).Speed
             End If
          Else
             If enemy(ind).bullet(i).clive = True Then
                enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx + enemy(ind).bullet(i).Speed
                enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy - enemy(ind).bullet(i).Speed
             End If
          End If
        End If
        If enemy(ind).bullet(i).buldir = 4 Then
          If enemy(ind).type = 1 Then
             If enemy(ind).bullet(i).llive = True Then
                enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx + enemy(ind).bullet(i).Speed * Sin(59 * pi / 180)
                enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly - enemy(ind).bullet(i).Speed * Cos(59 * pi / 180)
             End If
             If enemy(ind).bullet(i).rlive = True Then
                enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx + enemy(ind).bullet(i).Speed * Sin(59 * pi / 180)
                enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry - enemy(ind).bullet(i).Speed * Cos(59 * pi / 180)
             End If
          Else
             If enemy(ind).bullet(i).clive = True Then
                enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx + enemy(ind).bullet(i).Speed * Sin(59 * pi / 180)
                enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy - enemy(ind).bullet(i).Speed * Cos(59 * pi / 180)
             End If
          End If
        End If
        If enemy(ind).bullet(i).buldir = 5 Then
          If enemy(ind).type = 1 Then
             If enemy(ind).bullet(i).llive = True Then
                enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx + enemy(ind).bullet(i).Speed * Sin(74 * pi / 180)
                enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly - enemy(ind).bullet(i).Speed * Cos(74 * pi / 180)
             End If
             If enemy(ind).bullet(i).rlive = True Then
                enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx + enemy(ind).bullet(i).Speed * Sin(74 * pi / 180)
                enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry - enemy(ind).bullet(i).Speed * Cos(74 * pi / 180)
             End If
          Else
             If enemy(ind).bullet(i).clive = True Then
                enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx + enemy(ind).bullet(i).Speed * Sin(74 * pi / 180)
                enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy - enemy(ind).bullet(i).Speed * Cos(74 * pi / 180)
             End If
          End If
        End If
        If enemy(ind).bullet(i).buldir = 6 Then
          If enemy(ind).type = 1 Then
             If enemy(ind).bullet(i).llive = True Then
                enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx + enemy(ind).bullet(i).Speed
                enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly
             End If
             If enemy(ind).bullet(i).rlive = True Then
                enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx + enemy(ind).bullet(i).Speed
                enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry
             End If
          Else
             If enemy(ind).bullet(i).clive = True Then
                enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx + enemy(ind).bullet(i).Speed
                enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy
             End If
          End If
        End If
        If enemy(ind).bullet(i).buldir = 7 Then
          If enemy(ind).type = 1 Then
             If enemy(ind).bullet(i).llive = True Then
                enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx + enemy(ind).bullet(i).Speed * Sin(74 * pi / 180)
                enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly + enemy(ind).bullet(i).Speed * Cos(74 * pi / 180)
             End If
             If enemy(ind).bullet(i).rlive = True Then
                enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx + enemy(ind).bullet(i).Speed * Sin(74 * pi / 180)
                enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry + enemy(ind).bullet(i).Speed * Cos(74 * pi / 180)
             End If
          Else
             If enemy(ind).bullet(i).clive = True Then
                enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx + enemy(ind).bullet(i).Speed * Sin(74 * pi / 180)
                enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy + enemy(ind).bullet(i).Speed * Cos(74 * pi / 180)
             End If
          End If
        End If
        If enemy(ind).bullet(i).buldir = 8 Then
           If enemy(ind).type = 1 Then
              If enemy(ind).bullet(i).llive = True Then
                 enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx + enemy(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly + enemy(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
              If enemy(ind).bullet(i).rlive = True Then
                 enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx + enemy(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry + enemy(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
           Else
              If enemy(ind).bullet(i).clive = True Then
                 enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx + enemy(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy + enemy(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
          End If
        End If
        If enemy(ind).bullet(i).buldir = 9 Then
           If enemy(ind).type = 1 Then
              If enemy(ind).bullet(i).llive = True Then
                 enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx + enemy(ind).bullet(i).Speed
                 enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly + enemy(ind).bullet(i).Speed
              End If
              If enemy(ind).bullet(i).rlive = True Then
                 enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx + enemy(ind).bullet(i).Speed
                 enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry + enemy(ind).bullet(i).Speed
              End If
           Else
              If enemy(ind).bullet(i).clive = True Then
                 enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx + enemy(ind).bullet(i).Speed
                 enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy + enemy(ind).bullet(i).Speed
              End If
           End If
        End If
        If enemy(ind).bullet(i).buldir = 10 Then
           If enemy(ind).type = 1 Then
              If enemy(ind).bullet(i).llive = True Then
                 enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx + enemy(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly + enemy(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
              If enemy(ind).bullet(i).rlive = True Then
                 enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx + enemy(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry + enemy(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
           Else
              If enemy(ind).bullet(i).clive = True Then
                 enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx + enemy(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy + enemy(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
           End If
        End If
        If enemy(ind).bullet(i).buldir = 11 Then
           If enemy(ind).type = 1 Then
              If enemy(ind).bullet(i).llive = True Then
                 enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx + enemy(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly + enemy(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
              If enemy(ind).bullet(i).rlive = True Then
                 enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx + enemy(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry + enemy(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
           Else
              If enemy(ind).bullet(i).clive = True Then
                 enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx + enemy(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy + enemy(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
           End If
        End If
        If enemy(ind).bullet(i).buldir = 12 Then
          If enemy(ind).type = 1 Then
             If enemy(ind).bullet(i).llive = True Then
                enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx
                enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly + enemy(ind).bullet(i).Speed
             End If
             If enemy(ind).bullet(i).rlive = True Then
                enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx
                enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry + enemy(ind).bullet(i).Speed
             End If
          Else
             If enemy(ind).bullet(i).clive = True Then
                enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx
                enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy + enemy(ind).bullet(i).Speed
             End If
          End If
        End If
        If enemy(ind).bullet(i).buldir = 13 Then
          If enemy(ind).type = 1 Then
             If enemy(ind).bullet(i).llive = True Then
                enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx - enemy(ind).bullet(i).Speed * Sin(14 * pi / 180)
                enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly + enemy(ind).bullet(i).Speed * Cos(14 * pi / 180)
             End If
             If enemy(ind).bullet(i).rlive = True Then
                enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx - enemy(ind).bullet(i).Speed * Sin(14 * pi / 180)
                enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry + enemy(ind).bullet(i).Speed * Cos(14 * pi / 180)
             End If
          Else
             If enemy(ind).bullet(i).clive = True Then
                enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx - enemy(ind).bullet(i).Speed * Sin(14 * pi / 180)
                enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy + enemy(ind).bullet(i).Speed * Cos(14 * pi / 180)
             End If
          End If
        End If
        If enemy(ind).bullet(i).buldir = 14 Then
          If enemy(ind).type = 1 Then
             If enemy(ind).bullet(i).llive = True Then
                enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx - enemy(ind).bullet(i).Speed * Sin(29 * pi / 180)
                enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly + enemy(ind).bullet(i).Speed * Cos(29 * pi / 180)
             End If
             If enemy(ind).bullet(i).rlive = True Then
                enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx - enemy(ind).bullet(i).Speed * Sin(29 * pi / 180)
                enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry + enemy(ind).bullet(i).Speed * Cos(29 * pi / 180)
             End If
          Else
             If enemy(ind).bullet(i).clive = True Then
                enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx - enemy(ind).bullet(i).Speed * Sin(29 * pi / 180)
                enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy + enemy(ind).bullet(i).Speed * Cos(29 * pi / 180)
             End If
          End If
        End If
        If enemy(ind).bullet(i).buldir = 15 Then
           If enemy(ind).type = 1 Then
              If enemy(ind).bullet(i).llive = True Then
                 enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx - enemy(ind).bullet(i).Speed
                 enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly + enemy(ind).bullet(i).Speed
              End If
              If enemy(ind).bullet(i).rlive = True Then
                 enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx - enemy(ind).bullet(i).Speed
                 enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry + enemy(ind).bullet(i).Speed
              End If
           Else
              If enemy(ind).bullet(i).clive = True Then
                 enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx - enemy(ind).bullet(i).Speed
                 enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy + enemy(ind).bullet(i).Speed
              End If
           End If
        End If
        If enemy(ind).bullet(i).buldir = 16 Then
           If enemy(ind).type = 1 Then
              If enemy(ind).bullet(i).llive = True Then
                 enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx - enemy(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly + enemy(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
              If enemy(ind).bullet(i).rlive = True Then
                 enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx - enemy(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry + enemy(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
           Else
              If enemy(ind).bullet(i).clive = True Then
                 enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx - enemy(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy + enemy(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
           End If
        End If
        If enemy(ind).bullet(i).buldir = 17 Then
           If enemy(ind).type = 1 Then
              If enemy(ind).bullet(i).llive = True Then
                 enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx - enemy(ind).bullet(i).Speed * Sin(74 * pi / 180)
                 enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly + enemy(ind).bullet(i).Speed * Cos(74 * pi / 180)
              End If
              If enemy(ind).bullet(i).rlive = True Then
                 enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx - enemy(ind).bullet(i).Speed * Sin(74 * pi / 180)
                 enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry + enemy(ind).bullet(i).Speed * Cos(74 * pi / 180)
              End If
           Else
              If enemy(ind).bullet(i).clive = True Then
                 enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx - enemy(ind).bullet(i).Speed * Sin(74 * pi / 180)
                 enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy + enemy(ind).bullet(i).Speed * Cos(74 * pi / 180)
              End If
           End If
        End If
        If enemy(ind).bullet(i).buldir = 18 Then
           If enemy(ind).type = 1 Then
              If enemy(ind).bullet(i).llive = True Then
                 enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx - enemy(ind).bullet(i).Speed
                 enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly
              End If
              If enemy(ind).bullet(i).rlive = True Then
                 enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx - enemy(ind).bullet(i).Speed
                 enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry
              End If
           Else
              If enemy(ind).bullet(i).clive = True Then
                 enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx - enemy(ind).bullet(i).Speed
                 enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy
              End If
           End If
        End If
        If enemy(ind).bullet(i).buldir = 19 Then
           If enemy(ind).type = 1 Then
              If enemy(ind).bullet(i).llive = True Then
                 enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx - enemy(ind).bullet(i).Speed * Sin(74 * pi / 180)
                 enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly - enemy(ind).bullet(i).Speed * Cos(74 * pi / 180)
              End If
              If enemy(ind).bullet(i).rlive = True Then
                 enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx - enemy(ind).bullet(i).Speed * Sin(74 * pi / 180)
                 enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry - enemy(ind).bullet(i).Speed * Cos(74 * pi / 180)
              End If
           Else
              If enemy(ind).bullet(i).clive = True Then
                 enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx - enemy(ind).bullet(i).Speed * Sin(74 * pi / 180)
                 enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy - enemy(ind).bullet(i).Speed * Cos(74 * pi / 180)
              End If
           End If
        End If
        If enemy(ind).bullet(i).buldir = 20 Then
           If enemy(ind).type = 1 Then
              If enemy(ind).bullet(i).llive = True Then
                 enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx - enemy(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly - enemy(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
              If enemy(ind).bullet(i).rlive = True Then
                 enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx - enemy(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry - enemy(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
           Else
              If enemy(ind).bullet(i).clive = True Then
                 enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx - enemy(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy - enemy(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
           End If
        End If
        If enemy(ind).bullet(i).buldir = 21 Then
           If enemy(ind).type = 1 Then
              If enemy(ind).bullet(i).llive = True Then
                 enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx - enemy(ind).bullet(i).Speed
                 enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly - enemy(ind).bullet(i).Speed
              End If
              If enemy(ind).bullet(i).rlive = True Then
                 enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx - enemy(ind).bullet(i).Speed
                 enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry - enemy(ind).bullet(i).Speed
              End If
           Else
              If enemy(ind).bullet(i).clive = True Then
                 enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx - enemy(ind).bullet(i).Speed
                 enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy - enemy(ind).bullet(i).Speed
              End If
           End If
        End If
        If enemy(ind).bullet(i).buldir = 22 Then
           If enemy(ind).type = 1 Then
              If enemy(ind).bullet(i).llive = True Then
                 enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx - enemy(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly - enemy(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
              If enemy(ind).bullet(i).rlive = True Then
                 enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx - enemy(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry - enemy(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
           Else
              If enemy(ind).bullet(i).clive = True Then
                 enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx - enemy(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy - enemy(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
           End If
        End If
        If enemy(ind).bullet(i).buldir = 23 Then
           If enemy(ind).type = 1 Then
              If enemy(ind).bullet(i).llive = True Then
                 enemy(ind).bullet(i).lx = enemy(ind).bullet(i).lx - enemy(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 enemy(ind).bullet(i).ly = enemy(ind).bullet(i).ly - enemy(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
              If enemy(ind).bullet(i).rlive = True Then
                 enemy(ind).bullet(i).rx = enemy(ind).bullet(i).rx - enemy(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 enemy(ind).bullet(i).ry = enemy(ind).bullet(i).ry - enemy(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
           Else
              If enemy(ind).bullet(i).clive = True Then
                 enemy(ind).bullet(i).cx = enemy(ind).bullet(i).cx - enemy(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 enemy(ind).bullet(i).cy = enemy(ind).bullet(i).cy - enemy(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
           End If
        End If
        
        If Sqr(Abs(enemy(ind).bullet(i).lx - enemy(ind).x) * Abs(enemy(ind).bullet(i).lx - enemy(ind).x) + Abs(enemy(ind).bullet(i).ly - enemy(ind).y) * Abs(enemy(ind).bullet(i).ly - enemy(ind).y)) > 300 Then enemy(ind).bullet(i).llive = False
        If Sqr(Abs(enemy(ind).bullet(i).rx - enemy(ind).x) * Abs(enemy(ind).bullet(i).rx - enemy(ind).x) + Abs(enemy(ind).bullet(i).ry - enemy(ind).y) * Abs(enemy(ind).bullet(i).ry - enemy(ind).y)) > 300 Then enemy(ind).bullet(i).rlive = False
        If Sqr(Abs(enemy(ind).bullet(i).cx - enemy(ind).x) * Abs(enemy(ind).bullet(i).cx - enemy(ind).x) + Abs(enemy(ind).bullet(i).cy - enemy(ind).y) * Abs(enemy(ind).bullet(i).cy - enemy(ind).y)) > 300 Then enemy(ind).bullet(i).clive = False
        
        For j = 0 To 5
            If enemy(ind).bullet(i).lx >= robots(j).x - 70 And enemy(ind).bullet(i).lx <= robots(j).x + 70 And enemy(ind).bullet(i).ly <= robots(j).y + 70 And enemy(ind).bullet(i).ly >= robots(j).y - 70 Then
               robots(j).hit = robots(j).hit + 0.1
               enemy(ind).bullet(i).llive = False
            End If
            If enemy(ind).bullet(i).rx >= robots(j).x - 70 And enemy(ind).bullet(i).rx <= robots(j).x + 70 And enemy(ind).bullet(i).ry <= robots(j).y + 70 And enemy(ind).bullet(i).ry >= robots(j).y - 70 Then
               robots(j).hit = robots(j).hit + 0.1
               enemy(ind).bullet(i).rlive = False
            End If
            If enemy(ind).bullet(i).cx >= robots(j).x - 70 And enemy(ind).bullet(i).cx <= robots(j).x + 70 And enemy(ind).bullet(i).cy <= robots(j).y + 70 And enemy(ind).bullet(i).cy >= robots(j).y - 70 Then
               robots(j).hit = robots(j).hit + 0.2
               enemy(ind).bullet(i).clive = False
            End If
        Next j

        If enemy(ind).status = True Then
        If enemy(ind).type = 1 Then
           If enemy(ind).bullet(i).llive = True Then destSurface.BltFast -ViewPortX + Int(enemy(ind).bullet(i).lx), -ViewPortY + Int(enemy(ind).bullet(i).ly), dds, brect, DDBLTFAST_WAIT + DDBLTFAST_SRCCOLORKEY
           If enemy(ind).bullet(i).rlive = True Then destSurface.BltFast -ViewPortX + Int(enemy(ind).bullet(i).rx), -ViewPortY + Int(enemy(ind).bullet(i).ry), dds, brect, DDBLTFAST_WAIT + DDBLTFAST_SRCCOLORKEY
        Else
           If enemy(ind).bullet(i).clive = True Then destSurface.BltFast -ViewPortX + Int(enemy(ind).bullet(i).cx), -ViewPortY + Int(enemy(ind).bullet(i).cy), dds, brect, DDBLTFAST_WAIT + DDBLTFAST_SRCCOLORKEY
        End If
        End If

Next i

ReDim Preserve enemy(ind).bullet(UBound(enemy(ind).bullet) + 1)
If UBound(enemy(ind).bullet) > 20 Then ReDim enemy(ind).bullet(0)

End Sub
Public Sub rbullets(dds As DirectDrawSurface7, destSurface As DirectDrawSurface7, ind As Integer)
Dim i As Integer, j As Integer

If robots(ind).type = 1 Then
   brect.Bottom = 3
   brect.Right = 3
Else
   brect.Bottom = 6
   brect.Right = 6
End If
For i = 0 To UBound(robots(ind).bullet)
        If robots(ind).bullet(i).buldir = 0 Then
          If robots(ind).type = 1 Then
             If robots(ind).bullet(i).llive = True Then
                robots(ind).bullet(i).lx = robots(ind).bullet(i).lx
                robots(ind).bullet(i).ly = robots(ind).bullet(i).ly - robots(ind).bullet(i).Speed
             End If
             If robots(ind).bullet(i).rlive = True Then
                robots(ind).bullet(i).rx = robots(ind).bullet(i).rx
                robots(ind).bullet(i).ry = robots(ind).bullet(i).ry - robots(ind).bullet(i).Speed
             End If
          Else
             If robots(ind).bullet(i).clive = True Then
                robots(ind).bullet(i).cx = robots(ind).bullet(i).cx
                robots(ind).bullet(i).cy = robots(ind).bullet(i).cy - robots(ind).bullet(i).Speed
             End If
          End If
        End If
       If robots(ind).bullet(i).buldir = 1 Then
          If robots(ind).type = 1 Then
             If robots(ind).bullet(i).llive = True Then
                robots(ind).bullet(i).lx = robots(ind).bullet(i).lx + robots(ind).bullet(i).Speed * Sin(14 * pi / 180)
                robots(ind).bullet(i).ly = robots(ind).bullet(i).ly - robots(ind).bullet(i).Speed * Cos(14 * pi / 180)
             End If
             If robots(ind).bullet(i).rlive = True Then
                robots(ind).bullet(i).rx = robots(ind).bullet(i).rx + robots(ind).bullet(i).Speed * Sin(14 * pi / 180)
                robots(ind).bullet(i).ry = robots(ind).bullet(i).ry - robots(ind).bullet(i).Speed * Cos(14 * pi / 180)
             End If
          Else
             If robots(ind).bullet(i).clive = True Then
                robots(ind).bullet(i).cx = robots(ind).bullet(i).cx + robots(ind).bullet(i).Speed * Sin(14 * pi / 180)
                robots(ind).bullet(i).cy = robots(ind).bullet(i).cy - robots(ind).bullet(i).Speed * Cos(14 * pi / 180)
             End If
          End If
        End If
        If robots(ind).bullet(i).buldir = 2 Then
          If robots(ind).type = 1 Then
             If robots(ind).bullet(i).llive = True Then
                robots(ind).bullet(i).lx = robots(ind).bullet(i).lx + robots(ind).bullet(i).Speed * Sin(29 * pi / 180)
                robots(ind).bullet(i).ly = robots(ind).bullet(i).ly - robots(ind).bullet(i).Speed * Cos(29 * pi / 180)
             End If
             If robots(ind).bullet(i).rlive = True Then
                robots(ind).bullet(i).rx = robots(ind).bullet(i).rx + robots(ind).bullet(i).Speed * Sin(29 * pi / 180)
                robots(ind).bullet(i).ry = robots(ind).bullet(i).ry - robots(ind).bullet(i).Speed * Cos(29 * pi / 180)
             End If
          Else
             If robots(ind).bullet(i).clive = True Then
                robots(ind).bullet(i).cx = robots(ind).bullet(i).cx + robots(ind).bullet(i).Speed * Sin(29 * pi / 180)
                robots(ind).bullet(i).cy = robots(ind).bullet(i).cy - robots(ind).bullet(i).Speed * Cos(29 * pi / 180)
             End If
          End If
        End If
        If robots(ind).bullet(i).buldir = 3 Then
          If robots(ind).type = 1 Then
             If robots(ind).bullet(i).llive = True Then
                robots(ind).bullet(i).lx = robots(ind).bullet(i).lx + robots(ind).bullet(i).Speed
                robots(ind).bullet(i).ly = robots(ind).bullet(i).ly - robots(ind).bullet(i).Speed
             End If
             If robots(ind).bullet(i).rlive = True Then
                robots(ind).bullet(i).rx = robots(ind).bullet(i).rx + robots(ind).bullet(i).Speed
                robots(ind).bullet(i).ry = robots(ind).bullet(i).ry - robots(ind).bullet(i).Speed
             End If
          Else
             If robots(ind).bullet(i).clive = True Then
                robots(ind).bullet(i).cx = robots(ind).bullet(i).cx + robots(ind).bullet(i).Speed
                robots(ind).bullet(i).cy = robots(ind).bullet(i).cy - robots(ind).bullet(i).Speed
             End If
          End If
        End If
        If robots(ind).bullet(i).buldir = 4 Then
          If robots(ind).type = 1 Then
             If robots(ind).bullet(i).llive = True Then
                robots(ind).bullet(i).lx = robots(ind).bullet(i).lx + robots(ind).bullet(i).Speed * Sin(59 * pi / 180)
                robots(ind).bullet(i).ly = robots(ind).bullet(i).ly - robots(ind).bullet(i).Speed * Cos(59 * pi / 180)
             End If
             If robots(ind).bullet(i).rlive = True Then
                robots(ind).bullet(i).rx = robots(ind).bullet(i).rx + robots(ind).bullet(i).Speed * Sin(59 * pi / 180)
                robots(ind).bullet(i).ry = robots(ind).bullet(i).ry - robots(ind).bullet(i).Speed * Cos(59 * pi / 180)
             End If
          Else
             If robots(ind).bullet(i).clive = True Then
                robots(ind).bullet(i).cx = robots(ind).bullet(i).cx + robots(ind).bullet(i).Speed * Sin(59 * pi / 180)
                robots(ind).bullet(i).cy = robots(ind).bullet(i).cy - robots(ind).bullet(i).Speed * Cos(59 * pi / 180)
             End If
          End If
        End If
        If robots(ind).bullet(i).buldir = 5 Then
          If robots(ind).type = 1 Then
             If robots(ind).bullet(i).llive = True Then
                robots(ind).bullet(i).lx = robots(ind).bullet(i).lx + robots(ind).bullet(i).Speed * Sin(74 * pi / 180)
                robots(ind).bullet(i).ly = robots(ind).bullet(i).ly - robots(ind).bullet(i).Speed * Cos(74 * pi / 180)
             End If
             If robots(ind).bullet(i).rlive = True Then
                robots(ind).bullet(i).rx = robots(ind).bullet(i).rx + robots(ind).bullet(i).Speed * Sin(74 * pi / 180)
                robots(ind).bullet(i).ry = robots(ind).bullet(i).ry - robots(ind).bullet(i).Speed * Cos(74 * pi / 180)
             End If
          Else
             If robots(ind).bullet(i).clive = True Then
                robots(ind).bullet(i).cx = robots(ind).bullet(i).cx + robots(ind).bullet(i).Speed * Sin(74 * pi / 180)
                robots(ind).bullet(i).cy = robots(ind).bullet(i).cy - robots(ind).bullet(i).Speed * Cos(74 * pi / 180)
             End If
          End If
        End If
        If robots(ind).bullet(i).buldir = 6 Then
          If robots(ind).type = 1 Then
             If robots(ind).bullet(i).llive = True Then
                robots(ind).bullet(i).lx = robots(ind).bullet(i).lx + robots(ind).bullet(i).Speed
                robots(ind).bullet(i).ly = robots(ind).bullet(i).ly
             End If
             If robots(ind).bullet(i).rlive = True Then
                robots(ind).bullet(i).rx = robots(ind).bullet(i).rx + robots(ind).bullet(i).Speed
                robots(ind).bullet(i).ry = robots(ind).bullet(i).ry
             End If
          Else
             If robots(ind).bullet(i).clive = True Then
                robots(ind).bullet(i).cx = robots(ind).bullet(i).cx + robots(ind).bullet(i).Speed
                robots(ind).bullet(i).cy = robots(ind).bullet(i).cy
             End If
          End If
        End If
        If robots(ind).bullet(i).buldir = 7 Then
          If robots(ind).type = 1 Then
             If robots(ind).bullet(i).llive = True Then
                robots(ind).bullet(i).lx = robots(ind).bullet(i).lx + robots(ind).bullet(i).Speed * Sin(74 * pi / 180)
                robots(ind).bullet(i).ly = robots(ind).bullet(i).ly + robots(ind).bullet(i).Speed * Cos(74 * pi / 180)
             End If
             If robots(ind).bullet(i).rlive = True Then
                robots(ind).bullet(i).rx = robots(ind).bullet(i).rx + robots(ind).bullet(i).Speed * Sin(74 * pi / 180)
                robots(ind).bullet(i).ry = robots(ind).bullet(i).ry + robots(ind).bullet(i).Speed * Cos(74 * pi / 180)
             End If
          Else
             If robots(ind).bullet(i).clive = True Then
                robots(ind).bullet(i).cx = robots(ind).bullet(i).cx + robots(ind).bullet(i).Speed * Sin(74 * pi / 180)
                robots(ind).bullet(i).cy = robots(ind).bullet(i).cy + robots(ind).bullet(i).Speed * Cos(74 * pi / 180)
             End If
          End If
        End If
        If robots(ind).bullet(i).buldir = 8 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx + robots(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly + robots(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx + robots(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry + robots(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx + robots(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy + robots(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
           End If
        End If
        If robots(ind).bullet(i).buldir = 9 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx + robots(ind).bullet(i).Speed
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly + robots(ind).bullet(i).Speed
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx + robots(ind).bullet(i).Speed
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry + robots(ind).bullet(i).Speed
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx + robots(ind).bullet(i).Speed
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy + robots(ind).bullet(i).Speed
              End If
           End If
        End If
        If robots(ind).bullet(i).buldir = 10 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx + robots(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly + robots(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx + robots(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry + robots(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx + robots(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy + robots(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
           End If
        End If
        If robots(ind).bullet(i).buldir = 11 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx + robots(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly + robots(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx + robots(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry + robots(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx + robots(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy + robots(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
           End If
        End If
        If robots(ind).bullet(i).buldir = 12 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly + robots(ind).bullet(i).Speed
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry + robots(ind).bullet(i).Speed
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy + robots(ind).bullet(i).Speed
              End If
           End If
        End If
        If robots(ind).bullet(i).buldir = 13 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx - robots(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly + robots(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx - robots(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry + robots(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx - robots(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy + robots(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
           End If
        End If
        If robots(ind).bullet(i).buldir = 14 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx - robots(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly + robots(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx - robots(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry + robots(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx - robots(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy + robots(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
           End If
        End If
        If robots(ind).bullet(i).buldir = 15 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx - robots(ind).bullet(i).Speed
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly + robots(ind).bullet(i).Speed
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx - robots(ind).bullet(i).Speed
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry + robots(ind).bullet(i).Speed
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx - robots(ind).bullet(i).Speed
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy + robots(ind).bullet(i).Speed
              End If
           End If
         End If
        If robots(ind).bullet(i).buldir = 16 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx - robots(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly + robots(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx - robots(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry + robots(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx - robots(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy + robots(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
           End If
        End If
        If robots(ind).bullet(i).buldir = 17 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx - robots(ind).bullet(i).Speed * Sin(74 * pi / 180)
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly + robots(ind).bullet(i).Speed * Cos(74 * pi / 180)
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx - robots(ind).bullet(i).Speed * Sin(74 * pi / 180)
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry + robots(ind).bullet(i).Speed * Cos(74 * pi / 180)
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx - robots(ind).bullet(i).Speed * Sin(74 * pi / 180)
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy + robots(ind).bullet(i).Speed * Cos(74 * pi / 180)
              End If
           End If
        End If
        If robots(ind).bullet(i).buldir = 18 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx - robots(ind).bullet(i).Speed
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx - robots(ind).bullet(i).Speed
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx - robots(ind).bullet(i).Speed
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy
              End If
           End If
        End If
        If robots(ind).bullet(i).buldir = 19 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx - robots(ind).bullet(i).Speed * Sin(74 * pi / 180)
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly - robots(ind).bullet(i).Speed * Cos(74 * pi / 180)
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx - robots(ind).bullet(i).Speed * Sin(74 * pi / 180)
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry - robots(ind).bullet(i).Speed * Cos(74 * pi / 180)
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx - robots(ind).bullet(i).Speed * Sin(74 * pi / 180)
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy - robots(ind).bullet(i).Speed * Cos(74 * pi / 180)
              End If
           End If
        End If
        If robots(ind).bullet(i).buldir = 20 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx - robots(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly - robots(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx - robots(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry - robots(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx - robots(ind).bullet(i).Speed * Sin(59 * pi / 180)
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy - robots(ind).bullet(i).Speed * Cos(59 * pi / 180)
              End If
           End If
        End If
        If robots(ind).bullet(i).buldir = 21 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx - robots(ind).bullet(i).Speed
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly - robots(ind).bullet(i).Speed
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx - robots(ind).bullet(i).Speed
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry - robots(ind).bullet(i).Speed
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx - robots(ind).bullet(i).Speed
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy - robots(ind).bullet(i).Speed
              End If
           End If
         End If
        If robots(ind).bullet(i).buldir = 22 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx - robots(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly - robots(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx - robots(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry - robots(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx - robots(ind).bullet(i).Speed * Sin(29 * pi / 180)
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy - robots(ind).bullet(i).Speed * Cos(29 * pi / 180)
              End If
           End If
        End If
        If robots(ind).bullet(i).buldir = 23 Then
           If robots(ind).type = 1 Then
              If robots(ind).bullet(i).llive = True Then
                 robots(ind).bullet(i).lx = robots(ind).bullet(i).lx - robots(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 robots(ind).bullet(i).ly = robots(ind).bullet(i).ly - robots(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
              If robots(ind).bullet(i).rlive = True Then
                 robots(ind).bullet(i).rx = robots(ind).bullet(i).rx - robots(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 robots(ind).bullet(i).ry = robots(ind).bullet(i).ry - robots(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
           Else
              If robots(ind).bullet(i).clive = True Then
                 robots(ind).bullet(i).cx = robots(ind).bullet(i).cx - robots(ind).bullet(i).Speed * Sin(14 * pi / 180)
                 robots(ind).bullet(i).cy = robots(ind).bullet(i).cy - robots(ind).bullet(i).Speed * Cos(14 * pi / 180)
              End If
           End If
        End If
        
        If Sqr(Abs(robots(ind).bullet(i).lx - robots(ind).x) * Abs(robots(ind).bullet(i).lx - robots(ind).x) + Abs(robots(ind).bullet(i).ly - robots(ind).y) * Abs(robots(ind).bullet(i).ly - robots(ind).y)) > 300 Then robots(ind).bullet(i).llive = False
        If Sqr(Abs(robots(ind).bullet(i).rx - robots(ind).x) * Abs(robots(ind).bullet(i).rx - robots(ind).x) + Abs(robots(ind).bullet(i).ry - robots(ind).y) * Abs(robots(ind).bullet(i).ry - robots(ind).y)) > 300 Then robots(ind).bullet(i).rlive = False
        If Sqr(Abs(robots(ind).bullet(i).cx - robots(ind).x) * Abs(robots(ind).bullet(i).cx - robots(ind).x) + Abs(robots(ind).bullet(i).cy - robots(ind).y) * Abs(robots(ind).bullet(i).cy - robots(ind).y)) > 300 Then robots(ind).bullet(i).clive = False
        
        For j = 0 To 5
            If robots(ind).bullet(i).lx >= enemy(j).x - 80 And robots(ind).bullet(i).lx <= enemy(j).x + 80 And robots(ind).bullet(i).ly <= enemy(j).y + 80 And robots(ind).bullet(i).ly >= enemy(j).y - 80 Then
               enemy(j).hit = enemy(j).hit + 0.1
               robots(ind).bullet(i).llive = False
            End If
            If robots(ind).bullet(i).rx >= enemy(j).x - 80 And robots(ind).bullet(i).rx <= enemy(j).x + 80 And robots(ind).bullet(i).ry <= enemy(j).y + 80 And robots(ind).bullet(i).ry >= enemy(j).y - 80 Then
               enemy(j).hit = enemy(j).hit + 0.1
               robots(ind).bullet(i).rlive = False
            End If
            If robots(ind).bullet(i).cx >= enemy(j).x - 80 And robots(ind).bullet(i).cx <= enemy(j).x + 80 And robots(ind).bullet(i).cy <= enemy(j).y + 80 And robots(ind).bullet(i).cy >= enemy(j).y - 80 Then
               enemy(j).hit = enemy(j).hit + 0.2
               robots(ind).bullet(i).clive = False
            End If
        Next j
        
        If robots(ind).status = True Then
        If robots(ind).type = 1 Then
           If robots(ind).bullet(i).llive = True Then destSurface.BltFast -ViewPortX + Int(robots(ind).bullet(i).lx), -ViewPortY + Int(robots(ind).bullet(i).ly), dds, brect, DDBLTFAST_WAIT + DDBLTFAST_SRCCOLORKEY
           If robots(ind).bullet(i).rlive = True Then destSurface.BltFast -ViewPortX + Int(robots(ind).bullet(i).rx), -ViewPortY + Int(robots(ind).bullet(i).ry), dds, brect, DDBLTFAST_WAIT + DDBLTFAST_SRCCOLORKEY
        Else
           If robots(ind).bullet(i).clive = True Then destSurface.BltFast -ViewPortX + Int(robots(ind).bullet(i).cx), -ViewPortY + Int(robots(ind).bullet(i).cy), dds, brect, DDBLTFAST_WAIT + DDBLTFAST_SRCCOLORKEY
        End If
        End If
Next i

ReDim Preserve robots(ind).bullet(UBound(robots(ind).bullet) + 1)
If UBound(robots(ind).bullet) > 20 Then ReDim robots(ind).bullet(0)

End Sub


