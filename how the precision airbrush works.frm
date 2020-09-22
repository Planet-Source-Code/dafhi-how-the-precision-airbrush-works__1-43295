VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&next"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&back"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   4440
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   4080
      ScaleHeight     =   1515
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'how the precision airbrush works 1.0
'by dafhi

Dim sPress As Single
Dim iPress As Single 'i = 'increment'
Dim iDefin As Single

Dim sineParm1 As Single
Dim iSine1 As Single

Dim angle As Single
Dim iAngle As Single
Dim radius As Single

Dim PageIndex&
Dim XCenter&
Dim YCenter&
Dim XLEF&
Dim XRIT&
Dim YTOP&
Dim YBOT&
Dim RX&
Dim RY&
Dim RXS&
Dim RYS&
Dim RXE&
Dim RYE&
Dim XS1& 'CurrentX
Dim YS1& 'CurrentY


'arrow keys
Dim dnSizeRate As Single
Dim upSizeRate As Single
Dim rotLeftRate As Single
Dim rotRightRate As Single

Private Type PointAPI
 X As Long
 Y As Long
End Type

Private Type DimsAPI
 Width As Long
 Height As Long
End Type

Private Type PrecisionPointAPI
 px As Single
 py As Single
End Type

Dim LowLeft As PrecisionPointAPI
Dim LowRight As PrecisionPointAPI
Dim TopLeft As PrecisionPointAPI
Dim TopRight As PrecisionPointAPI
Dim Center1 As PrecisionPointAPI

Private Type BGRAQUAD
 Blue As Byte
 Green As Byte
 Red As Byte
 Alpha As Byte
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As BGRAQUAD
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Dim BMPadBytes&
Dim BMPFileHeader As BITMAPFILEHEADER   'Holds the file header
Dim BMPInfoHeader As BITMAPINFOHEADER   'Holds the info header
Dim BMPData() As Byte                   'Holds the pixel data

Private Type SAFEARRAY1D
 cDims As Integer
 fFeatures As Integer
 cbElements As Long
 cLocks As Long
 pvData As Long
 cElements As Long
 lLbound As Long
End Type

Private Type AnimSurfaceInfo
 Dib() As BGRAQUAD
 LDib() As Long
 TopRight As PointAPI
 Dims As DimsAPI
 halfW As Single
 halfH As Single
 CBWidth As Long
 SA1D As SAFEARRAY1D
 SA1D_L As SAFEARRAY1D
 EraseDib() As Long
 LBotLeftErase() As Long
 LTopLeftErase() As Long
 LEraseWidth() As Long
 EraseSpriteCount As Long
End Type

Private BM As BITMAP
Private PicDib As AnimSurfaceInfo
Private FormDib As AnimSurfaceInfo

Private Type PicBmp
    Size As Long
    Type As PictureTypeConstants
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy&)
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject&, ByVal nCount&, lpObject As Any) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc&, pBitmapInfo As BITMAPINFO, ByVal un&, lplpVoid&, ByVal handle&, ByVal dw&) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle&, IPic As IPicture) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Dim TimeElapsed As Double
Dim Freq As Currency
Dim TimeNow As Currency
Dim LastTime As Currency
Dim FrameCounter As Integer
Dim FPS As Integer
Dim TickSum As Double

Private Const B255 As Byte = 255
Private Const pi As Single = 3.1415926535
Private Const twoPi As Single = 2 * pi
Private Const BrightPine As Long = 256 * 118
Private Const UserComment As Long = vbBlue + 52 + 58& * 256
Private Const Gray128 As Long = 128 + 128 * 256& + 128& * 65536
Private Const UserComment2 As Long = 0 + 128 * 256& + 128& * 65536

Dim bRunning As Boolean

Private Enum MixMode
 CSolid
 CShift
 invert
End Enum

Private Type AirbrushStruct
 Red As Byte
 Green As Byte
 Blue As Byte
 Pressure As Byte
 definition As Single
 center_x As Single
 center_y As Single
 diameter As Single
 bFilledCircle As Boolean
 bDoErase As Boolean
 CMode As MixMode
End Type

'BlitAirbrush
Dim AddDrawHeight&

'BlitAirbrush, EraseAirbrushes&
Dim AddDrawWidth&
Dim DrawX&
Dim DrawY&

'ColorShift
Dim rgb_shift_intensity!
Dim sR!
Dim sG!
Dim sB!

Dim maximu!
Dim minimu!
Dim iSubt!
Dim bytMaxMin_diff As Byte

'AnimQuadSkew, EraseQuadSkew
Dim X4&
Dim Y4&
Dim plotx!
Dim ploty!
Dim ix!
Dim iy!
Dim iix!
Dim iiy!
Dim ixSave!
Dim iySave!
Dim tmpRise!
Dim tmpRun!
Dim ImgWide& 'Written to by TruecolorBmpToBGRA1D
Dim ImgHigh& 'same with this
Dim ImgWideM1&
Dim ImgHighM1&
Dim xLeft!
Dim yLeft!
Dim ixLeft!
Dim iyLeft!
Dim diff!
Dim Elem&
Dim DibIndex&
Dim ImageBytes() As BGRAQUAD

Dim LColorSave&

''ShowPage
Dim base_y!
Dim ceil_y!
Dim cone_base!
Dim mid_x!
Dim left_x!
Dim right_x!
Dim cone_height!
Dim cone_diam!
Dim cone_top!
Dim box_ceil!

Dim pageBrushCenterX!
Dim pageBrushCenterY!
Dim pageRadius!
Dim Grid_Magnification&

Dim BGR&

Dim BGRed&, BGGreen&, BGBlue&
Dim FGRed&, FGGreen&, FGBlue&

Dim blit_x!, blit_y!
Dim bRegularDistFormula As Boolean

Dim RXL&
Dim Airbrush As AirbrushStruct


Private Sub cmdBack_Click()
cmdNext.Enabled = True
PageIndex = PageIndex - 1
InfoCaption
ShowPage
End Sub
Private Sub cmdNext_Click()
cmdBack.Enabled = True
PageIndex = PageIndex + 1
InfoCaption
ShowPage
End Sub
Private Sub InfoCaption()
Caption = "Page " & PageIndex + 1 & ":  "
Select Case PageIndex
Case 0
 Caption = Caption + "Visualizing the brush as a cone"
Case 1
 Caption = Caption + "The application of brush Pressure"
Case 2
 Caption = Caption + "Understanding Brush Definition"
Case 3
 Caption = Caption + "One Set of Defaults"
Case 4
 Caption = Caption + "Behind the scenes in the blit sub"
Case 5
 Caption = Caption + "Behind the scenes part 2"
Case 6
 Caption = Caption + "Some Background Info"
Case 7
 Caption = Caption + "Precision means using the decimal point"
Case 8
 Caption = Caption + "Close-up view .. compute left side of brush rect"
Case 9
 Caption = Caption + "Left column and brush center x have initial delta"
Case 10
 Caption = Caption + "Top row has initial delta with brush center y"
Case 11
 Caption = Caption + "Most of the loop"
End Select
End Sub

Private Sub Form_Load()
ScaleMode = vbPixels
Picture1.ScaleMode = vbPixels

Airbrush.diameter = 100
Airbrush.Pressure = 255
Airbrush.definition = 1

Form1.AutoRedraw = True

End Sub

Private Sub CleanUp()
'Important - if you create AnimSurfaceInfo types,
'clean up their .Dib and .LDib pointers like this
 CopyMemory ByVal VarPtrArray(PicDib.Dib), 0&, 4
 CopyMemory ByVal VarPtrArray(PicDib.LDib), 0&, 4
 CopyMemory ByVal VarPtrArray(FormDib.Dib), 0&, 4
 CopyMemory ByVal VarPtrArray(FormDib.LDib), 0&, 4
End Sub


Private Sub Form_Activate()
Dim aspectR!
Dim Red As Byte
Dim Green As Byte
Dim Blue As Byte
Dim sR3!
Dim sG3!
Dim sB3!
Dim colorshift_rate!

AnimPicSurface Picture1, PicDib

Center1.px = PicDib.halfW
Center1.py = PicDib.halfH

blit_x = PicDib.halfW
blit_y = PicDib.halfH

InfoCaption
ShowPage

'solid color fill
Red = 158
Green = 128
Blue = 155
LColorSave = RGB(Blue, Green, Red)
For RX = 0 To UBound(PicDib.Dib)
 PicDib.LDib(RX) = LColorSave
 PicDib.EraseDib(RX) = LColorSave
Next

''Temp
ImgWide = 80
ImgHigh = 80
MakeBGRA1D ImageBytes, ImgWide&, ImgHigh&, 255, 255, 0
''comment out TruecolorBmp.. below to see MakeBGRA1D work

''Writes to ImgWide, ImgHigh
'TruecolorBmpToBGRA1D ImageBytes, "hola.bmp", ImgWide&, ImgHigh&


'Center1.px = ScaleWidth / 2&
'Center1.py = ScaleHeight / 2&

''More Temp
aspectR = ImgWide / ImgHigh

iAngle = 0.2!
radius = Sqr((ImgWide / 2) ^ 2 + (ImgHigh / 2) ^ 2)

ImgWideM1 = ImgWide - 1
ImgHighM1 = ImgHigh - 1

bRunning = True

QueryPerformanceFrequency Freq
QueryPerformanceCounter LastTime

sB3 = 255 'full saturation .. 1 of 3 elements of course
colorshift_rate = 0.5

Do While bRunning
 If PicDib.Dims.Height > 0& Then

  'time-based modelling
  QueryPerformanceCounter TimeNow

  'AnimQuadSkew PicDib
    
  rgb_shift_intensity = colorshift_rate * TimeElapsed
  iSine1 = TimeElapsed
  If rgb_shift_intensity > 1 Then rgb_shift_intensity = 1
  sR = 255 'sR3
  sG = 255 'sG3
  sB = 255 'sB3
  Airbrush.Red = sR
  Airbrush.Green = sG
  Airbrush.Blue = sB
  ColorShift
  sR3 = sR
  sG3 = sG
  sB3 = sB
  Airbrush.bDoErase = True
  Airbrush.bFilledCircle = True
  BlitAirbrush PicDib, blit_x, blit_y, True
  'BlitAirbrush FormDib, Rnd * FormDib.Dims.Width, Rnd * FormDib.Dims.Height
  ''uncomment AnimPicSurface in Form_Resize
  
  '' increments
  sineParm1 = sineParm1 + iSine1
  If sineParm1 > twoPi Then sineParm1 = sineParm1 - twoPi
  
  Picture1.Refresh
  Refresh
  
  'Erase everything
  EraseAirbrushes PicDib
  EraseAirbrushes FormDib
  
  EraseQuadSkew PicDib

  'Compute FPS and the #1 time-base variable, TimeElapsed
  TimeElapsed = (TimeNow - LastTime) / Freq
  LastTime = TimeNow
  TickSum = TickSum + TimeElapsed
      
  'Calculate current frames per second
  If TickSum > 1 Then
     FPS = FrameCounter / TickSum
     FrameCounter = 0
     TickSum = 0
     'Caption = "FPS: " & FPS
  End If
    
  FrameCounter = FrameCounter + 1

  angle = angle + iAngle * TimeElapsed
  If angle > twoPi Then angle = angle - twoPi

 End If 'FormDib.Dims.Height > 0
 
 DoEvents
 
Loop

CleanUp

End

End Sub


Private Sub Form_Resize()
'AnimPicSurface Form1, FormDib
End Sub
Private Sub AnimQuadSkew(Surf As AnimSurfaceInfo)
Dim ixHigh!
Dim iyHigh!

LowLeft.px = Center1.px - radius * Cos(angle)
LowLeft.py = Center1.py - radius * Sin(angle)

'low right is next.  if we had an image of 3 wide
'there would be a step of 2 to low right corner

'low right corner is a run/rise flip from center
'of rotation and low left corner
tmpRun = Center1.px - LowLeft.px
tmpRise = Center1.py - LowLeft.py

'the other 2 points are almost as easy
TopRight.px = Center1.px + tmpRun
TopRight.py = Center1.py + tmpRise

'normally, x is associated with 'run'
LowRight.px = Center1.px + tmpRise
LowRight.py = Center1.py - tmpRun


TopLeft.px = Center1.px - tmpRise
TopLeft.py = Center1.py + tmpRun

'ix = 'increment x'
ix = (LowRight.px - LowLeft.px) / ImgWideM1
iy = (LowRight.py - LowLeft.py) / ImgWideM1

ixHigh = (TopRight.px - TopLeft.px) / ImgWideM1
iyHigh = (TopRight.py - TopLeft.py) / ImgWideM1

iix = (ixHigh - ix) / ImgHighM1
iiy = (iyHigh - iy) / ImgHighM1

xLeft = LowLeft.px
yLeft = LowLeft.py

'marker's increment
ixLeft = (TopLeft.px - LowLeft.px) / ImgHighM1
iyLeft = (TopLeft.py - LowLeft.py) / ImgHighM1

ixSave = ix
iySave = iy

Elem = 0 'for ImageBytes() array

For Y4 = 1 To ImgHigh
'reset plot with each row
 plotx = xLeft
 ploty = yLeft
 For X4 = 1 To ImgWide

  'round plotx
  RX = Int(plotx)
  diff = plotx - RX
  If diff >= 0.5 Then RX = RX + 1&
  
  'round ploty
  RY = Int(ploty)
  diff = ploty - RY
  If diff >= 0.5 Then RY = RY + 1&
  
  DibIndex = RX + PicDib.Dims.Width * RY
  If DibIndex >= 0 And DibIndex < PicDib.SA1D.cElements Then
  PicDib.Dib(DibIndex).Blue = ImageBytes(Elem).Blue
  PicDib.Dib(DibIndex).Green = ImageBytes(Elem).Green
  PicDib.Dib(DibIndex).Red = ImageBytes(Elem).Red
  End If
  
  plotx = plotx + ix
  ploty = ploty + iy
  Elem = Elem + 1&
  
 Next X4
 
 xLeft = xLeft + ixLeft
 yLeft = yLeft + iyLeft
 ix = ix + iix
 iy = iy + iiy
 
Next Y4
End Sub
Private Sub EraseQuadSkew(Surf As AnimSurfaceInfo)
ix = ixSave
iy = iySave
xLeft = LowLeft.px
yLeft = LowLeft.py
For Y4 = 1 To ImgHigh
 plotx = xLeft
 ploty = yLeft
 For X4 = 1 To ImgWide
 
  'round plotx
  RX = Int(plotx)
  diff = plotx - RX
  If diff >= 0.5 Then RX = RX + 1&
  
  'round ploty
  RY = Int(ploty)
  diff = ploty - RY
  If diff >= 0.5 Then RY = RY + 1&
  
  DibIndex = RX + PicDib.Dims.Width * RY
  If DibIndex >= 0 And DibIndex < PicDib.SA1D.cElements Then
  PicDib.LDib(DibIndex) = LColorSave
  End If
  plotx = plotx + ix
  ploty = ploty + iy
  Elem = Elem + 1&
 Next X4
 xLeft = xLeft + ixLeft
 yLeft = yLeft + iyLeft
 ix = ix + iix
 iy = iy + iiy
Next Y4
End Sub

Private Sub BlitAirbrush(Surf As AnimSurfaceInfo, X!, Y!, Optional bSolidDot As Boolean)
Dim ClipBot&
Dim ClipWidthLeft&
Dim DrawRight&
Dim DrawLeft&
Dim DrawTop&
Dim DrawBot&
Dim brush_radius!
Dim brush_height!
Dim brush_slope!
Dim height_Sq!
Dim px!
Dim py!
Dim RX&
Dim RY&
Dim RXR&
Dim RYT&
Dim RYB&
Dim Left_&
Dim delta_ySq!
Dim delta_y!
Dim delta_x!
Dim delta_left!
Dim deltas_xy_Sq!
Dim Bright&
Dim sR2!
Dim sG2!
Dim sB2!
Dim Press_x_2 As Long

 If Airbrush.diameter > 0 Then
  
 brush_radius = Airbrush.diameter / 2
 
 brush_height = Airbrush.Pressure * Airbrush.definition
 brush_slope = brush_height / brush_radius
 
 height_Sq = brush_height * brush_height
 
 DrawLeft = RealRound(X - brush_radius)
 DrawBot = RealRound(Y - brush_radius)
 DrawRight = RealRound(X + brush_radius)
 DrawTop = RealRound(Y + brush_radius)
 
 If DrawLeft < 0 Then DrawLeft = 0
 If DrawBot < 0 Then DrawBot = 0
 If DrawRight > Surf.TopRight.X Then DrawRight = Surf.TopRight.X
 If DrawTop > Surf.TopRight.Y Then DrawTop = Surf.TopRight.Y
 
 delta_left = (DrawLeft - X) * brush_slope
 delta_y = (DrawBot - Y) * brush_slope
  
 AddDrawWidth = DrawRight - DrawLeft
 AddDrawHeight = DrawTop - DrawBot
 
 DrawBot = DrawBot * Surf.Dims.Width + DrawLeft
 DrawTop = DrawBot + Surf.Dims.Width * AddDrawHeight
 
 If Airbrush.bDoErase Then
  
  Surf.EraseSpriteCount = Surf.EraseSpriteCount + 1&
   
  ReDim Preserve Surf.LBotLeftErase(1& To Surf.EraseSpriteCount)
  ReDim Preserve Surf.LTopLeftErase(1& To Surf.EraseSpriteCount)
  ReDim Preserve Surf.LEraseWidth(1& To Surf.EraseSpriteCount)
  
  Surf.LBotLeftErase(Surf.EraseSpriteCount) = DrawBot
  Surf.LTopLeftErase(Surf.EraseSpriteCount) = DrawTop
  Surf.LEraseWidth(Surf.EraseSpriteCount) = AddDrawWidth
  
  If Airbrush.CMode = CSolid Then
  
  If bSolidDot Then
  
  If bRegularDistFormula Then
  
  For DrawY& = DrawBot& To DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = delta_left!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight&
    Bright& = Sqr#(delta_x * delta_x + delta_ySq)
     If Bright& > Airbrush.Pressure Then Bright& = Airbrush.Pressure
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     sR = sR! + Bright * (Airbrush.Red - sR!) / B255
     sG = sG! + Bright * (Airbrush.Green - sG!) / B255
     sB = sB! + Bright * (Airbrush.Blue - sB!) / B255
     Surf.Dib(DrawX).Blue = sB
     Surf.Dib(DrawX).Green = sG
     Surf.Dib(DrawX).Red = sR
    delta_x! = delta_x! + brush_slope!
   Next
   delta_y! = delta_y! + brush_slope!
  Next
  
  Else 'invert distance formula
  
  For DrawY& = DrawBot& To DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = delta_left!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight&
    deltas_xy_Sq = delta_x * delta_x + delta_ySq
    If deltas_xy_Sq < height_Sq Then
    Bright& = brush_height - Sqr#(deltas_xy_Sq)
     If Bright& > Airbrush.Pressure Then Bright& = Airbrush.Pressure
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     sR = sR! + Bright * (Airbrush.Red - sR!) / B255
     sG = sG! + Bright * (Airbrush.Green - sG!) / B255
     sB = sB! + Bright * (Airbrush.Blue - sB!) / B255
     Surf.Dib(DrawX).Blue = sB
     Surf.Dib(DrawX).Green = sG
     Surf.Dib(DrawX).Red = sR
    End If
    delta_x! = delta_x! + brush_slope!
   Next
   delta_y! = delta_y! + brush_slope!
  Next
  
  End If
  
  Else 'circle outline
  
  For DrawY& = DrawBot& To DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = delta_left!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight&
    Bright& = brush_height - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > Airbrush.Pressure Then Bright& = Press_x_2 - Bright
     If Bright > 0& Then
      sB! = Surf.Dib(DrawX).Blue
      sG! = Surf.Dib(DrawX).Green
      sR! = Surf.Dib(DrawX).Red
      Surf.Dib(DrawX).Blue = sB + Bright * (Airbrush.Blue - sB!) / B255
      Surf.Dib(DrawX).Green = sG + Bright * (Airbrush.Green - sG!) / B255
      Surf.Dib(DrawX).Red = sR + Bright * (Airbrush.Red - sR!) / B255
     End If
    End If
    delta_x! = delta_x! + brush_slope!
   Next
   delta_y! = delta_y! + brush_slope!
  Next DrawY
  
  End If
  
  ElseIf Airbrush.CMode = invert Then
  
  If bSolidDot Then
  
  For DrawY& = DrawBot& To DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y! 'speeds up what's inside the next loop
   delta_x! = delta_left! 'starting at the left side with each new row.
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight&
    Bright& = brush_height - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright& > 0& Then
     If Bright > Airbrush.Pressure Then Bright = Airbrush.Pressure
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     sR2 = 255! - sR
     sG2 = 255! - sG
     sB2 = 255! - sB
     Surf.Dib(DrawX).Blue = sB + Bright * (sB2 - sB!) / B255
     Surf.Dib(DrawX).Green = sG + Bright * (sG2 - sG!) / B255
     Surf.Dib(DrawX).Red = sR + Bright * (sR2 - sR!) / B255
    End If
    delta_x! = delta_x! + brush_slope!
   Next
   delta_y! = delta_y! + brush_slope!
  Next
  
  Else 'circle outline
  
  For DrawY& = DrawBot& To DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = delta_left!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight&
    Bright& = brush_height - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > Airbrush.Pressure Then Bright& = Press_x_2 - Bright
     If Bright > 0& Then
      sB! = Surf.Dib(DrawX).Blue
      sG! = Surf.Dib(DrawX).Green
      sR! = Surf.Dib(DrawX).Red
      sR2 = 255& - sR
      sG2 = 255& - sG
      sB2 = 255& - sB
      Surf.Dib(DrawX).Blue = sB + Bright * (sB2 - sB!) / B255
      Surf.Dib(DrawX).Green = sG + Bright * (sG2 - sG!) / B255
      Surf.Dib(DrawX).Red = sR + Bright * (sR2 - sR!) / B255
     End If
    End If
    delta_x! = delta_x! + brush_slope!
   Next
   delta_y! = delta_y! + brush_slope!
  Next DrawY
  
  End If
  
  ElseIf Airbrush.CMode = CShift Then
  
  If bSolidDot Then
  
  For DrawY& = DrawBot& To DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y! 'speeds up what's inside the next loop
   delta_x! = delta_left! 'starting at the left side with each new row.
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight&
    Bright& = brush_height - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright& > 0& Then
     If Bright > Airbrush.Pressure Then Bright = Airbrush.Pressure
     rgb_shift_intensity = Bright& / B255
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     ColorShift
     Surf.Dib(DrawX).Blue = sB
     Surf.Dib(DrawX).Green = sG
     Surf.Dib(DrawX).Red = sR
    End If
    delta_x! = delta_x! + brush_slope!
   Next
   delta_y! = delta_y! + brush_slope!
  Next
  
  Else 'circle outline
  
  For DrawY& = DrawBot& To DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = delta_left!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight&
    Bright& = brush_height - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > Airbrush.Pressure Then Bright& = Press_x_2 - Bright
     If Bright > 0& Then
      rgb_shift_intensity = Bright& / B255
      sB! = Surf.Dib(DrawX).Blue
      sG! = Surf.Dib(DrawX).Green
      sR! = Surf.Dib(DrawX).Red
      ColorShift
      Surf.Dib(DrawX).Blue = sB
      Surf.Dib(DrawX).Green = sG
      Surf.Dib(DrawX).Red = sR
     End If
    End If
    delta_x! = delta_x! + brush_slope!
   Next
   delta_y! = delta_y! + brush_slope!
  Next DrawY
  
  End If
  
  End If 'AirBrush.CMode = CSolid
 
  
 Else 'not erasing
  
    
  If Airbrush.CMode = CSolid Then
  
  If bSolidDot Then
  
  For DrawY& = DrawBot& To DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = delta_left!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight&
    Bright& = brush_height - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > Airbrush.Pressure Then Bright& = Airbrush.Pressure
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     Surf.Dib(DrawX).Blue = sB + Bright * (Airbrush.Blue - sB!) / B255
     Surf.Dib(DrawX).Green = sG + Bright * (Airbrush.Green - sG!) / B255
     Surf.Dib(DrawX).Red = sR + Bright * (Airbrush.Red - sR!) / B255
     Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
    End If
    delta_x! = delta_x! + brush_slope!
   Next
   delta_y! = delta_y! + brush_slope!
  Next
  
  Else 'circle outline
  
  For DrawY& = DrawBot& To DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = delta_left!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight&
    Bright& = brush_height - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > Airbrush.Pressure Then Bright& = Press_x_2 - Bright
     If Bright > 0& Then
      sB! = Surf.Dib(DrawX).Blue
      sG! = Surf.Dib(DrawX).Green
      sR! = Surf.Dib(DrawX).Red
      Surf.Dib(DrawX).Blue = sB + Bright * (Airbrush.Blue - sB!) / B255
      Surf.Dib(DrawX).Green = sG + Bright * (Airbrush.Green - sG!) / B255
      Surf.Dib(DrawX).Red = sR + Bright * (Airbrush.Red - sR!) / B255
      Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
     End If
    End If
    delta_x! = delta_x! + brush_slope!
   Next DrawX
   delta_y! = delta_y! + brush_slope!
  Next DrawY
  
  End If 'outlined
 
  ElseIf Airbrush.CMode = invert Then
  
  If bSolidDot Then
  
  For DrawY& = DrawBot& To DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = delta_left!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight&
    Bright& = brush_height - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright& > 0& Then
     If Bright > Airbrush.Pressure Then Bright = Airbrush.Pressure
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     sR2 = 255& - sR
     sG2 = 255& - sG
     sB2 = 255& - sB
     Surf.Dib(DrawX).Blue = sB + Bright * (sB2 - sB!) / B255
     Surf.Dib(DrawX).Green = sG + Bright * (sG2 - sG!) / B255
     Surf.Dib(DrawX).Red = sR + Bright * (sR2 - sR!) / B255
     Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
    End If
    delta_x! = delta_x! + brush_slope!
   Next DrawX
   delta_y! = delta_y! + brush_slope!
  Next DrawY
  
  Else 'circle outline
  
  For DrawY& = DrawBot& To DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = delta_left!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight&
    Bright& = brush_height - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > Airbrush.Pressure Then Bright& = Press_x_2 - Bright
     If Bright > 0& Then
      sB! = Surf.Dib(DrawX).Blue
      sG! = Surf.Dib(DrawX).Green
      sR! = Surf.Dib(DrawX).Red
      sR2 = 255& - sR
      sG2 = 255& - sG
      sB2 = 255& - sB
      Surf.Dib(DrawX).Blue = sB + Bright * (sB2 - sB!) / B255
      Surf.Dib(DrawX).Green = sG + Bright * (sG2 - sG!) / B255
      Surf.Dib(DrawX).Red = sR + Bright * (sR2 - sR!) / B255
      Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
     End If
    End If
    delta_x! = delta_x! + brush_slope!
   Next DrawX
   delta_y! = delta_y! + brush_slope!
  Next DrawY
  
  End If 'outlined
  
  ElseIf Airbrush.CMode = CShift Then
  
  If bSolidDot Then
  
  For DrawY& = DrawBot& To DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = delta_left!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight&
    Bright& = brush_height - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright& > 0& Then
     If Bright > Airbrush.Pressure Then Bright = Airbrush.Pressure
     rgb_shift_intensity = Bright& / B255
     sB! = Surf.Dib(DrawX).Blue
     sG! = Surf.Dib(DrawX).Green
     sR! = Surf.Dib(DrawX).Red
     ColorShift
     Surf.Dib(DrawX).Blue = sB
     Surf.Dib(DrawX).Green = sG
     Surf.Dib(DrawX).Red = sR
     Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
    End If
    delta_x! = delta_x! + brush_slope!
   Next DrawX
   delta_y! = delta_y! + brush_slope!
  Next DrawY
  
  Else 'circle outline
  
  For DrawY& = DrawBot& To DrawTop& Step Surf.Dims.Width
   delta_ySq! = delta_y! * delta_y!
   delta_x! = delta_left!
   DrawRight = DrawY + AddDrawWidth
   For DrawX& = DrawY To DrawRight&
    Bright& = brush_height - Sqr#(delta_x! * delta_x! + delta_ySq!)
    If Bright > 0& Then
     If Bright& > Airbrush.Pressure Then Bright& = Press_x_2 - Bright
     If Bright > 0& Then
      rgb_shift_intensity = Bright& / B255
      sB! = Surf.Dib(DrawX).Blue
      sG! = Surf.Dib(DrawX).Green
      sR! = Surf.Dib(DrawX).Red
      ColorShift
      Surf.Dib(DrawX).Blue = sB
      Surf.Dib(DrawX).Green = sG
      Surf.Dib(DrawX).Red = sR
      Surf.EraseDib(DrawX) = Surf.LDib(DrawX)
     End If
    End If
    delta_x! = delta_x! + brush_slope!
   Next DrawX
   delta_y! = delta_y! + brush_slope!
  Next DrawY
  
  End If 'bSolid
  
  End If 'AirBrush.CMode
  
 End If 'erasing
  
 End If 'Airbrush.diameter > 0
  
End Sub
Private Sub EraseAirbrushes(Surf As AnimSurfaceInfo)
Dim LBotLeft&
Dim LTopLeft&
Dim LEraseWide&
Dim N&

 For N& = 1& To Surf.EraseSpriteCount&
  
  LBotLeft = Surf.LBotLeftErase(N)
  LTopLeft = Surf.LTopLeftErase(N)
  LEraseWide = Surf.LEraseWidth(N)
  For DrawY = LBotLeft To LTopLeft Step Surf.Dims.Width
   AddDrawWidth = DrawY + LEraseWide
   For DrawX = DrawY To AddDrawWidth
    Surf.LDib(DrawX) = Surf.EraseDib(DrawX)
   Next
  Next
  
 Next N& 'Next Sprite
 
 Surf.EraseSpriteCount = 0&

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
Case vbKeyEscape
 bRunning = False
 
Case vbKeyUp
 upSizeRate = upSizeRate + 0.04
 radius = radius + upSizeRate
 dnSizeRate = 0
 
Case vbKeyDown
 dnSizeRate = dnSizeRate + 0.04
 radius = radius - dnSizeRate
 upSizeRate = 0
 
Case vbKeyLeft
 iAngle = rotLeftRate
 rotLeftRate = rotLeftRate + 0.02
 rotRightRate = 0
 
Case vbKeyRight
 iAngle = -rotRightRate
 rotRightRate = rotRightRate + 0.02
 rotLeftRate = 0
 
End Select

End Sub
Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub
Private Sub cmdNext_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub
Private Sub cmdBack_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub


Private Sub TruecolorBmpToBGRA1D(Bytes() As BGRAQUAD, strFilename$, RetWide&, RetHigh&)
Dim TrackY&
Dim ScanLineWidthBytes&
Dim XLng& '1d.  XLng marks the position.  YLng is a reference.
Dim YLng&
Dim AddDrawWidthBytes&
Dim DrawRight&
Dim TopLeft&
Dim WidthBytes&
Dim X_Max&
Dim DrawX&

 TrueColorBMPToData strFilename
  
 'These are used to reference 1d array file bytes
 WidthBytes& = BMPInfoHeader.biWidth * 3& + BMPadBytes

 If WidthBytes > 0 Then
 
  RetWide = BMPInfoHeader.biWidth
  RetHigh = BMPInfoHeader.biHeight
  
  ReDim Bytes(0 To BMPInfoHeader.biWidth * BMPInfoHeader.biHeight - 1)
  
  X_Max = RetWide - 1
  DrawRight = X_Max * 3
  
  TopLeft = WidthBytes * (RetHigh - 1&)
  ScanLineWidthBytes = RetWide
  For YLng& = 0& To TopLeft Step WidthBytes
   DrawX = TrackY
   AddDrawWidthBytes& = YLng& + DrawRight&
   For XLng& = YLng& To AddDrawWidthBytes& Step 3&
    Bytes(DrawX).Blue = BMPData(XLng&)
    Bytes(DrawX).Green = BMPData(XLng& + 1&)
    Bytes(DrawX).Red = BMPData(XLng& + 2&)
    DrawX = DrawX + 1&
   Next XLng
   TrackY = TrackY + ScanLineWidthBytes
  Next YLng

 End If 'Widthbytes > 0
  
End Sub

Private Sub MakeBGRA1D(Bytes() As BGRAQUAD, Wide&, High&, Red As Byte, Green As Byte, Blue As Byte)
Dim TrackY&
Dim ScanLineWidthBytes&
Dim XLng& '1d.  XLng marks the position.  YLng is a reference.
Dim YLng&
Dim AddDrawWidthBytes&
Dim DrawRight&
Dim TopLeft&
Dim WidthBytes&
Dim X_Max&
Dim DrawX&
Dim N&

 N = 3 * Wide
 BMPadBytes = ((N + 3) And &HFFFFFFFC) - N
  
 'These are used to reference 1d array file bytes
 WidthBytes& = Wide * 3& + BMPadBytes

 If WidthBytes > 0 Then
  
  ReDim Bytes(0 To Wide * High - 1)
  
  X_Max = Wide - 1
  DrawRight = X_Max * 3
  
  TopLeft = WidthBytes * (High - 1&)
  ScanLineWidthBytes = Wide '* 4&
  For YLng& = 0& To TopLeft Step WidthBytes
   DrawX = TrackY
   AddDrawWidthBytes& = YLng& + DrawRight&
   For XLng& = YLng& To AddDrawWidthBytes& Step 3&
    Bytes(DrawX).Blue = Blue
    Bytes(DrawX).Green = Green
    Bytes(DrawX).Red = Red
    DrawX = DrawX + 1&
   Next XLng
   TrackY = TrackY + ScanLineWidthBytes
  Next YLng

 End If 'Widthbytes > 0
  
End Sub

Private Sub TrueColorBMPToData(strFilename$)
Dim N&
On Error GoTo errout:

 Open (App.Path & "\" & strFilename) For Binary As #1
   Get #1, 1, BMPFileHeader
   Get #1, , BMPInfoHeader
   With BMPInfoHeader
    N = 3 * .biWidth '(red, green, blue) * width
    BMPadBytes = ((N + 3) And &HFFFFFFFC) - N
    ReDim BMPData(.biHeight * (BMPadBytes + .biWidth * .biBitCount / 8))
   End With
   Get #1, , BMPData
 Close #1
 Exit Sub
 
errout:
 
 BMPInfoHeader.biWidth = 150
 BMPInfoHeader.biHeight = 150
 BMPInfoHeader.biBitCount = 24
 With BMPInfoHeader
  N = 3 * .biWidth
  BMPadBytes = ((N + 3) And &HFFFFFFFC) - N
  ReDim BMPData(.biHeight * (BMPadBytes + .biWidth * .biBitCount / 8))
  For N = 0 To UBound(BMPData)
   BMPData(N) = Rnd * 255
  Next
 End With

End Sub

Private Function CreatePicture(ByVal nWidth&, ByVal nHeight&, ByVal BitDepth&) As Picture
Dim Pic As PicBmp, IID_IDispatch As GUID
Dim BMI As BITMAPINFO
With BMI.bmiHeader
.biSize = Len(BMI.bmiHeader)
.biWidth = nWidth
.biHeight = nHeight
.biPlanes = 1
.biBitCount = BitDepth
End With
Pic.hBmp = CreateDIBSection(0, BMI, 0, 0, 0, 0)
IID_IDispatch.Data1 = &H20400: IID_IDispatch.Data4(0) = &HC0: IID_IDispatch.Data4(7) = &H46
Pic.Size = Len(Pic)
Pic.Type = vbPicTypeBitmap
OleCreatePictureIndirect Pic, IID_IDispatch, 1, CreatePicture
If CreatePicture = 0 Then Set CreatePicture = Nothing
End Function

Private Sub AnimPicSurface(Obj As Object, Surf As AnimSurfaceInfo)
 
 Surf.Dims.Width = Obj.ScaleWidth
 Surf.Dims.Height = Obj.ScaleHeight
 Surf.TopRight.X = Surf.Dims.Width - 1
 Surf.TopRight.Y = Surf.Dims.Height - 1
 
 Surf.halfW = Surf.Dims.Width / 2&
 Surf.halfH = Surf.Dims.Height / 2&
 
 Surf.EraseSpriteCount = 0
 
 'Destroy any pointer this array may have had
 CopyMemory ByVal VarPtrArray(Surf.Dib), 0&, 4
 
 If Surf.Dims.Height > 0 Then
 
 'Allocate memory to the Refresh buffer
 Obj.Picture = CreatePicture(Surf.Dims.Width, Surf.Dims.Height, 32)
 GetObjectAPI Obj.Picture, Len(BM), BM
 With Surf.SA1D
 .cbElements = 4
 .cDims = 1
 .lLbound = 0
 .cElements = BM.bmHeight * BM.bmWidth
 .pvData = BM.bmBits
 End With
 CopyMemory ByVal VarPtrArray(Surf.Dib), VarPtr(Surf.SA1D), 4
 
 With Surf.SA1D_L
 .cbElements = 4
 .cDims = 1
 .lLbound = 0
 .cElements = BM.bmHeight * BM.bmWidth
 .pvData = BM.bmBits
 ReDim Surf.EraseDib(.cElements)
 End With
 CopyMemory ByVal VarPtrArray(Surf.LDib), VarPtr(Surf.SA1D_L), 4
 
 Surf.CBWidth = BM.bmWidthBytes
 End If
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
bRunning = False
End Sub

Public Function RealRound(ByVal sngValue!) As Long
 'This function rounds .5 up
 
 RealRound = Int(sngValue)
 If sngValue - RealRound >= 0.5! Then RealRound = RealRound + 1&

End Function
Public Function RealRound2(ByVal sngValue!) As Long
Dim diff!
 'This function rounds .5 down
 
 RealRound2 = Int(sngValue)
 diff = sngValue - RealRound2
 If diff > 0.5! Then RealRound2 = RealRound2 + 1&

End Function

Private Sub ShowPage()

Cls

Select Case PageIndex
Case 0
 cmdBack.Enabled = False
 Airbrush.definition = 1
 Airbrush.Pressure = 255
 tmr1.Enabled = False
 
 ceil_y = 15
 cone_height = 50
 cone_diam = 100
 
 base_y = ceil_y + cone_height
 left_x = ceil_y / 2
 
 right_x = left_x + cone_diam
 mid_x = (right_x + left_x) / 2
 
 cone_base = ceil_y + cone_height
 
 ''box
 Line (left_x, ceil_y)-(right_x, base_y), , B
 
 'cone
 Line (left_x, cone_base)-(mid_x, ceil_y)
 Line (right_x, cone_base)-(mid_x, ceil_y)
 
 CurrentY = 70
 Print
 Print "The first thing I do when visualizing the airbrush shape"
 Print "is imagine a cone."
 
Case 1
 tmr1.Enabled = True
 Airbrush.definition = 1
 Print " Brush Pressure"
 
 cone_base = ceil_y + cone_height
 sPress = (Cos(sineParm1 * 5) / 2 + 0.5)
 cone_top = base_y - cone_height * sPress
 
 sPress = sPress * B255
 Airbrush.Pressure = sPress
 
 ''box
 Line (left_x, ceil_y)-(right_x, ceil_y), vbWhite
 ForeColor = vbWhite
 Print " Max = 255"
 ForeColor = 0
 Line (left_x, cone_base)-(right_x, cone_top), , B
 
 'cone
 Line (left_x, cone_base)-(mid_x, cone_top)
 Line (right_x, cone_base)-(mid_x, cone_top)
 
 CurrentY = 70
 Print
 Print "Brush 'Pressure' is a simple matter of multiplication."

Case 2
 tmr1.Enabled = True
 DefinitionFrame Cos(sineParm1 * 5) + 1, 100

 Print
 Print "Brush 'definition' has an additional multiplicative effect,"
 Print "where the intended response of the brush has a 'cap'"
 Print "when definition is greater than 1.0"

 'Print "Note how I dropped Pressure in this example to give"
 'Print "you an idea of how the two 'properties' work together."

Case 3
 tmr1.Enabled = False
 Airbrush.Pressure = 255
 Airbrush.definition = 2
 DefinitionFrame Airbrush.definition, Airbrush.Pressure, 0.5
 Print "First, start off with a Type, like this."
 Print "I am leaving out other properties like"
 Print "Red, Green, and Blue for simplicity."
 ForeColor = BrightPine
 Print "'============================="
 ForeColor = 0
 Print
 PrintBrushHeightInfo2
 Print
 DefaultsInfo
 
Case 4
 Airbrush.definition = 2
 DefinitionFrame Airbrush.definition, Airbrush.Pressure, 0.5
 Print
 Print "This is what I meant by multiplication with"
 Print "respect to Pressure and definition."
 Print
 Print
 Print "Private sub BlitAirbrush()"
 ForeColor = BrightPine
 Print
 ForeColor = vbBlue
 Print " brush_height ";: ForeColor = 0: Print "= ";
 Print "Airbrush.Pressure * Airbrush.definition"
 Print " .."
 Print
 Print
 Print "End Sub"
 Line (mid_x, cone_top)-(mid_x, cone_base), vbBlue
 tmr1.Enabled = False
 
Case 5
 bRegularDistFormula = False
 BehindTheScenesInfo True
 tmr1.Enabled = True
 ForeColor = BrightPine
 Print
 Print " 'Remember  slope = rise / run ?"
 ForeColor = 0
 Print
 Print " .."
 Print "End Sub"
 
Case 6
 tmr1.Enabled = False
 bRegularDistFormula = True
 Airbrush.diameter = 100
 Airbrush.Pressure = 255
 Airbrush.definition = 1
 DeltaPage
 Picture1.PSet (5, 5), vbBlack
 
Case 7
 bRegularDistFormula = False
 tmr1.Enabled = True
 InfoPixelCenters vbWhite * Int(Rnd + 0.5), True
 Airbrush.diameter = 0.2

Case 8
 bRegularDistFormula = False
 InfoPixelCenters vbBlack, False
 cone_diam = 4 * Cos(sineParm1) + 4
 Airbrush.diameter = cone_diam
 Airbrush.Pressure = 255
 Airbrush.definition = 2
 blit_x = PicDib.halfW
 BrushBoxPage cone_diam, vbWhite
 
Case 9
 Delta_BaseLeftPage 145, 100, 20, 3, 4, 2
 bRegularDistFormula = True
 ForeColor = BrightPine
 Print
 Print " 'Notice what is happening with the blue, first."
 ForeColor = vbBlue
 Print " delta_left ";: ForeColor = 0
 Print "= (";: ForeColor = UserComment2:
 Print "DrawLeft - x";: ForeColor = 0
 Print ") * brush_slope"
 Print
 Print
 Print " For DrawY = DrawBot To DrawTop"
 Print "  delta_x = ";: ForeColor = vbBlue: Print "delta_left"
 Print
 ForeColor = 0
 Print "  For DrawX = DrawLeft To DrawRight"
 Print
 ForeColor = BrightPine
 Print "  ...": Print "  ..."
 Print "  ..." ': Print
 
 ForeColor = BrightPine
 Print '"  'again, starts negative, moves toward zero, then positive"
 ForeColor = 0
 Print "   delta_x ";: ForeColor = 0: Print "= ";
 Print "delta_x";: ForeColor = 0: Print " + brush_slope"
 ForeColor = 0
 Print "  Next DrawX"
 ForeColor = BrightPine
 Print " ..."
 
 Airbrush.diameter = 8
 Airbrush.Pressure = 255
 Airbrush.definition = 1
 blit_x = PicDib.halfW + 2 * Cos(sineParm1)
 
Case 10

 CurrentY = 108
 ForeColor = BrightPine
 Print "  'Here is the initial delta_y, and the outer For-Next loop."
 ForeColor = vbBlue
 Print " delta_y ";: ForeColor = 0: Print "= (DrawBot - y) * brush_slope"
 Print
 Print " For DrawY = DrawBot To DrawTop" ' Step Surf.Dims.Width"
 Print: Print: Print
 ForeColor = BrightPine
 Print "  ...": Print "  ..."
 Print "  ...": Print
 ForeColor = vbBlue
 Print: Print
 Print
 Print "  delta_y ";: ForeColor = 0:
 Print "= ";: ForeColor = vbBlue: Print "delta_y ";: ForeColor = 0
 Print "+ brush_slope"
 Print " Next DrawY"

Case 11
 bRegularDistFormula = False
 Airbrush.diameter = 8
 Airbrush.Pressure = 255
 Airbrush.definition = 1
 CurrentY = 69
 Print "  And this pretty much wraps it up for the solid brush."
 Print "  Until next time  :)"
 Print
 Print " delta_left ";: ForeColor = 0
 Print "= (DrawLeft - x) * brush_slope"
 Print " delta_y = (DrawBot - y) * brush_slope"
 Print
 Print " For DrawY = DrawBot To DrawTop"
 Print "  delta_ySq = delta_y * delta_y"
 Print "  delta_x = delta_left"
 ForeColor = 0
 Print "  For DrawX = DrawLeft To DrawRight"
 Print "   Bright = brush_height - Sqr(delta_x * delta_x + delta_ySq)"
 Print "   If Bright > 0 Then"
 Print "    If Bright > Airbrush.Pressure Then Bright = Airbrush.Pressure"
 Print "    .. ";: ForeColor = BrightPine: Print "'regular Alpha-blending stuff"
 ForeColor = 0
 Print "   End If"
 Print "   delta_x = delta_x + brush_slope"
 Print "  Next DrawX"
 Print "  delta_y = delta_y + brush_slope"
 Print " Next DrawY"
 cmdNext.Enabled = False
 
End Select
End Sub
Private Sub PrintBrushHeightInfo2()
 Print "Private Type AirbrushStruct"
 ForeColor = 0
 Print " diameter As Single"
 Print " Pressure As Byte"
 Print " definition As Single"
 Print "End Type"
 Print
 Print "Dim Airbrush as AirbrushStruct"
End Sub
Private Sub BrushDiameterInfo()
 Print "Private Type AirbrushStruct"
 Print " diameter ";
 Print "As Single"
 ForeColor = 0
 Print " Pressure As Byte"
 Print " definition As Single"
 Print "End Type"
 Print
End Sub
Private Sub DefaultsInfo()
 Print "Private Sub Form_Load"
 Print " Airbrush.diameter = 100";: ForeColor = BrightPine
 Print "  'Assign some defaults!": ForeColor = 0
 Print " Airbrush.Pressure = " & Airbrush.Pressure
 Print " Airbrush.definition = " & Airbrush.definition
 Print "End Sub"
 tmr1.Interval = 50
 tmr1.Enabled = False
End Sub
Private Sub BehindTheScenesInfo(Optional Blinking As Boolean)
Dim BGR1&

 DefinitionFrame 1, 255
 Print
 Print "Private Sub BlitAirbrush()"
 Print
 ForeColor = BrightPine
 ForeColor = 0
 Print " brush_height = Airbrush.Pressure * Airbrush.definition"
 Print " brush_radius = Airbrush.diameter / 2"
 If Blinking And Rnd > 0.5 Then BGR1 = vbBlue
 ForeColor = BGR1
 Print " brush_slope";: ForeColor = 0
 Print " = brush_height / brush_radius"
 XS1 = CurrentX
 YS1 = CurrentY
 Line (left_x, cone_base)-(mid_x, cone_top), BGR1
 CurrentX = XS1
 CurrentY = YS1
 Airbrush.diameter = 100
End Sub
Private Sub DeltaPage()
 Print "The brush looks like this (at least with my method)"
 Print "unless I put a 'flip' in the generated alpha."
 Print
 Print "I compute 'every' pixel's Alpha using distance formula,"
 Print "where delta x means any pixel's center x distance from"
 Print "brush center x .."
 Print
 Print "So you can see how, when relying on the distance for-"
 Print "mula as I do, you will get an increased brightness as"
 Print "you move away from brush center as shown."
 Print
 Print "This is how I make up for that"
 Print " Alpha = brush_height - Sqr#(delta_x * delta_x + delta_y * delta_y)"
End Sub
Private Sub InfoPixelCenters(CenterPixelColor As Long, bool_ShowGridDots As Boolean)
 
 RYS = CurrentY + 10
 RXS = CurrentX
 
 RXE = RXS + 200
 RYE = RYS + 200
 
 Grid_Magnification = 10
 
 VertGrid RXS, RXE, RYS, RYE, Grid_Magnification
 HorizGrid RXS, RXE, RYS, RYE, Grid_Magnification

 If bool_ShowGridDots Then
  GridDots RXS, RXE, RYS, RYE, Grid_Magnification, CenterPixelColor
  LColorSave = ForeColor
  ForeColor = CenterPixelColor
  Print
  Print " Pixel Centers"
  ForeColor = LColorSave
  Print
  Print "Here is how I begin to visualize deltas x and y of a pixel"
  Print "and the brush center."
 End If
 
End Sub
Private Sub BrushCenterPage(PrintColor As Long, bool_qBrushCenter As Boolean)
 LColorSave = ForeColor
 ForeColor = PrintColor
 pageBrushCenterX = RXS + 103
 pageBrushCenterY = RYS + 102
 PSet (pageBrushCenterX, pageBrushCenterY)
 If bool_qBrushCenter Then
  CurrentY = CurrentY - 5
  CurrentX = CurrentX + 5
  Print "Brush Center"
 End If
 ForeColor = LColorSave
End Sub
Private Sub BrushBoxPage(brush_radi!, SquareColor&)
 pageRadius = brush_radi * Grid_Magnification
 BrushCenterPage vbBlack, False
 LColorSave = ForeColor
 ForeColor = SquareColor
 PSet (pageBrushCenterX, pageBrushCenterY)
 RXS = RealRound(pageBrushCenterX - pageRadius)
 RXE = RealRound(pageBrushCenterX) '- 6
 Circle (pageBrushCenterX, pageBrushCenterY), pageRadius
 Line (RXS, pageBrushCenterY)- _
      (RXE, pageBrushCenterY), , B
 CurrentX = CurrentX + 2
 CurrentY = CurrentY
 ForeColor = LColorSave
 
 RYS = 10
 RXS = RealRound(pageBrushCenterX) - Grid_Magnification * RealRound(brush_radi) - 1
 SimTransparentPixelColumn RealRound(RXS) + 3, 1, RYS, 200, vbBlue, 255
 
 Print
 RXS = CurrentX
 Print "And with this, I compute which column of pixels has the"
 CurrentX = RXS
 Print "left side of the brush"
 Print
 Print "Private Sub BlitAirbrush(";: ForeColor = vbMagenta
 Print "x ";: ForeColor = 0: Print "As Single, y ...)"
 CurrentX = RXS
 Print "  .. ";: ForeColor = BrightPine: Print Tab(21); "'RealRound is my rounding function"
 ForeColor = vbBlue
 CurrentX = RXS
 Print "  DrawLeft ";: ForeColor = 0: Print "= RealRound(";: ForeColor = vbMagenta
 Print "x ";: ForeColor = 0: Print "- brush_radius)"
 CurrentX = RXS
 ForeColor = BrightPine
 ForeColor = 0

End Sub
Private Sub SimTransparentPixelColumn(X1&, XWidth&, Y1&, YHeight&, RectColor&, BytAlpha&)
Dim X2&
Dim Y2&

 X2 = X1 + XWidth - 1
 Y2 = Y1 + YHeight - 1
 
 For DrawY = Y1 To Y2
  For DrawX = X1 To X2
   BGR = Point(DrawX, DrawY)
   BGRed = BGR And &HFF
   BGGreen = (BGR And 65280) / 512&
   BGBlue = (BGR And 16711680) / 131072
   FGRed = RectColor And &HFF
   FGGreen = (RectColor And 65280) / 512&
   FGBlue = (RectColor And 16711680) / 131072
   PSet (DrawX, DrawY), RGB(BGRed / 2 + FGRed / 2, _
                          BGGreen + FGGreen, _
                          BGBlue + FGBlue)
  Next
 Next

End Sub
Private Sub Delta_BaseLeftPage(X_Center_Base&, Col_Base&, Grid_Scale&, brush_height!, brush_radius!, sineRadius!)
Dim brush_center_x!
Dim gen_proj_len!
Dim gen_x_sineRad!
Dim proj_left_edge!
Dim scaleRadius!
Dim RPLen&
Dim RCX&
Dim initialBright!
 
 gen_proj_len = Cos(sineParm1)
 gen_x_sineRad = gen_proj_len * sineRadius
 brush_center_x = X_Center_Base + gen_x_sineRad
 scaleRadius = Grid_Scale * brush_radius
 
 RCX = RealRound(brush_center_x)
 
 RPLen = RCX - X_Center_Base
 
 RXL = X_Center_Base + Grid_Scale * RPLen - scaleRadius
 
 RYS = Col_Base - 50
 
 CurrentY = RYS - 40
 CurrentX = 10
 
 ForeColor = 0
 
 'add to DrawX to have PSet repr. horizontal pixL centers
 RXE = RealRound(Grid_Scale / 2)
 
 RXS = X_Center_Base - RXE
 
 'Base of the columns
 Line (RXS, Col_Base)-(RXE, Col_Base), RGB(150, 150, 150)
 
 brush_center_x = X_Center_Base + gen_x_sineRad * Grid_Scale
 
 Line (brush_center_x, Col_Base)-(RXL, Col_Base), UserComment2
 
 For DrawX = RXS To RXE Step -Grid_Scale
  Line (DrawX, Col_Base)-(DrawX, RYS)
  PSet (DrawX + RXE, Col_Base), vbWhite
 Next
 
 proj_left_edge = brush_center_x - scaleRadius
 
 '"cone" tip y
 RYS = Col_Base - Grid_Scale * brush_height
 
 ForeColor = vbWhite
 Line (brush_center_x, Col_Base)-(brush_center_x, RYS)
 Line (brush_center_x, RYS)-(proj_left_edge, Col_Base)
 
 ForeColor = vbBlue
 
 initialBright = Grid_Scale * brush_height * (RXL - proj_left_edge) / _
                 (brush_center_x - proj_left_edge)
                 
 Line (RXL, Col_Base)-(RXL, Col_Base - initialBright)
 
 CurrentY = Col_Base - 5
 Print
End Sub

Private Sub VertGrid(ByVal X1&, ByVal X2&, ByVal Y1&, ByVal Y2&, ByVal XSTP&, Optional BGR& = Gray128)
For RX = X1 To X2 Step XSTP&
 Line (RX, Y1)-(RX, Y2), BGR
Next
End Sub
Private Sub HorizGrid(ByVal X1&, ByVal X2&, ByVal Y1&, ByVal Y2&, ByVal YSTP&, Optional BGR& = Gray128)
For RY = Y1 To Y2 Step YSTP&
 Line (X1, RY)-(X2, RY), BGR
Next
End Sub
Private Sub GridDots(ByVal sngX1!, ByVal sngX2!, ByVal sngY1!, ByVal sngY2!, ByVal sStep!, Optional BGR& = vbWhite)
Dim stepDiv2!
Dim sngX!
Dim sngY!

stepDiv2 = sStep / 2

sngX1 = sngX1 + stepDiv2 - 0.001
sngX2 = sngX2 - stepDiv2
sngY1 = sngY1 + stepDiv2 - 0.001
sngY2 = sngY2 - stepDiv2
For sngY = sngY1 To sngY2 Step sStep
 For sngX = sngX1 To sngX2 Step sStep
 PSet (sngX, sngY), BGR
 Next
Next

End Sub
Private Sub DefinitionFrame(sDef1 As Single, ByVal chPress As Byte, Optional box_height_scale As Single = 1)
 tmr1.Enabled = True
 
 Airbrush.definition = sDef1
 
 Print " Brush definition = " & Round(Airbrush.definition, 1)
 cone_base = ceil_y + cone_height
 sPress = chPress / B255
 
 box_ceil = cone_base - cone_height * chPress * box_height_scale / B255
 
 cone_top = base_y - cone_height * sPress * Airbrush.definition * box_height_scale
 
 sPress = sPress * B255
 Airbrush.Pressure = chPress
 
 ''box
 Line (left_x, box_ceil)-(right_x, box_ceil), vbWhite
 ForeColor = vbWhite
 Print " Pressure = "; chPress; ""
 ForeColor = 0
 Line (left_x, cone_base)-(right_x, cone_top), , B
 
 'cone
 Line (left_x, cone_base)-(mid_x, cone_top)
 Line (right_x, cone_base)-(mid_x, cone_top)
  
 CurrentY = cone_base

 Print

End Sub
Private Sub ColorShift()

 If sR < sB Then
  If sR < sG Then
   If sG < sB Then
    bytMaxMin_diff = sB - sR
    sG = sG - bytMaxMin_diff * rgb_shift_intensity
    If sG < sR Then
     iSubt = sR - sG
     sG = sR
     sR = sR + iSubt
    End If
   Else
    bytMaxMin_diff = sG - sR
    sB = sB + bytMaxMin_diff * rgb_shift_intensity
    If sB > sG Then
     iSubt = sB - sG
     sB = sG
     sG = sG - iSubt
    End If
   End If
  Else
   bytMaxMin_diff = sB - sG
   sR = sR + bytMaxMin_diff * rgb_shift_intensity
   If sR > sB Then
    iSubt = sR - sB
    sR = sB
    sB = sB - iSubt
   End If
  End If
 ElseIf sR > sG Then
  If sB < sG Then
   bytMaxMin_diff = sR - sB
   sG = sG + bytMaxMin_diff * rgb_shift_intensity
   If sG > sR Then
    iSubt = sG - sR
    sG = sR
    sR = sR - iSubt
   End If
  Else
   bytMaxMin_diff = sR - sG
   sB = sB - bytMaxMin_diff * rgb_shift_intensity
   If sB < sG Then
    iSubt = sG - sB
    sB = sG
    sG = sG + iSubt
   End If
  End If
 Else
  bytMaxMin_diff = sG - sB
  sR = sR - bytMaxMin_diff * rgb_shift_intensity
  If sR < sB Then
   iSubt = sB - sR
   sR = sB
   sB = sB + iSubt
  End If
 End If
 
End Sub

Private Sub tmr1_Timer()
 ShowPage
End Sub
