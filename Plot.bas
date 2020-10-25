Attribute VB_Name = "PlotGraphics"
Option Explicit

'API declarations font handling
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * 32
  'lfFaceName(1 To LF_FACESIZE) As Byte 'original type declaration from API Viewer
End Type

Public Enum FontWeights
  FW_DONTCARE = 0
  FW_THIN = 100
  FW_EXTRALIGHT = 200
  FW_LIGHT = 300
  FW_NORMAL = 400
  FW_MEDIUM = 500
  FW_SEMIBOLD = 600
  FW_BOLD = 700
  FW_EXTRABOLD = 800
  FW_HEAVY = 900
End Enum

Public Enum FontCharSets
  ANSI_CHARSET = 0
  DEFAULT_CHARSET = 1
  SYMBOL_CHARSET = 2
  MAC_CHARSET = 77
  SHIFTJIS_CHARSET = 128
  HANGEUL_CHARSET = 129
  CHINESEGIG5_CHARSET = 136
  OEM_CHARSET = 255
End Enum

Public Enum FontOutPrecisions
  OUT_DEFAULT_PRECIS = 0
  OUT_STRING_PRECIS = 1
  OUT_CHARACTER_PRECIS = 2
  OUT_STROKE_PRECIS = 3
  OUT_TT_PRECIS = 4
  OUT_DEVICE_PRECIS = 5
  OUT_RASTER_PRECIS = 6
  OUT_TT_ONLY_PRECIS = 7
  OUT_OUTLINE_PRECIS = 8
End Enum

Public Enum FontClipPrecisions
  CLIP_DEFAULT_PRECIS = 0
  CLIP_CHARACTER_PRECIS = 1
  CLIP_STROKE_PRECIS = 2
  CLIP_LH_ANGLES = 16
  CLIP_TT_ALWAYS = 32
  CLIP_EMBEDDED = 128
  CLIP_TO_PATH = 4097
End Enum

Public Enum FontQuality
  DEFAULT_QUALITY = 0
  DRAFT_QUALITY = 1
  PROOF_QUALITY = 2
End Enum

Public Enum FontPitch
  DEFAULT_PITCH = 0
  FIXED_PITCH = 1
  VARIABLE_PITCH = 2
End Enum

Public Enum FontFamily
  FF_DONTCARE = 0
  FF_ROMAN = 16
  FF_SWISS = 32
  FF_MODERN = 48
  FF_SCRIPT = 64
  FF_DECORATIVE = 80
End Enum


'public constants used in demo-forms
Public dPlot() As Double
Public udtMyGraphLayout As GRAPHIC_LAYOUT

'declaration of UDT's (User Defined Types)
Public Type GRAPHIC_LAYOUT
  XTitle As String 'title X-axis
  Ytitle As String 'title Y-axis
  blnOrigin As Boolean 'origin is included for only pos/neg values when true
  blnGridLine As Boolean 'Gridlines are shown when true
  lStart As Long 'index of start x-Range
  lEnd As Long 'index of end x-Range
  asX As Double 'trace in array to function as "X-value"
  asY() As Variant 'Y-traces to plot
  DrawTrace() As DRAWN_AS
  X0 As Double 'minimum value of domain X-values to draw
  X1 As Double 'maximum value of domain X-values to draw
  Y0 As Double 'minimum value of domain Y-values to draw
  Y1 As Double 'maximum value of domain Y-values to draw
End Type

Public Enum DRAWN_AS
  AS_POINT
  AS_CONLINE
  AS_BAR
End Enum

Public Type COORDINATE
  X As Single
  Y As Single
End Type
  

'public declaration of screen variables - in twips
Public twp_XLeftMargin As Single 'left margin
Public twp_XRightMargin As Single 'right margin
Public twp_YTopMargin As Single 'top margin
Public twp_YBottomMargin As Single 'bottom margin
Public twp_YRange As Single 'full Y-Range
Public twp_XRange As Single 'full X-Range

'public declaration of value variables - in their own units
Public val_XMin As Double 'minimum value X
Public val_XRange As Double 'full X-Range X-values
Public val_YMin As Double 'minimum value Y
Public val_YRange As Double 'full Y-Range Y-values

Public Sub Plot(frmSpec As Form, dArSpec() As Double, udtLayoutSpec As GRAPHIC_LAYOUT)

'declaration of screen variables - in twips
Dim twp_YTick As Single 'size of Y tick
Dim twp_XTick As Single 'size of X tick
Dim twp_Y0 As Single 'Y0: origin (as in 1st quadrant)
Dim twp_X0 As Single 'X0: origin (as in 1st quadrant)
Dim twp_Y0Tr As Single 'transferrred Y0: origin (as in 4 quadrants)
Dim twp_X0Tr As Single 'transferrred X0: origin (as in 4 quadrants)


'declaration of value variables - in their own units
Dim val_XMax As Double 'maximum value X
Dim val_YMax As Double 'maximum value Y
Dim val_X As Double 'value of X-value
Dim val_Y As Double 'value of Y-value

'declaration of dimensionless variables (ratios)
Dim XRatio As Double 'quotient of val_X and val_XRange
Dim YRatio As Double 'quotient of val_Y and val_YRange
Dim NumYTicks As Integer 'number of ticks Y-axis
Dim NumXTicks As Integer 'number of ticks X-axis

'declaration of general variables
Dim nI As Integer 'counter
Dim nTrace As Variant 'the traces to be plotted
Dim clr_Plot(10) As Long 'array with colors
Dim udtFont As LOGFONT 'to create a logical font type
Dim lHandleFont As Long 'handle for new (logical) font
Dim lOldFont As Long 'handle of old font
Dim lRetVal As Long 'acts for storing return value


'*************  initialise screen  ******************
  'screen: clear and define drawwidth
  frmSpec.Cls
  frmSpec.DrawWidth = 1
  
  'screenRange, screenorigin and X/Y-ticks
  twp_XTick = 80
  twp_YTick = 80
  twp_XLeftMargin = 800
  twp_XRightMargin = 200
  twp_YTopMargin = 300
  twp_YBottomMargin = 500
  twp_Y0 = frmSpec.ScaleHeight - twp_YBottomMargin
  twp_X0 = twp_XLeftMargin
  twp_YRange = frmSpec.ScaleHeight - twp_YBottomMargin - twp_YTopMargin
  twp_XRange = frmSpec.ScaleWidth - twp_XLeftMargin - twp_XRightMargin
  
  'font (and colors)
  frmSpec.Font.Name = "MS Sans Serif"
  frmSpec.Font.Size = 8
  frmSpec.Font.Bold = False
  frmSpec.Font.Italic = False
  clr_Plot(0) = RGB(100, 20, 0)
  clr_Plot(1) = RGB(0, 0, 255)
  clr_Plot(2) = RGB(255, 0, 0)
  clr_Plot(3) = RGB(255, 255, 0)
  clr_Plot(4) = RGB(93, 255, 201)
  clr_Plot(5) = RGB(0, 255, 255)
  clr_Plot(6) = RGB(210, 25, 210)
  clr_Plot(7) = RGB(255, 255, 255)
  clr_Plot(8) = RGB(255, 0, 115)
  clr_Plot(9) = RGB(0, 0, 115)
  clr_Plot(10) = RGB(80, 80, 80)
  
  'logical font
  With udtFont
    .lfEscapement = 200
    .lfFaceName = "Arial" & Chr$(0)
    .lfHeight = (9 * -20) / Screen.TwipsPerPixelY
  End With
    

'*************  determine Xmin, Xmax, Ymin and Ymax  *************
'Xmin
val_XMin = udtLayoutSpec.X0
If val_XMin > 0 And udtLayoutSpec.blnOrigin = True Then
  val_XMin = 0
End If
'if val_Xmin<0 lower twp_XleftMargin to show more in window
If val_XMin < 0 Then
  twp_XLeftMargin = 600
  twp_X0 = twp_XLeftMargin
  twp_XRange = frmSpec.ScaleWidth - twp_XLeftMargin - twp_XRightMargin
End If

'Xmax
val_XMax = udtLayoutSpec.X1
If val_XMax < 0 And udtLayoutSpec.blnOrigin = True Then
  val_XMax = 0
End If
If val_XMax = val_XMin Then
  val_XMin = val_XMin - 1
  val_XMax = val_XMax + 1
End If
val_XRange = val_XMax - val_XMin

'Ymin
val_YMin = udtLayoutSpec.Y0
If val_YMin > 0 And udtLayoutSpec.blnOrigin = True Then
  val_YMin = 0
End If

'Ymax
val_YMax = udtLayoutSpec.Y1
If val_YMax < 0 And udtLayoutSpec.blnOrigin = True Then
  val_YMax = 0
End If
If val_YMax = val_YMin Then
  val_YMin = val_YMin - 1
  val_YMax = val_YMax + 1
End If
val_YRange = val_YMax - val_YMin
  

'*************  prepare SpanX and twp_X0Tr  *****************
'determine Pl_SpanX
Dim Pl_SpanX As Double 'span between two ticks in own units
Dim nExp As Integer 'help to determine Pl_SpanX and Pl_SpanY

nExp = 0
If (val_XMax - val_XMin) < 1 And (val_XMax - val_XMin) > 0 Then
  Do While (val_XMax - val_XMin) < 1
    nExp = nExp + 1
    val_XMax = val_XMax * 10
    val_XMin = val_XMin * 10
  Loop
  Pl_SpanX = 10 ^ (-nExp)
  val_XMax = val_XMax * 10 ^ (-nExp) 'correct val_Xmax to original value
  val_XMin = val_XMin * 10 ^ (-nExp) 'correct val_Xmin to original value
Else
  Pl_SpanX = 1
  Do While val_XRange / Pl_SpanX > 20
    If val_XRange / Pl_SpanX > 20 Then
      Pl_SpanX = Pl_SpanX * 2
    End If
    If val_XRange / Pl_SpanX > 20 Then
      Pl_SpanX = Pl_SpanX * 2.5
    End If
    If val_XRange / Pl_SpanX > 20 Then
      Pl_SpanX = Pl_SpanX * 2
    End If
  Loop
End If

'determine twp_X0Tr (Translated twp_X0; twp_X0 is position origin in twips)
If val_XMin < 0 And val_XMax > 0 Then 'positive and negative X-values
  twp_X0Tr = twp_X0 - (val_XMin / val_XRange) * twp_XRange ' "-" because of negative value val_Xmin!
ElseIf val_XMin < 0 And val_XMax <= 0 Then 'only negative values X-values
  twp_X0Tr = twp_X0 + twp_XRange 'axis at end of X-Range
Else
End If


'*************  prepare SpanY and twp_Y0Tr  *****************
'determine Pl_SpanY
Dim Pl_SpanY As Double 'span between two ticks in own units

nExp = 0
If (val_YMax - val_YMin) < 1 And (val_YMax - val_YMin) > 0 Then
  Do While (val_YMax - val_YMin) < 1
    nExp = nExp + 1
    val_YMax = val_YMax * 10
    val_YMin = val_YMin * 10
  Loop
  Pl_SpanY = 10 ^ (-nExp)
  val_YMax = val_YMax * 10 ^ (-nExp)
  val_YMin = val_YMin * 10 ^ (-nExp)
Else
  Pl_SpanY = 1
  Do While val_YRange / Pl_SpanY > 20
    If val_YRange / Pl_SpanY > 20 Then
      Pl_SpanY = Pl_SpanY * 2
    End If
    If val_YRange / Pl_SpanY > 20 Then
      Pl_SpanY = Pl_SpanY * 2.5
    End If
    If val_YRange / Pl_SpanY > 20 Then
      Pl_SpanY = Pl_SpanY * 2
    End If
  Loop
End If

'determine twp_Y0Tr (Translated twp_Y0; twp_Y0 is position origin in twips)
If val_YMin < 0 And val_YMax > 0 Then 'positive and negative Y-values
  twp_Y0Tr = twp_Y0 + (val_YMin / val_YRange) * twp_YRange
ElseIf val_YMin < 0 And val_YMax <= 0 Then 'only negative values Y-values
  twp_Y0Tr = twp_Y0 - twp_YRange
Else '1st quadrant is shown
End If


'************  plot Y-gridlines  (vertical)  **************

Dim Dummy1 As Double
Dim Dummy2 As Double
Dim OffSetX As Double
Dim twp_OffSetX As Double
Dim twp_XTickRange As Double

Dummy1 = Int(val_XMin / Pl_SpanX)
Dummy2 = Int(val_XMax / Pl_SpanX)
NumXTicks = Dummy2 - Dummy1
OffSetX = (val_XMin - Pl_SpanX * Int(val_XMin / Pl_SpanX)) 'offsetX in own units
twp_OffSetX = twp_XRange * OffSetX / val_XRange 'offsetX in twips
Dummy2 = (val_XMax - Pl_SpanX * Int(val_XMax / Pl_SpanX)) 'difference between val_Xmax and highest Ytick lable
Dummy2 = twp_XRange * Dummy2 / val_XRange 'and now in twips
twp_XTickRange = twp_XRange * ((twp_XRange + twp_OffSetX - Dummy2) / twp_XRange) / NumXTicks

If udtLayoutSpec.blnGridLine = True Then
  If val_XMin = Pl_SpanX * Int(val_XMin / Pl_SpanX) Then 'val_Xmin = gridline
    frmSpec.Line (twp_X0, twp_Y0)-(twp_X0, twp_Y0 - twp_YRange), &H80000016
  End If
  For nI = 1 To NumXTicks
    frmSpec.Line (twp_X0 - twp_OffSetX + nI * twp_XTickRange, twp_Y0)- _
    (twp_X0 - twp_OffSetX + nI * twp_XTickRange, twp_Y0 - twp_YRange), &H80000016
  Next nI
End If


'************  plot X-gridlines  (horizontal)  **************

Dim OffSetY As Double
Dim twp_OffSetY As Double
Dim twp_YTickRange As Double

Dummy1 = Int(val_YMin / Pl_SpanY)
Dummy2 = Int(val_YMax / Pl_SpanY)
NumYTicks = Dummy2 - Dummy1 - 1
If NumYTicks = 0 Then
  NumYTicks = 1
End If
OffSetY = Pl_SpanY - (val_YMin - Pl_SpanY * Int(val_YMin / Pl_SpanY))
twp_OffSetY = twp_YRange * OffSetY / val_YRange
Dummy2 = (val_YMax - Pl_SpanY * Int(val_YMax / Pl_SpanY))
Dummy2 = twp_YRange * Dummy2 / val_YRange
twp_YTickRange = (twp_YRange - twp_OffSetY - Dummy2) / NumYTicks
If (Int(val_YMax / Pl_SpanY) - Int(val_YMin / Pl_SpanY)) = 1 Then
  NumYTicks = 0
End If 'otherwise labeling incorrect

If udtLayoutSpec.blnGridLine = True Then 'plot gridlines
  If val_YMin = Pl_SpanY * Int(val_YMin / Pl_SpanY) Then 'val_Ymin = gridline
    frmSpec.Line (twp_X0, twp_Y0)-(twp_X0 + twp_XRange, twp_Y0), &H80000016
  End If
  For nI = 0 To NumYTicks 'rest of gridlines
    frmSpec.Line (twp_X0, twp_Y0 - twp_OffSetY - nI * twp_YTickRange)- _
    (twp_X0 + twp_XRange, twp_Y0 - twp_OffSetY - nI * twp_YTickRange), &H80000016
  Next nI
End If
  
  
'*************  plot datapoints for every trace  *******************
Dim nNumTraces As Integer

nNumTraces = 0
For Each nTrace In udtLayoutSpec.asY()
  nNumTraces = nNumTraces + 1
  Select Case udtLayoutSpec.DrawTrace(nNumTraces - 1)
  
  Case AS_CONLINE
    'find value starting point
    val_X = dArSpec(udtLayoutSpec.lStart, udtLayoutSpec.asX)
    val_Y = dArSpec(udtLayoutSpec.lStart, nTrace)
    XRatio = (val_X - val_XMin) / val_XRange
    YRatio = (val_Y - val_YMin) / val_YRange
    frmSpec.CurrentX = twp_X0 + XRatio * twp_XRange
    frmSpec.CurrentY = twp_Y0 - YRatio * twp_YRange
    'find rest
    For nI = udtLayoutSpec.lStart + 1 To udtLayoutSpec.lEnd Step 1
      val_X = dArSpec(nI, udtLayoutSpec.asX)
      val_Y = dArSpec(nI, nTrace)
      XRatio = (val_X - val_XMin) / val_XRange
      YRatio = (val_Y - val_YMin) / val_YRange
      frmSpec.Line -(twp_X0 + XRatio * twp_XRange, twp_Y0 - YRatio * twp_YRange), clr_Plot(nNumTraces Mod 10)
    Next nI
    'clear lines outside drawing pane
    frmSpec.Line (twp_X0, twp_Y0 - twp_YRange - 10)-(twp_X0 + twp_XRange, 0), frmSpec.BackColor, BF 'above drawing pane
    frmSpec.Line (twp_X0, twp_Y0 + 20)-(frmSpec.Width, frmSpec.Height), frmSpec.BackColor, BF 'below drawing pane
    frmSpec.Line (twp_X0 + twp_XRange + 20, frmSpec.Height)-(frmSpec.Width, 0), frmSpec.BackColor, BF 'below drawing pane
    frmSpec.Line (0, 0)-(twp_XLeftMargin - 20, frmSpec.Height), frmSpec.BackColor, BF 'below drawing pane

  Case AS_BAR
    For nI = udtLayoutSpec.lStart To udtLayoutSpec.lEnd Step 1
      val_X = dArSpec(nI, udtLayoutSpec.asX)
      val_Y = dArSpec(nI, nTrace)
      XRatio = (val_X - val_XMin) / val_XRange
      If XRatio >= 0 And XRatio <= 1 Then
        YRatio = (val_Y - val_YMin) / val_YRange
        If YRatio > 1 Then YRatio = 1
        If YRatio < 0 Then YRatio = 0
        If val_YMin >= 0 Then
          frmSpec.Line (twp_X0 + XRatio * twp_XRange, twp_Y0)- _
          (twp_X0 + XRatio * twp_XRange, twp_Y0 - YRatio * twp_YRange), clr_Plot(nNumTraces Mod 10)
        Else
          frmSpec.Line (twp_X0 + XRatio * twp_XRange, twp_Y0Tr)- _
          (twp_X0 + XRatio * twp_XRange, twp_Y0 - YRatio * twp_YRange), clr_Plot(nNumTraces Mod 10)
        End If
      End If
    Next nI
  
  Case AS_POINT
    frmSpec.DrawWidth = 2
    For nI = udtLayoutSpec.lStart To udtLayoutSpec.lEnd Step 1
      val_X = dArSpec(nI, udtLayoutSpec.asX)
      val_Y = dArSpec(nI, nTrace)
      XRatio = (val_X - val_XMin) / val_XRange
      YRatio = (val_Y - val_YMin) / val_YRange
      frmSpec.ForeColor = clr_Plot(nNumTraces Mod 10)
      If XRatio >= 0 And XRatio <= 1 And YRatio >= 0 And YRatio <= 1 Then
        frmSpec.PSet (twp_X0 + XRatio * twp_XRange, twp_Y0 - YRatio * twp_YRange)
      End If
      frmSpec.ForeColor = vbBlack
    Next nI
    frmSpec.DrawWidth = 1
  End Select
Next nTrace

  
'*************  plot ticks Y-axis  ************
If val_XMin < 0 Then
  If val_YMin = Pl_SpanY * Int(val_YMin / Pl_SpanY) Then 'val_Ymin = gridline + tick
    frmSpec.Line (twp_X0Tr - twp_XTick, twp_Y0)-(twp_X0Tr, twp_Y0), vbBlack
  End If
  For nI = 0 To NumYTicks
    frmSpec.Line (twp_X0Tr - twp_XTick, twp_Y0 - twp_OffSetY - nI * twp_YTickRange)- _
    (twp_X0Tr, twp_Y0 - twp_OffSetY - nI * twp_YTickRange), vbBlack
  Next nI
Else
  If val_YMin = Pl_SpanY * Int(val_YMin / Pl_SpanY) Then 'val_Ymin = gridline + tick
    frmSpec.Line (twp_X0 - twp_XTick, twp_Y0)-(twp_X0, twp_Y0), vbBlack
  End If
  For nI = 0 To NumYTicks
    frmSpec.Line (twp_X0 - twp_XTick, twp_Y0 - twp_OffSetY - nI * twp_YTickRange)- _
    (twp_X0, twp_Y0 - twp_OffSetY - nI * twp_YTickRange), vbBlack
  Next nI
End If


'************  plot labels to ticks from Y-axis  *************
Dim nLenYLable As Integer 'length of lable Y-axis ticks

frmSpec.ForeColor = clr_Plot(0)
If val_XMin < 0 Then
  If val_YMin = Pl_SpanY * Int(val_YMin / Pl_SpanY) Then 'val_Ymin = gridline + lable
    If Abs(val_YMax - val_YMin) > 100000 Or Abs(val_YMax - val_YMin) < 0.00001 Then
      frmSpec.CurrentX = twp_X0Tr - (Len(Format(val_YMin + Pl_SpanY * nI, "Scientific")) + 2) * twp_XTick
      frmSpec.CurrentY = twp_Y0 - twp_YTick
      frmSpec.Print Format(val_YMin, "Scientific")
    Else
      nLenYLable = Len(Trim(Str$(Abs(val_YMin))))
      frmSpec.CurrentX = twp_X0Tr - twp_XLeftMargin + (5 - nLenYLable) * twp_XTick
      frmSpec.CurrentY = twp_Y0 - twp_YTick
      frmSpec.Print val_YMin
    End If
  End If 'plot lable for val_Ymin = gridline
  For nI = 0 To NumYTicks
    If Abs(val_YMax - val_YMin) > 100000 Or Abs(val_YMax - val_YMin) < 0.00001 Then
      frmSpec.CurrentX = twp_X0Tr - (Len(Format(val_YMin + Pl_SpanY * nI, "Scientific")) + 2) * twp_XTick
      frmSpec.CurrentY = twp_Y0 - twp_OffSetY - nI * twp_YTickRange - twp_YTick
      frmSpec.Print Format(val_YMin + OffSetY + Pl_SpanY * nI, "Scientific")
    Else
      nLenYLable = Len(Trim(Str$(Abs(val_YMin + OffSetY + Pl_SpanY * nI))))
      frmSpec.CurrentX = twp_X0Tr - twp_XLeftMargin + (5 - nLenYLable) * twp_XTick
      frmSpec.CurrentY = twp_Y0 - twp_OffSetY - nI * twp_YTickRange - twp_YTick
      frmSpec.Print val_YMin + OffSetY + Pl_SpanY * nI
    End If
  Next nI
Else
  If val_YMin = Pl_SpanY * Int(val_YMin / Pl_SpanY) Then 'val_Ymin = gridline + lable
    If Abs(val_YMax - val_YMin) > 100000 Or Abs(val_YMax - val_YMin) < 0.00001 Then
      frmSpec.CurrentX = twp_X0 - (Len(Format(val_YMin + Pl_SpanY * nI, "Scientific")) + 2) * twp_XTick
      frmSpec.CurrentY = twp_Y0 - twp_YTick
      frmSpec.Print Format(val_YMin, "Scientific")
    Else
      nLenYLable = Len(Trim(Str$(Abs(val_YMin))))
      frmSpec.CurrentX = twp_X0 - twp_XLeftMargin + (7 - nLenYLable) * twp_XTick
      frmSpec.CurrentY = twp_Y0 - twp_YTick
      frmSpec.Print val_YMin
    End If
  End If 'plot lable for val_Ymin = gridline
  For nI = 0 To NumYTicks
    If Abs(val_YMax - val_YMin) > 100000 Or Abs(val_YMax - val_YMin) < 0.00001 Then
      frmSpec.CurrentX = twp_X0 - (Len(Format(val_YMin + Pl_SpanY * nI, "Scientific")) + 2) * twp_XTick
      frmSpec.CurrentY = twp_Y0 - twp_OffSetY - nI * twp_YTickRange - twp_YTick
      frmSpec.Print Format(val_YMin + OffSetY + Pl_SpanY * nI, "Scientific")
    Else
      nLenYLable = Len(Trim(Str$(Abs(val_YMin + OffSetY + Pl_SpanY * nI))))
      frmSpec.CurrentX = twp_X0 - twp_XLeftMargin + (7 - nLenYLable) * twp_XTick
      frmSpec.CurrentY = twp_Y0 - twp_OffSetY - nI * twp_YTickRange - twp_YTick
      frmSpec.Print val_YMin + OffSetY + Pl_SpanY * nI
    End If
  Next nI
End If
frmSpec.ForeColor = vbBlack


'**********  plot Y-axis and title  ***********
If val_XMin < 0 And val_XMax > 0 Then 'all four quadrants are shown
  frmSpec.Line (twp_X0Tr, twp_Y0 + twp_YTick)-(twp_X0Tr, twp_Y0 - twp_YRange), vbBlack 'Y-axis
  'prepare position title Y-axis and plot
  frmSpec.CurrentX = 10 + twp_X0Tr - (Len(udtLayoutSpec.Ytitle) + 0.5) * twp_XTick / 2
  frmSpec.CurrentY = 10
  frmSpec.Print udtLayoutSpec.Ytitle
ElseIf val_XMin < 0 And val_XMax <= 0 Then '3rd quadrant is shown
  frmSpec.Line (twp_X0Tr, twp_Y0 + twp_YTick)-(twp_X0Tr, twp_Y0 - twp_YRange), vbBlack 'Y-axis
  'prepare position title Y-axis and plot
  frmSpec.CurrentX = 10 + twp_X0Tr - (Len(udtLayoutSpec.Ytitle) + 0.5) * twp_XTick
  frmSpec.CurrentY = 10
  frmSpec.Print udtLayoutSpec.Ytitle
Else '1st quadrant is shown
  frmSpec.Line (twp_X0, twp_Y0 + twp_YTick)-(twp_X0, twp_Y0 - twp_YRange), vbBlack 'Y-axis
  'prepare position title Y-axis and plot
  frmSpec.CurrentX = twp_XLeftMargin
  frmSpec.CurrentY = 10
  frmSpec.Print udtLayoutSpec.Ytitle
End If
  
  
'**********  plot X-axis and title  **********
If val_YMin < 0 And val_YMax > 0 Then 'all four quadrants are shown
  frmSpec.Line (twp_X0 - twp_XTick, twp_Y0Tr)-(twp_X0 + twp_XRange, twp_Y0Tr), vbBlack 'X-axis
  frmSpec.CurrentX = 10
  frmSpec.CurrentY = twp_Y0Tr - 3 * twp_YTick
  frmSpec.Print udtLayoutSpec.XTitle
ElseIf val_YMin < 0 And val_YMax <= 0 Then '3rd quadrant is shown
  frmSpec.Line (twp_X0 - twp_XTick, twp_Y0Tr)-(twp_X0 + twp_XRange, twp_Y0Tr), vbBlack 'X-axis
  frmSpec.CurrentX = 10
  frmSpec.CurrentY = twp_Y0Tr - 3 * twp_YTick
  frmSpec.Print udtLayoutSpec.XTitle
Else '1st quadrant is shown
  frmSpec.Line (twp_X0 - twp_XTick, twp_Y0)-(twp_X0 + twp_XRange, twp_Y0), vbBlack 'X-axis
  frmSpec.CurrentX = 10
  frmSpec.CurrentY = frmSpec.ScaleHeight - twp_YBottomMargin
  frmSpec.Print udtLayoutSpec.XTitle
End If
  

'**********  plot ticks X-axis  **********
If val_YMin < 0 Then
  If val_XMin = Pl_SpanX * Int(val_XMin / Pl_SpanX) Then 'val_Xmin = Xtick
    frmSpec.Line (twp_X0, twp_Y0Tr + twp_XTick)-(twp_X0, twp_Y0Tr), vbBlack
  End If
  For nI = 1 To NumXTicks
    frmSpec.Line (twp_X0 - twp_OffSetX + nI * twp_XTickRange, twp_Y0Tr + twp_XTick)- _
    (twp_X0 - twp_OffSetX + nI * twp_XTickRange, twp_Y0Tr), vbBlack
  Next nI
Else
  If val_XMin = Pl_SpanX * Int(val_XMin / Pl_SpanX) Then 'val_Xmin = Xtick
    frmSpec.Line (twp_X0, twp_Y0 + twp_XTick)-(twp_X0, twp_Y0), vbBlack
  End If
  For nI = 1 To NumXTicks
    frmSpec.Line (twp_X0 - twp_OffSetX + nI * twp_XTickRange, twp_Y0 + twp_XTick)- _
    (twp_X0 - twp_OffSetX + nI * twp_XTickRange, twp_Y0), vbBlack
  Next nI
End If


'**********  plot labels to ticks from X-axis  **********
Dim nLenXLable As Integer 'length of lable X-axis ticks

frmSpec.ForeColor = clr_Plot(0)
If val_YMin < 0 Then 'val_Ymin < 0 implies that twp_Y0Tr has to used
  If val_XMin = Pl_SpanX * Int(val_XMin / Pl_SpanX) Then 'val_Xmin = Xlable
    If Abs(val_XMax - val_XMin) > 100000 Or Abs(val_XMax - val_XMin) < 0.00001 Then 'scientific notation
      If NumXTicks > 10 Then 'plot lables under an angle
        lHandleFont = CreateFontIndirect(udtFont)
        lOldFont = SelectObject(frmSpec.hdc, lHandleFont) 'save old font
        For nI = 0 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick + 40
          frmSpec.CurrentY = twp_Y0Tr + twp_XTick * 4
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
        lRetVal = SelectObject(frmSpec.hdc, lOldFont) 'reload old font
        lRetVal = DeleteObject(lHandleFont)
      Else 'plot horizontal (as standard)
        For nI = 0 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick / 2
          frmSpec.CurrentY = twp_Y0Tr + twp_XTick * 2
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
      End If
    Else 'non-scientific notation
      For nI = 0 To NumXTicks
        nLenXLable = Len(Trim(Str$(Pl_SpanX * Int((val_XMin + Pl_SpanX * nI) / Pl_SpanX)))) 'to avoid to get long lables (this will misplace the label under the tick)
        frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
        - nLenXLable * twp_XTick / 2
        frmSpec.CurrentY = twp_Y0Tr + twp_XTick * 2
        frmSpec.Print val_XMin - OffSetX + Pl_SpanX * nI
      Next nI
    End If
  Else 'val_Xmin doesn't get a lable
    If Abs(val_XMax - val_XMin) > 100000 Or Abs(val_XMax - val_XMin) < 0.00001 Then 'scientific notation
      If NumXTicks > 10 Then 'plot lables under an angle
        lHandleFont = CreateFontIndirect(udtFont)
        lOldFont = SelectObject(frmSpec.hdc, lHandleFont)
        For nI = 1 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick + 40
          frmSpec.CurrentY = twp_Y0Tr + twp_XTick * 4
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
        lRetVal = SelectObject(frmSpec.hdc, lOldFont)
        lRetVal = DeleteObject(lHandleFont)
      Else 'plot horizontal (as standard)
        For nI = 1 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick / 2
          frmSpec.CurrentY = twp_Y0Tr + twp_XTick * 2
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
      End If
    Else 'non-scientific notation
      For nI = 1 To NumXTicks
        nLenXLable = Len(Trim(Str$(Pl_SpanX * Int((val_XMin + Pl_SpanX * nI) / Pl_SpanX))))
        frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
        - nLenXLable * twp_XTick / 2
        frmSpec.CurrentY = twp_Y0Tr + twp_XTick * 2
        frmSpec.Print val_XMin - OffSetX + Pl_SpanX * nI
      Next nI
    End If
  End If
Else 'val_Ymin >= 0 ; this implies that twp_Y0 is used
  If val_XMin = Pl_SpanX * Int(val_XMin / Pl_SpanX) Then 'val_Xmin = Xlable
    If Abs(val_XMax - val_XMin) > 100000 Or Abs(val_XMax - val_XMin) < 0.00001 Then 'scientific notation
      If NumXTicks > 10 Then 'plot lables under an angle
        lHandleFont = CreateFontIndirect(udtFont)
        lOldFont = SelectObject(frmSpec.hdc, lHandleFont)
        For nI = 0 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick + 40
          frmSpec.CurrentY = twp_Y0 + twp_XTick * 4
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
        lRetVal = SelectObject(frmSpec.hdc, lOldFont)
        lRetVal = DeleteObject(lHandleFont)
      Else 'plot horizontal (as standard)
        For nI = 0 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick / 2
          frmSpec.CurrentY = twp_Y0 + twp_XTick * 2
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
      End If
    Else 'non-scientific notation
      For nI = 0 To NumXTicks
        nLenXLable = Len(Trim(Str$(Pl_SpanX * Int((val_XMin + Pl_SpanX * nI) / Pl_SpanX))))
        frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
        - nLenXLable * twp_XTick / 2
        frmSpec.CurrentY = twp_Y0 + twp_XTick * 2
        frmSpec.Print Str(val_XMin - OffSetX + Pl_SpanX * nI)
      Next nI
    End If
  Else
    If Abs(val_XMax - val_XMin) > 100000 Or Abs(val_XMax - val_XMin) < 0.00001 Then 'scientific notation
      If NumXTicks > 10 Then 'plot lables under an angle
        lHandleFont = CreateFontIndirect(udtFont)
        lOldFont = SelectObject(frmSpec.hdc, lHandleFont)
        For nI = 1 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick + 40
          frmSpec.CurrentY = twp_Y0 + twp_XTick * 4
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
        lRetVal = SelectObject(frmSpec.hdc, lOldFont)
        lRetVal = DeleteObject(lHandleFont)
      Else 'plot horizontal (as standard)
        For nI = 1 To NumXTicks
          frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
          - (Len(Format(val_XMin + Pl_SpanX * nI, "Scientific"))) * twp_XTick / 2
          frmSpec.CurrentY = twp_Y0 + twp_XTick * 2
          frmSpec.Print Format(val_XMin - OffSetX + Pl_SpanX * nI, "Scientific")
        Next nI
      End If
    Else 'non-scientific notation
      For nI = 1 To NumXTicks
        nLenXLable = Len(Trim(Str$(Pl_SpanX * Int((val_XMin + Pl_SpanX * nI) / Pl_SpanX))))
        frmSpec.CurrentX = twp_X0 - twp_OffSetX + nI * twp_XTickRange _
        - nLenXLable * twp_XTick / 2
        frmSpec.CurrentY = twp_Y0 + twp_XTick * 2
        frmSpec.Print Str(val_XMin - OffSetX + Pl_SpanX * nI)
      Next nI
    End If
  End If
End If
frmSpec.ForeColor = vbBlack


End Sub

Public Sub SetZoomValues(twp_StartXMD As Single, twp_StartYMD As Single, _
 twp_WidthMD As Single, twp_HeightMD As Single)
 
Dim val_ZoomXStart As Double
Dim val_ZoomXEnd As Double
Dim val_ZoomYStart As Double
Dim val_ZoomYEnd As Double
 
'check if boundaries of drawing pane are crossed
If twp_StartXMD < twp_XLeftMargin Then twp_StartXMD = twp_XLeftMargin
If twp_StartXMD > (twp_XLeftMargin + twp_XRange) Then twp_StartXMD = (twp_XLeftMargin + twp_XRange)
If twp_StartYMD < twp_YTopMargin Then twp_StartYMD = twp_YTopMargin
If twp_StartYMD > (twp_YTopMargin + twp_YRange) Then twp_StartYMD = (twp_YTopMargin + twp_YRange)
If twp_WidthMD > twp_XRange Then twp_WidthMD = twp_XRange
If twp_HeightMD > twp_YRange Then twp_HeightMD = twp_YRange

'calculate X-value
val_ZoomXStart = val_XMin + val_XRange * (twp_StartXMD - twp_XLeftMargin) / twp_XRange
val_ZoomXEnd = val_XMin + val_XRange * (twp_StartXMD + twp_WidthMD - twp_XLeftMargin) / twp_XRange

'calculate Y-value
val_ZoomYEnd = val_YMin + val_YRange * (twp_YRange - (twp_StartYMD - twp_YTopMargin)) / twp_YRange
val_ZoomYStart = val_YMin + val_YRange * (twp_YRange - (twp_StartYMD + twp_HeightMD - twp_YTopMargin)) / twp_YRange

With udtMyGraphLayout
  .X0 = val_ZoomXStart
  .X1 = val_ZoomXEnd
  .Y0 = val_ZoomYStart
  .Y1 = val_ZoomYEnd
End With
  

End Sub

Public Function GetValues(twp_XPosition As Single, twp_Yposition As Single) As COORDINATE

Dim val_CoordX As Double 'value X-coordinate
Dim val_CoordY As Double 'value Y-coordinate
Dim flg_OutsidePane As Boolean 'true when clicked outside drawing pane

'check if boundaries of drawing pane are crossed
If twp_XPosition < twp_XLeftMargin Then flg_OutsidePane = True
If twp_XPosition > (twp_XLeftMargin + twp_XRange) Then flg_OutsidePane = True
If twp_Yposition < twp_YTopMargin Then flg_OutsidePane = True
If twp_Yposition > (twp_YTopMargin + twp_YRange) Then flg_OutsidePane = True

'get X-coordinate
If flg_OutsidePane = False Then
  val_CoordX = val_XMin + val_XRange * (twp_XPosition - twp_XLeftMargin) / twp_XRange
Else
  val_CoordX = 0
End If

'get Y-coordinate
If flg_OutsidePane = False Then
  val_CoordY = val_YMin + val_YRange * (twp_YRange - (twp_Yposition - twp_YTopMargin)) / twp_YRange
Else
  val_CoordY = 0
End If

With GetValues
  .X = val_CoordX
  .Y = val_CoordY
End With


End Function


