VERSION 5.00
Begin VB.Form frmSetCond 
   Caption         =   "Conditions"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   2985
   Begin VB.CommandButton cmdOK 
      Caption         =   "Plot"
      Height          =   435
      Left            =   900
      TabIndex        =   19
      Top             =   8940
      Width           =   975
   End
   Begin VB.Frame fraQuadrant 
      Caption         =   "Draw data for quadrant: "
      Height          =   1815
      Left            =   60
      TabIndex        =   31
      Top             =   7020
      Width           =   2775
      Begin VB.OptionButton optQuad 
         Caption         =   "all quadrants"
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   18
         Top             =   1440
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optQuad 
         Caption         =   "4th quadrant"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   1140
         Width           =   1335
      End
      Begin VB.OptionButton optQuad 
         Caption         =   "3rd quadrant"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   16
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optQuad 
         Caption         =   "2nd quadrant"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Top             =   540
         Width           =   1335
      End
      Begin VB.OptionButton optQuad 
         Caption         =   "1st quadrant"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraDrawnAs 
      Caption         =   "Drawn demo signal(s) As"
      Height          =   1995
      Left            =   60
      TabIndex        =   30
      Top             =   4860
      Width           =   2775
      Begin VB.CheckBox chkSignal 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   1305
         TabIndex        =   37
         Top             =   540
         Width           =   255
      End
      Begin VB.CheckBox chkSignal 
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   2220
         TabIndex        =   33
         Top             =   540
         Width           =   255
      End
      Begin VB.CheckBox chkSignal 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   32
         Top             =   540
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.OptionButton optDrawAs 
         Caption         =   "bars"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   1440
         Width           =   675
      End
      Begin VB.OptionButton optDrawAs 
         Caption         =   "line"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   1140
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optDrawAs 
         Caption         =   "points"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.Frame fraDemo1 
         Caption         =   "Sig 1"
         Height          =   1515
         Left            =   120
         TabIndex        =   34
         Top             =   300
         Width           =   855
      End
      Begin VB.Frame fraDemo2 
         Caption         =   "Sig 2"
         Height          =   1515
         Left            =   1080
         TabIndex        =   35
         Top             =   300
         Width           =   735
         Begin VB.OptionButton optDrawAs 
            Caption         =   "bars"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   40
            Top             =   1140
            Width           =   255
         End
         Begin VB.OptionButton optDrawAs 
            Caption         =   "bars"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   39
            Top             =   840
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optDrawAs 
            Caption         =   "bars"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   38
            Top             =   540
            Width           =   255
         End
      End
      Begin VB.Frame fraDem03 
         Caption         =   "Sig 3"
         Height          =   1515
         Left            =   1980
         TabIndex        =   36
         Top             =   300
         Width           =   675
         Begin VB.OptionButton optDrawAs 
            Caption         =   "bars"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   43
            Top             =   1140
            Width           =   255
         End
         Begin VB.OptionButton optDrawAs 
            Caption         =   "bars"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   42
            Top             =   840
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton optDrawAs 
            Caption         =   "bars"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   41
            Top             =   540
            Width           =   255
         End
      End
   End
   Begin VB.Frame fraLayout 
      Caption         =   "Layout"
      Height          =   4395
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      Begin VB.TextBox txtValueX1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   3150
         Width           =   555
      End
      Begin VB.TextBox txtValueX0 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   2748
         Width           =   555
      End
      Begin VB.TextBox txtYTitle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox txtXTitle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   3552
         Width           =   855
      End
      Begin VB.TextBox txtValueY1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   2346
         Width           =   555
      End
      Begin VB.TextBox txtValueY0 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1944
         Width           =   555
      End
      Begin VB.TextBox txtX1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "1000"
         Top             =   1542
         Width           =   555
      End
      Begin VB.TextBox txtX0 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "0"
         Top             =   1140
         Width           =   555
      End
      Begin VB.CheckBox chkGridLine 
         Caption         =   "Check1"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkOrigin 
         Caption         =   "Check1"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lblEndX1 
         AutoSize        =   -1  'True
         Caption         =   "maximum value X-range"
         Height          =   195
         Left            =   840
         TabIndex        =   27
         Top             =   3210
         Width           =   1680
      End
      Begin VB.Label lblStartX0 
         AutoSize        =   -1  'True
         Caption         =   "minimum value X-range"
         Height          =   195
         Left            =   840
         TabIndex        =   26
         Top             =   2808
         Width           =   1635
      End
      Begin VB.Label lblYTitle 
         AutoSize        =   -1  'True
         Caption         =   "Title Y-axis"
         Height          =   195
         Left            =   1140
         TabIndex        =   29
         Top             =   4020
         Width           =   765
      End
      Begin VB.Label lblXTitle 
         AutoSize        =   -1  'True
         Caption         =   "Title X-axis"
         Height          =   195
         Left            =   1140
         TabIndex        =   28
         Top             =   3612
         Width           =   765
      End
      Begin VB.Label lblEndY 
         AutoSize        =   -1  'True
         Caption         =   "maximum value Y-range"
         Height          =   195
         Left            =   840
         TabIndex        =   25
         Top             =   2406
         Width           =   1680
      End
      Begin VB.Label lblStartY 
         AutoSize        =   -1  'True
         Caption         =   "minimum value Y-range"
         Height          =   195
         Left            =   840
         TabIndex        =   24
         Top             =   2004
         Width           =   1635
      End
      Begin VB.Label lblIndexEnd 
         AutoSize        =   -1  'True
         Caption         =   "index End X-range"
         Height          =   195
         Left            =   840
         TabIndex        =   23
         Top             =   1602
         Width           =   1305
      End
      Begin VB.Label lblIndexStart 
         AutoSize        =   -1  'True
         Caption         =   "index Start X-range"
         Height          =   195
         Left            =   840
         TabIndex        =   22
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label lblGridLine 
         AutoSize        =   -1  'True
         Caption         =   "Gridlines?"
         Height          =   195
         Left            =   420
         TabIndex        =   21
         Top             =   780
         Width           =   690
      End
      Begin VB.Label lblOrigin 
         AutoSize        =   -1  'True
         Caption         =   "Show origin?"
         Height          =   195
         Left            =   420
         TabIndex        =   20
         Top             =   420
         Width           =   915
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmSetCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public n As Integer 'counter
Public m As Integer 'counter
Public dYmax As Double
Public dYmin As Double

Private Sub cmdOK_Click()

'create functions value to draw
CreateValues

'Define window Layout
DefineLayout
  
'plot the data
Plot frmDraw, dPlot, udtMyGraphLayout
  
    
End Sub


Private Sub Form_Load()

Load frmDraw
frmDraw.Show vbModeless
frmDraw.Width = Screen.Width - frmSetCond.Width
frmDraw.ScaleHeight = frmSetCond.ScaleHeight
frmDraw.Left = frmSetCond.Width
txtXTitle.Text = "X-values"
txtYTitle.Text = "Y-values"
chkSignal(0).Value = 1
cmdOK_Click

End Sub
Public Sub DefineLayout()

Dim nSum As Integer 'sum of number of traces to plot

With udtMyGraphLayout
  .XTitle = txtXTitle.Text
  .Ytitle = txtYTitle.Text
  If chkOrigin.Value = 0 Then
    .blnOrigin = False
    Else
    .blnOrigin = True
  End If
  If chkGridLine.Value = 0 Then
    .blnGridLine = False
    Else
    .blnGridLine = True
  End If
  'X-range
  If Abs((Val(txtValueX1.Text) - Val(txtValueX0.Text))) >= 0 And txtValueX1.Text <> "" _
  And txtValueX0.Text <> "" Then
    .X0 = Val(txtValueX0.Text)
    .X1 = Val(txtValueX1.Text)
    Else
    .X0 = dPlot(LBound(dPlot, 1), 0)
    .X1 = dPlot(UBound(dPlot, 1), 0)
  End If
  'Y-range
  If Abs((Val(txtValueY1.Text) - Val(txtValueY0.Text))) >= 0 And txtValueY1.Text <> "" _
  And txtValueY0.Text <> "" Then
    .Y0 = Val(txtValueY0.Text)
    .Y1 = Val(txtValueY1.Text)
    Else
    .Y0 = dYmin 'dPlot(LBound(dPlot, 1), 2)
    .Y1 = dYmax 'dPlot(UBound(dPlot, 1), 2)
  End If
  'index start-X
  If Val(txtX0.Text) >= LBound(dPlot, 1) And Val(txtX0.Text) <= UBound(dPlot, 1) _
  And (Val(txtX1.Text) - Val(txtX0.Text)) > 0 Then
    .lStart = Val(txtX0.Text)
    Else
    .lStart = LBound(dPlot, 1)
    txtX0.Text = Str(LBound(dPlot, 1))
    txtX1.Text = Str(UBound(dPlot, 1))
  End If
  'index end-X
  If Val(txtX1.Text) >= LBound(dPlot, 1) And Val(txtX1.Text) <= UBound(dPlot, 1) _
  And (Val(txtX1.Text) - Val(txtX0.Text)) > 0 Then
    .lEnd = Val(txtX1.Text)
    Else
    .lEnd = UBound(dPlot, 1)
    txtX0.Text = Str(LBound(dPlot, 1))
    txtX1.Text = Str(UBound(dPlot, 1))
  End If
  .asX = 0
  
  'check number of demo-signals to plot
  nSum = 0
  For n = 0 To 2
    If chkSignal(n).Value = 1 Then
      nSum = nSum + 1
    End If
  Next n
  ReDim .asY(nSum - 1) 'nSum traces to plot
  ReDim .DrawTrace(nSum - 1) 'nSum traces to plot
  
  'define traces to plot
  nSum = 0
  For n = 0 To 2
    If chkSignal(n).Value = 1 Then
      .asY(nSum) = n + 1 'plot trace checked (=dplot(n,1))
      If optDrawAs(n * 3 + 0).Value = True Then
        .DrawTrace(nSum) = AS_POINT
      End If
      If optDrawAs(n * 3 + 1).Value = True Then
        .DrawTrace(nSum) = AS_CONLINE
      End If
      If optDrawAs(n * 3 + 2).Value = True Then
        .DrawTrace(nSum) = AS_BAR
      End If
      nSum = nSum + 1
    End If
  Next n
End With

End Sub

Private Sub CreateValues()

Dim nOffSetIndex As Integer
Dim nOffSetValue As Integer
Dim nSign As Integer

If optQuad(0).Value = True Then '1st quadrant
  nOffSetValue = 10
  nOffSetIndex = 25
  nSign = 1
End If
If optQuad(1).Value = True Then '2nd quadrant
  nOffSetValue = -30
  nOffSetIndex = 25
  nSign = -1
End If
If optQuad(2).Value = True Then '3rd quadrant
  nOffSetValue = -30
  nOffSetIndex = -125
  nSign = -1
End If
If optQuad(3).Value = True Then '4th quadrant
  nOffSetValue = 10
  nOffSetIndex = -125
  nSign = 1
End If
If optQuad(4).Value = True Then 'all quadrants
  nOffSetValue = -5
  nOffSetIndex = -50
  nSign = 1
End If


ReDim dPlot(100, 3) As Double
Randomize

For n = 0 To 100
  dPlot(n, 0) = (n + nOffSetIndex)
  dPlot(n, 1) = 0.01 * (n - 50) ^ 2 + nOffSetValue 'Sqr(n) + nSign * 5 * Sin(0.1 * n) + Rnd + nOffSetValue
  dPlot(n, 2) = nSign * 8 * Cos(0.1 * n) + Rnd + nOffSetValue
  dPlot(n, 3) = Sin(n ^ 2) + Rnd + nOffSetValue
Next n

'determine Ymin and Ymax
dYmax = dPlot(0, 1)
dYmin = dPlot(0, 1)
For m = 1 To 3
  For n = 0 To UBound(dPlot, 1)
    If dYmax < dPlot(n, m) Then
      dYmax = dPlot(n, m)
    End If
    If dYmin > dPlot(n, m) Then
      dYmin = dPlot(n, m)
    End If
  Next n
Next m
  


End Sub

Private Sub mnuExit_Click()

Unload frmDraw
Unload frmShowValues
Unload Me
End Sub
