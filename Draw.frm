VERSION 5.00
Begin VB.Form frmDraw 
   AutoRedraw      =   -1  'True
   Caption         =   "Graphic"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   10215
   Begin VB.Shape shpMouseDraw 
      Height          =   495
      Left            =   1380
      Top             =   6960
      Visible         =   0   'False
      Width           =   675
   End
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fStartXMD As Single 'startposition X of rectangle
Public fStartYMD As Single 'startposition Y of rectangle
Public fHeightMD As Single 'difference Y direction of rectangle
Public fWidthMD As Single 'difference X direction of rectangle

Private Sub Form_DblClick()

Call frmSetCond.DefineLayout
Plot frmDraw, dPlot, udtMyGraphLayout

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim udtMyCoord As COORDINATE
Dim sCoord As String 'containing coordinate-info
Dim nLenStr As Integer 'length of string coordinate-info

If Button = 1 And Shift = vbShiftMask Then 'SHIFT + left mouse
  shpMouseDraw.Left = X
  fStartXMD = X
  shpMouseDraw.Top = Y
  fStartYMD = Y
  shpMouseDraw.Width = 100
  shpMouseDraw.Height = 100
  shpMouseDraw.BorderStyle = 3
  shpMouseDraw.Visible = True
End If

If Button = 2 Then 'Right mouse click
  udtMyCoord = GetValues(X, Y)
  Load frmShowValues
  frmShowValues.Top = frmDraw.Top + Y
  frmShowValues.Left = frmDraw.Left + X
  frmShowValues.Show vbModeless
  frmShowValues.CurrentX = 10
  frmShowValues.CurrentY = 10
  sCoord = "(X,Y)=(" & Str$(udtMyCoord.X) & " , " & Str$(udtMyCoord.Y) & ")"
  nLenStr = Len(sCoord)
  frmShowValues.Width = nLenStr * 80
  If (frmShowValues.Left + frmShowValues.Width) > (frmDraw.Width + frmDraw.Left) Then
    frmShowValues.Left = frmDraw.Width + frmDraw.Left - frmShowValues.Width
  End If
  frmShowValues.Print sCoord
End If
  
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 And Shift = vbShiftMask Then 'SHIFT + left mouse
  If (X - fStartXMD) > 0 Then
    shpMouseDraw.Width = Abs(X - fStartXMD)
  Else
    shpMouseDraw.Left = X
    shpMouseDraw.Width = Abs(fStartXMD - X)
  End If
  If (Y - fStartYMD) > 0 Then
    shpMouseDraw.Height = Abs(Y - fStartYMD)
  Else
    shpMouseDraw.Top = Y
    shpMouseDraw.Height = Abs(fStartYMD - Y)
  End If
            
End If

End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 And Shift = vbShiftMask Then 'SHIFT + left mouse
  shpMouseDraw.Visible = False
  If (X - fStartXMD) < 0 Then fStartXMD = fStartXMD - Abs(X - fStartXMD)
  If (Y - fStartYMD) < 0 Then fStartYMD = fStartYMD - Abs(Y - fStartYMD)
  fWidthMD = Abs(X - fStartXMD)
  fHeightMD = Abs(Y - fStartYMD)
  SetZoomValues fStartXMD, fStartYMD, fWidthMD, fHeightMD
  Plot frmDraw, dPlot, udtMyGraphLayout 'plot the zoomed area
  
End If
  

End Sub
