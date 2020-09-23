Attribute VB_Name = "modLineArt"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const VK_ESCAPE = &H1B

' Pre-Defined Stuff
Public Const MaxLines As Integer = 12
Public Const MaxPoints As Integer = 4
Public Const SnapUnit As Single = 80

' Track Mouse Clicks and Point Definitions
Public cPoint As Integer               ' Click Count
Public xPoint(MaxPoints) As Integer    ' X-Coord of Click
Public yPoint(MaxPoints) As Integer    ' Y-Coord of Click

' Coordinates for points along line
Dim xa(MaxLines) As Integer
Dim ya(MaxLines) As Integer
Dim xb(MaxLines) As Integer
Dim yb(MaxLines) As Integer

' Line Art Options
Public Enum LineArtOption
 PointA_LineBC = 0
 LineAB_LineBC = 1
 LineAB_LineCD = 2
End Enum

Public Sub LineArt(pct As PictureBox, a As Integer, b As Integer, c As Integer, d As Integer)
 ' point A - Starting Point of Line AB
 ' point B - Ending Point of Line AB
 ' point C - Starting Point of Line CD
 ' point D - Ending Point of Line CD
 Dim t As Integer
 Dim xDiff As Single
 Dim yDiff As Single
 
 ' Calculate "Slope" of line AB
 xDiff = (xPoint(a) - xPoint(b)) / MaxLines
 yDiff = (yPoint(a) - yPoint(b)) / MaxLines
 
 ' Define Points along line AB
 For t = 0 To MaxLines
  xa(t) = xPoint(a) - t * xDiff
  ya(t) = yPoint(a) - t * yDiff
 Next t
 
 ' Calculate "Slope" of line CD
 xDiff = (xPoint(c) - xPoint(d)) / MaxLines
 yDiff = (yPoint(c) - yPoint(d)) / MaxLines
 
 ' Define Points along line CD
 For t = 0 To MaxLines
  xb(t) = xPoint(c) - t * xDiff
  yb(t) = yPoint(c) - t * yDiff
 Next t
 
 pct.Line (xPoint(a), yPoint(a))-(xPoint(b), yPoint(b))
 For t = 0 To MaxLines
  pct.Line (xa(t), ya(t))-(xb(t), yb(t))
 Next t
 pct.Line (xPoint(c), yPoint(c))-(xPoint(d), yPoint(d))

End Sub

Public Function Snap(unit As Single, coord As Single) As Single
 Dim halfUnit As Single
 
 halfUnit = unit / 2
 
 Snap = coord - (coord + halfUnit) Mod unit
End Function
 
Public Function LineColor(cp As Integer) As Long
 
 If cp = 3 Then
  LineColor = RGB(0, 200, 200)
 Else
  LineColor = RGB(0, 0, 255)
 End If
End Function

