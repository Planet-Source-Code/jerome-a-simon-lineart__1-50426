VERSION 5.00
Begin VB.Form frmLineArt 
   Caption         =   "Form1"
   ClientHeight    =   3300
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2505
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   2505
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctLineArt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1155
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   1095
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   0
      Width           =   1035
   End
   Begin VB.PictureBox pctBackBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1155
      Left            =   540
      ScaleHeight     =   1095
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Menu mnuLine 
      Caption         =   "&Option"
      Begin VB.Menu mnuLineOption 
         Caption         =   "Point A - Line &BC"
         Index           =   0
      End
      Begin VB.Menu mnuLineOption 
         Caption         =   "Line AB - Line B&C"
         Index           =   1
      End
      Begin VB.Menu mnuLineOption 
         Caption         =   "Line AB - Line C&D"
         Index           =   2
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "&Image"
      Begin VB.Menu mnuImageClear 
         Caption         =   "&Clear"
      End
   End
End
Attribute VB_Name = "frmLineArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 cPoint = 0
 LineType PointA_LineBC
 
End Sub

Private Sub Form_Resize()
 pctLineArt.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
 pctBackBuffer.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
 
End Sub

Private Sub mnuImageClear_Click()
 pctBackBuffer.Cls

End Sub

Private Sub mnuLineOption_Click(Index As Integer)
 LineType Index
 cPoint = 0

End Sub

Private Sub pctLineArt_KeyPress(KeyAscii As Integer)
 If KeyAscii = VK_ESCAPE Then
  If cPoint > 0 Then
   cPoint = cPoint - 1
  End If
 End If

End Sub

Private Sub pctLineArt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim t As Integer
 
 x = Snap(SnapUnit, x)
 y = Snap(SnapUnit, y)
 
 ' Record Mouse Coords
 cPoint = cPoint + 1
 xPoint(cPoint) = x
 yPoint(cPoint) = y
 
 If mnuLineOption(LineAB_LineBC).Checked And cPoint = 2 Then
  cPoint = 3
  xPoint(cPoint) = xPoint(2)
  yPoint(cPoint) = yPoint(2)
 End If
 If mnuLineOption(PointA_LineBC).Checked And cPoint = 1 Then
  cPoint = 2
  xPoint(cPoint) = xPoint(1)
  yPoint(cPoint) = yPoint(1)
 End If
 
 If Not cPoint < MaxPoints Then
  LineArt pctBackBuffer, cPoint - 3, cPoint - 2, cPoint - 1, cPoint
  cPoint = 0
 End If

End Sub

Private Sub pctLineArt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim tBar As String
 Dim t As Integer
 
 ' Clear pctLineArt
 'pctLineArt.Picture = pctBackBuffer.Picture
 With pctLineArt
  BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, pctBackBuffer.hDC, 0, 0, SRCCOPY
 End With
 
 ' Snap Cursor to Points
 x = Snap(SnapUnit, x)
 y = Snap(SnapUnit, y)
 pctLineArt.Line (x, 0)-(x, pctLineArt.ScaleHeight), RGB(0, 150, 0)
 pctLineArt.Line (0, y)-(pctLineArt.ScaleWidth, y), RGB(0, 150, 0)
 pctLineArt.Circle (x, y), SnapUnit, RGB(150, 0, 0)
 
 If cPoint = 0 Then
  ' No Points Defined - Nothing to do!
  Exit Sub
 End If
 
 ' Atleast One Point Defined Set first "Line" point
 t = 1
 pctLineArt.Line (xPoint(t), yPoint(t))-(xPoint(t), yPoint(t))
    
 ' Draw "Other" lines (if any)
 Do While t < cPoint
  t = t + 1
  pctLineArt.Line (xPoint(t - 1), yPoint(t - 1))-(xPoint(t), yPoint(t)), LineColor(t)
 Loop
 
 ' Fallow Mouse Pointer - from last point
 pctLineArt.Line -(x, y), LineColor(t + 1)
 If cPoint < MaxPoints Then
  If cPoint = 3 Then
   xPoint(4) = x
   yPoint(4) = y
   LineArt pctLineArt, cPoint - 2, cPoint - 1, cPoint, cPoint + 1
  End If
 Else
  LineArt pctBackBuffer, cPoint - 3, cPoint - 2, cPoint - 1, cPoint
 End If
 
End Sub

Private Sub LineType(lType As Integer)
 Dim t As Integer
 
 For t = 0 To LineAB_LineCD
  mnuLineOption(t).Checked = False
 Next t
 mnuLineOption(lType).Checked = True
 
End Sub

