VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTranslution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Translucent Forms...
Private Declare Function ReleaseDC Lib "USER32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "USER32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

'For Dragging Borderless Forms...
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

'Prevents function recursion...
Public iRecursion As Boolean
Public tColor As Long
Public Sub DragForm(Who As Form)

On Local Error Resume Next

'Move the borderless form...
Call ReleaseCapture
Call SendMessage(Who.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub
Public Sub MakeTranslucent(Who As Variant, Optional tColor As Long)  'Was (Who as Object) before...

On Local Error Resume Next

Dim HW As Long
Dim HA As Long
Dim iLeft As Integer
Dim iTop As Integer
Dim iWidth As Integer
Dim iHeight As Integer

If IsMissing(tColor) Or tColor = 0 Then
    tColor = RGB(0, 0, 200)
End If

Who.AutoRedraw = True
Who.Hide

DoEvents

HW = GetDesktopWindow()
HA = GetDC(HW)

'Get the Left, Top, Width and Height of the Form...
iLeft = Who.Left / Screen.TwipsPerPixelX
iTop = Who.Top / Screen.TwipsPerPixelY '+ 25    If using a form with a titlebar (border)...
iWidth = Who.ScaleWidth
iHeight = Who.ScaleHeight

'Now, Transfer the contents of the Desktop Window to the Form...
Call BitBlt(Who.hdc, 0, 0, iWidth, iHeight, HA, iLeft, iTop, SRCCOPY) 'iLeft + 4    If using a form with a titlebar (border)...

'Show...
Who.Picture = Who.Image
'Who.Show

'Release the DC...
Call ReleaseDC(HW, HA)

'Add color...
Who.DrawMode = 9
Who.ForeColor = tColor
Who.Line (0, 0)-(iWidth, iHeight), , BF

End Sub

