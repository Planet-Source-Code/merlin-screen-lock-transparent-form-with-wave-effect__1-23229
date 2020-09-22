VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPicWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Klasse:    ClsPicWave
'Author:    Merlin
'Funktion:  Picture Animation mit Wellen Algorythmus
'Eingabe:   Source hDC,Destination hDC,Wellenart,Amplitude,Anzahl der Wellen,
           'Geschwindigkeit
'Import     Function BitBlt von gdi32

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020


Private Const Pi As Double = 3.1456 / 180
Private WaveScene As Long
Private PicWidth As Long
Private PicHeight As Long
Private Dist() As Variant

'Horizontal =1
'Verticle   =2
'Both       =3
'Cosine Wav =4
'2Tape Wave =5
'Ripple     =6

Private m_Speed As Long
Private m_Cancel As Boolean
Private m_WaveArt As Integer
Private m_Amplitude As Long
Private m_Waves As Long
Public SourcePic As Variant
Public DestinationPic As Variant


Public Property Let Speed(Value As Long)
    If Value < 1 Then Value = 1
    If Value > 20 Then Value = 20
    m_Speed = Value
End Property

Property Get Speed() As Long
    Speed = m_Speed
End Property

Public Property Let Cancel(Value As Boolean)
    m_Cancel = Value
End Property

Public Property Get Cancel() As Boolean
    Cancel = m_Cancel
End Property

Public Property Let Waves(Value As Long)
    If Value < 1 Then Value = 1
    If Value > 15 Then Value = 15
    m_Waves = Value
End Property

Public Property Get Waves() As Long
    Waves = m_Waves
End Property

Public Property Let Amplitude(Value As Long)
    If Value > 100 Then Value = 100
    If Value < -100 Then Value = -100
    m_Amplitude = Value
End Property

Property Get Amplitude() As Long
    Amplitude = m_Amplitude
End Property

Public Property Let WaveArt(Value As Integer)
    If Value > 6 Then Value = 6
    If Value < 1 Then Value = 1
    m_WaveArt = Value
End Property

Public Property Get WaveArt() As Integer
    WaveArt = m_WaveArt
End Property




Public Function Wave()
PicInit

    Do
        
        DoEvents
        If WaveScene >= 1080 - 20 Or WaveScene < 0 Then WaveScene = 0
        On Error Resume Next
        WaveScene = WaveScene + 20
         
        Select Case WaveArt

            'Horizontal
            Case Is = 1
                For x = 1 To PicWidth
                    DoEvents
                    Angel = WaveScene + (x * Waves)
                    Wave = Sin(Pi * Angel) * Amplitude
                    BitBlt DestinationPic.hdc, x, Wave, 2, PicHeight, SourcePic.hdc, x, 0, SRCCOPY
                Next x
                DestinationPic.Refresh
                
            'Verticle
            Case Is = 2
                For x = 1 To PicHeight
                    DoEvents
                    Angel = WaveScene + (x * Waves)
                    Wave = Sin(Pi * Angel) * Amplitude
                    BitBlt DestinationPic.hdc, Wave, x, PicWidth, 2, SourcePic.hdc, 0, x, SRCCOPY
                Next
                DestinationPic.Refresh
                
            'Both
            Case Is = 3
            
                For x = 1 To PicHeight
                    DoEvents
                    Angel = WaveScene + (x * Waves)
                    Wave = Sin(Pi * Angel) * Amplitude
                    BitBlt DestinationPic.hdc, Wave, x + Wave, PicWidth, 4, SourcePic.hdc, 0, x, SRCCOPY
                Next
                DestinationPic.Refresh

            'Cosine Wav
            Case Is = 4
            
                For x = 1 To PicHeight
                    DoEvents
                    Angel = WaveScene + (x * Waves)
                    Wave = 1 / Sin(Pi * Angel) * Amplitude
                    BitBlt DestinationPic.hdc, Wave, x, PicWidth, 4, SourcePic.hdc, 0, x, SRCCOPY
                Next
                DestinationPic.Refresh
        
            '2Tape Wave
            Case Is = 5
                
                For x = 1 To PicWidth Step Speed
                    DoEvents
                    Angel = WaveScene + (x * Waves)
                    Wave = Sin(Pi * Angel) * Amplitude
                    For y = 1 To PicHeight Step Speed
                        Angel2 = WaveScene + (y * Waves)
                        Wave2 = Cos(Pi * Angel2) * Amplitude
                        BitBlt DestinationPic.hdc, x + Wave2, y + Wave, Speed, Speed, SourcePic.hdc, x, y, SRCCOPY
                    Next y
                Next x
                DestinationPic.Refresh
        
            'Ripple
            Case Is = 6
                For x = 1 To PicWidth Step Speed
                    DoEvents
                    For y = 1 To PicHeight Step Speed
                        Angel = WaveScene + (Dist(x, y) * Waves)
                        Wave = Cos(Pi * Angel) * Amplitude
                        BitBlt DestinationPic.hdc, x + Wave, y + Wave, Speed + 1, Speed + 1, SourcePic.hdc, x, y, SRCCOPY
                    Next y
                Next x
                DestinationPic.Refresh
        
         
         End Select
    
    Loop While Cancel = False

End Function

Private Sub PicInit()

Cancel = False
DestinationPic.ScaleMode = 3
PicWidth = DestinationPic.ScaleWidth
PicHeight = DestinationPic.ScaleHeight
DestinationPic.AutoRedraw = True

If WaveArt = 6 Then
    DoTable
End If

End Sub


Private Sub DoTable()
    ReDim Dist(1 To PicWidth, 1 To PicHeight) As Variant
    For x = 1 To PicWidth
        For y = 1 To PicHeight
            Dist(x, y) = Distance(PicWidth / 2, PicHeight / 2, x, y)
        Next y
    Next x
End Sub

Private Function Distance(sx, sy, Ex, Ey) As Long
    Distance = Sqr((Ex - sx) ^ 2 + (Ey - sy) ^ 2)
End Function

Private Sub Class_Initialize()
    Cancel = True
End Sub
