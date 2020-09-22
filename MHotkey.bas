Attribute VB_Name = "MHotkey"
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_SHIFT = &H4
Private Const PM_REMOVE = &H1
Private Const WM_HOTKEY = &H312

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private bCancel As Boolean
Private Sub ProcessMessages()
    Dim Message As Msg
    'loop until bCancel is set to True
    Do While Not bCancel
        'wait for a message
        WaitMessage
        'check if it's a HOTKEY-message
        If PeekMessage(Message, Parent.hWnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
            If cWave.Cancel = True Then WaterStart
        End If
        'let the operating system process other events
        DoEvents
    Loop
End Sub
Public Sub RegHotkey()
    bCancel = False
    
    'register the Ctrl-Alt-End hotkey
    RegisterHotKey Parent.hWnd, &HBFFF&, MOD_CONTROL Or MOD_ALT, vbKeyEnd
    
    'process the Hotkey messages
    ProcessMessages
End Sub

Public Sub unRegHotkey()
    bCancel = True
    'unregister hotkey
    Call UnregisterHotKey(Parent.hWnd, &HBFFF&)
End Sub
