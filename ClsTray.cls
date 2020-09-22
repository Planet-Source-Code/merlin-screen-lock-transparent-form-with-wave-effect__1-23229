VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
      'TrayModule Basic Module
      'Coded by Michael Barnathan
      'E-mail at webmaster@programmerfaq.com
      'Declare a user-defined variable to pass to the Shell_NotifyIcon
      'function.
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

      'Declare the constants for the API function. These constants can be
      'found in the header file Shellapi.h.

      'The following constants are the messages sent to the
      'Shell_NotifyIcon function to add, modify, or delete an icon from the
      'taskbar status area.
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

      'The following constant is the message sent when a mouse event occurs
      'within the rectangular boundaries of the icon in the taskbar status
      'area.
Private Const WM_MOUSEMOVE = &H200

      'The following constants are the flags that indicate the valid
      'members of the NOTIFYICONDATA data type.
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

      'The following constants are used to determine the mouse input on the
      'the icon in the taskbar status area.

      'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up

      'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

      'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" _
         Alias "Shell_NotifyIconA" _
         (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      'Dimension a variable as the user-defined data type.
Private nid As NOTIFYICONDATA

Public Sub AddToTray(TrayIcon, TrayText As String, TrayForm As Form)
         'Set the individual values of the NOTIFYICONDATA data type.
         nid.cbSize = Len(nid)
         nid.hWnd = TrayForm.hWnd
         nid.uId = vbNull
         nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         nid.uCallBackMessage = WM_MOUSEMOVE
         nid.hIcon = TrayIcon 'You can replace form1.icon with loadpicture=("icon's file name")
         nid.szTip = TrayText & vbNullChar

         'Call the Shell_NotifyIcon function to add the icon to the taskbar
         'status area.
         Shell_NotifyIcon NIM_ADD, nid
         TrayForm.Hide
End Sub
Public Sub ModifyTray(TrayIcon, TrayText As String, TrayForm As Form)
         'Set the individual values of the NOTIFYICONDATA data type.
         nid.cbSize = Len(nid)
         nid.hWnd = TrayForm.hWnd
         nid.uId = vbNull
         nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         nid.uCallBackMessage = WM_MOUSEMOVE
         nid.hIcon = TrayIcon 'You can replace form1.icon with loadpicture=("icon's file name")
         nid.szTip = TrayText & vbNullChar

         'Call the Shell_NotifyIcon function to modify the icon to the taskbar
         'status area.
         Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Public Sub RemoveFromTray()
    Shell_NotifyIcon NIM_DELETE, nid
End Sub


Public Function RespondToTray(x As Single)
          'Call this sub from the mousemove event on a form
          'Event occurs when the mouse pointer is within the rectangular
          'boundaries of the icon in the taskbar status area.
          RespondToTray = 0
          Dim Msg As Long
          Dim sFilter As String
          If Parent.ScaleMode <> 3 Then Msg = x / Screen.TwipsPerPixelX Else: Msg = x
          Select Case Msg
             Case WM_LBUTTONDOWN
             Case WM_LBUTTONUP
             Case WM_LBUTTONDBLCLK 'Left button double-clicked
             RespondToTray = 1
             Case WM_RBUTTONDOWN 'Right button pressed
             RespondToTray = 2
             Case WM_RBUTTONUP
             Case WM_RBUTTONDBLCLK
          End Select
End Function

Public Sub ShowFormAgain(TrayForm As Form)
    Call RemoveFromTray
    TrayForm.Show
End Sub




