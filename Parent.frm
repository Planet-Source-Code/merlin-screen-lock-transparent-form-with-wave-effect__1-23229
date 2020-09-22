VERSION 5.00
Begin VB.Form Parent 
   AutoRedraw      =   -1  'True
   Caption         =   "Screen Lock by Merlin"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Parent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuLockScreen 
         Caption         =   "Lock Screen"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuLockScreenStart 
         Caption         =   "Lock Screen at Startup"
      End
      Begin VB.Menu mnuAutostart 
         Caption         =   "Autostart"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUninstall 
         Caption         =   "Uninstall"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Parent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    cTray.AddToTray Me.Icon, Me.Caption, Me
    
    cSet.DefaultSetting
    'syncronisiere Menusettings
    Parent.mnuAutostart.Checked = cSet.Autostart
    Parent.mnuLockScreenStart.Checked = cSet.AutostartSL
    
    If cSet.Alarm = True Or cSet.AutostartSL = True Then
        WaterStart
    End If
    
    RegHotkey
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If cTray.RespondToTray(x) = 2 Then
        PopupMenu mnu, , , , mnuLockScreen
    End If
    
    If cTray.RespondToTray(x) = 1 Then
        WaterStart
    End If

End Sub

Private Sub mnuAutostart_Click()
    With mnuAutostart
        If .Checked = False Then
            .Checked = True
            cSet.Autostart = True
        Else
            .Checked = False
            cSet.Autostart = False
            mnuLockScreenStart.Checked = False
            cSet.AutostartSL = False
        End If
    End With
End Sub

Private Sub mnuChangePassword_Click()
    If cSet.Password = Chr(10) & Chr(11) & Chr(75) Then
        FPassword.Text1.Enabled = False
        FPassword.Frame1.Enabled = False
        FPassword.Text2.Enabled = True
        FPassword.Frame2.Enabled = True
        FPassword.Text3.Enabled = True
        FPassword.Frame3.Enabled = True
    Else
        FPassword.Text1.Enabled = True
        FPassword.Frame1.Enabled = True
        FPassword.Text2.Enabled = False
        FPassword.Frame2.Enabled = False
        FPassword.Text3.Enabled = False
        FPassword.Frame3.Enabled = False
    End If
    FPassword.Show
End Sub

Private Sub mnuExit_Click()
    cTray.RemoveFromTray
    Set cTray = Nothing
    Set cSet = Nothing
    unRegHotkey
    Unload Me
    End
End Sub

Private Sub mnuLockScreen_Click()
    WaterStart
End Sub

Private Sub mnuLockScreenStart_Click()
    With mnuLockScreenStart
        If .Checked = False Then
            .Checked = True
            cSet.AutostartSL = True
            mnuAutostart.Checked = True
            cSet.Autostart = True
        Else
            .Checked = False
            cSet.AutostartSL = False
        End If
    End With
End Sub

Private Sub mnuUninstall_Click()
    If MsgBox("Are you sure ?", vbYesNo Or vbExclamation Or vbMsgBoxSetForeground, "Uninstall Screen Lock") = vbYes Then
        If cSet.Password = Chr(10) & Chr(11) & Chr(75) Then
            cSet.ClearSetting
            cTray.RemoveFromTray
            Set cTray = Nothing
            Set cSet = Nothing
            unRegHotkey
            Unload Me
            End
        Else
            FPassword2.Show
        End If
    
    End If
End Sub
