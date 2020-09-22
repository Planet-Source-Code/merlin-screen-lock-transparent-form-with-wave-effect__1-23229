VERSION 5.00
Begin VB.Form FPassword2 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Screen Lock - Uninstall"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3030
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3030
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Please Enter Password"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2775
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "FPassword2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
 If Text1.Text <> cSet.Password Then
        MsgBox "Incorrect Password", vbExclamation Or vbOKOnly, "Screen Lock - Change Password"
        Text1.Text = ""
        Text1.SetFocus
    
    ElseIf Text1.Text = cSet.Password Then
        cSet.ClearSetting
        cTray.RemoveFromTray
        Set cTray = Nothing
        Set cSet = Nothing
        unRegHotkey
        Unload Me
        End
    
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Text1.Text <> cSet.Password Then
        MsgBox "Incorrect Password", vbExclamation Or vbOKOnly, "Screen Lock - Change Password"
        Text1.Text = ""
        Text1.SetFocus
    
    ElseIf Text1.Text = cSet.Password Then
        cSet.ClearSetting
        cTray.RemoveFromTray
        Set cTray = Nothing
        Set cSet = Nothing
        unRegHotkey
        Unload Me
        End
    
    End If
End If

End Sub
