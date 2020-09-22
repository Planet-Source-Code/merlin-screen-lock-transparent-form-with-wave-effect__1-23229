VERSION 5.00
Begin VB.Form FPassword 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Screen Lock - Change Password"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   6000
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Confirm New Password"
      Height          =   615
      Left            =   3120
      TabIndex        =   4
      Top             =   840
      Width           =   2775
      Begin VB.TextBox Text3 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Please Enter New Password"
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Please Enter Old Password"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "FPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Select Case Text1.Enabled

    Case Is = True
        If Text1.Text = cSet.Password Then
            Text1.Enabled = False
            Frame1.Enabled = False
            Text2.Enabled = True
            Frame2.Enabled = True
            Text3.Enabled = True
            Frame3.Enabled = True
            Text2.SetFocus
        Else
            MsgBox "Incorrect Password", vbExclamation Or vbOKOnly, "Screen Lock - Change Password"
            Text1.Text = ""
            Text1.SetFocus
        End If
    
    Case Is = False
        
        If Text2.Text <> Text3.Text Then
            MsgBox "Different Passwords", vbExclamation Or vbOKOnly, "Screen Lock - Change Password"
        ElseIf Text1.Text = cSet.Password Or cSet.Password = Chr(10) & Chr(11) & Chr(75) Then
            
            If Text2.Text = "" Then
                cSet.Password = Chr(10) & Chr(11) & Chr(75)
            Else
                cSet.Password = Text2.Text
            End If
            
            MsgBox "Password successful changed.", vbInformation Or vbOKOnly, "Screen Lock - Change Password"
            Unload Me
        End If

End Select
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text1.Text = cSet.Password Or cSet.Password = Chr(10) & Chr(11) & Chr(75) Then
            Text1.Enabled = False
            Frame1.Enabled = False
            Text2.Enabled = True
            Frame2.Enabled = True
            Text3.Enabled = True
            Frame3.Enabled = True
            Text2.SetFocus
        Else
            MsgBox "Incorrect Password", vbExclamation Or vbOKOnly, "Screen Lock - Change Password"
            Text1.Text = ""
            Text1.SetFocus
        End If
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text2.Text <> Text3.Text Then
            MsgBox "Different Passwords", vbExclamation Or vbOKOnly, "Screen Lock - Change Password"
        ElseIf Text1.Text = cSet.Password Or cSet.Password = Chr(10) & Chr(11) & Chr(75) Then
            
            If Text2.Text = "" Then
                cSet.Password = Chr(10) & Chr(11) & Chr(75)
            Else
                cSet.Password = Text2.Text
            End If
            
            MsgBox "Password successful changed.", vbInformation Or vbOKOnly, "Screen Lock - Change Password"
            Unload Me
        End If
    End If
End Sub
