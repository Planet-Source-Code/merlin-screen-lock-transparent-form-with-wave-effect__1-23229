Attribute VB_Name = "MStart"
Public cTray As New ClsTray
Public cWave As New ClsPicWave
Public cTrans As New ClsTranslution
Public cSet As New ClsSettings

Public Sub Main()

'Prüft ob Anwendung schon besteht
If App.PrevInstance = True Then
    MsgBox "Screen Lock is always running.", vbExclamation Or vbOKOnly, "Screen Lock Warning"
    Exit Sub
End If

Load Parent

End Sub

