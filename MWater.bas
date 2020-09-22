Attribute VB_Name = "MWater"
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long


Public Sub WaterExit()
    cWave.Cancel = True
    FalseSysKey False
    ShowCursor 1
    Unload Form1Template
    Set cWave = Nothing
    Set cTrans = Nothing
    Unload Form1
    cSet.Alarm = False
End Sub


Public Sub WaterStart()
FalseSysKey True
cSet.Alarm = True
ShowCursor 0
Form1.Timer1.Enabled = True

Form1Template.Show
cTrans.tColor = RGB(0, 255, 255)
cTrans.MakeTranslucent Form1Template, cTrans.tColor
Form1Template.Visible = False

Form1.Show

Set cWave.SourcePic = Form1Template
Set cWave.DestinationPic = Form1
cWave.Amplitude = 4
cWave.WaveArt = 3
cWave.Waves = 4
cWave.Speed = 10
cWave.Wave


End Sub

Public Sub FalseSysKey(bDisabled As Boolean)
    ' Disables Control Alt Delete Breaking as well as Ctrl-Escape
    Dim x As Long
    x = SystemParametersInfo(97, bDisabled, CStr(1), 0)

End Sub

