VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public cCrypt As New ClsMerlinCrypt
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long


Public Property Get Autostart() As Boolean
Dim sTemp As String, lReturn As Long
    
    sTemp = String(255, Chr(0))
    lReturn = GetProfileString("Screen Lock", "Autostart", "", sTemp, 255)
    Autostart = Left(sTemp, lReturn)
End Property

Public Property Let Autostart(Value As Boolean)
    WriteProfileString "Screen Lock", "Autostart", Value
    If Value = True Then WriteToRegistry
    If Value = False Then DelFromRegistry
End Property

Public Property Get AutostartSL() As Boolean
Dim sTemp As String, lReturn As Long
    
    sTemp = String(255, Chr(0))
    lReturn = GetProfileString("Screen Lock", "AutostartSL", "", sTemp, 255)
    AutostartSL = Left(sTemp, lReturn)
End Property

Public Property Let AutostartSL(Value As Boolean)
    WriteProfileString "Screen Lock", "AutostartSL", Value
End Property

Public Property Get Alarm() As Boolean
Dim sTemp As String, lReturn As Long
    
    sTemp = String(255, Chr(0))
    lReturn = GetProfileString("Screen Lock", "Alarm", "", sTemp, 255)
    Alarm = Left(sTemp, lReturn)
End Property


Public Property Let Alarm(Value As Boolean)
    WriteProfileString "Screen Lock", "Alarm", Value
End Property

Public Property Get Password() As String
Dim sTemp As String, lReturn As Long
    
    sTemp = String(255, Chr(0))
    lReturn = GetProfileString("Screen Lock", "Password", "", sTemp, 255)
    Password = Encode(Left(sTemp, lReturn))

End Property


Public Property Let Password(sValue As String)
    WriteProfileString "Screen Lock", "Password", Decode(sValue)
End Property

Public Sub ClearSetting()
    DefaultSetting
    WriteProfileSection "Screen Lock", ""
    DelFromRegistry
End Sub

Public Sub DefaultSetting()
Dim sTemp As String, lReturn As Long
sTemp = String(255, Chr(0))
    
    If GetProfileString("Screen Lock", "About", "", sTemp, 255) = 0 Then
        
        WriteProfileSection "Screen Lock", ""
        
        WriteProfileString "Screen Lock", "About", "Merlin"
        WriteProfileString "Screen Lock", "Password", Decode(Chr(10) & Chr(11) & Chr(75))
        WriteProfileString "Screen Lock", "Alarm", False
        WriteProfileString "Screen Lock", "Autostart", True
        WriteToRegistry
        WriteProfileString "Screen Lock", "AutostartSL", False
        
        Path = App.Path
        If Right(Path, 1) <> "\" Then Path = Path & "\"
        Shell "notepad.exe " & Path & "readme.txt", vbNormalFocus
    
    End If
End Sub

Private Function Decode(sText As String) As String
Decode = cCrypt.Decode(sText, Text)
End Function

Private Function Encode(sChiffre As String) As String
Encode = cCrypt.Encode(sChiffre, Text)
End Function

Private Sub Class_Terminate()
Set cCrypt = Nothing
End Sub
