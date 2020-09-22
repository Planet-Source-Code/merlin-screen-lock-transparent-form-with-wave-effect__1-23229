Attribute VB_Name = "MRegistry"
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Sub WriteToRegistry()
    Dim sPath As String
    sPath = App.Path
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    sPath = sPath & "Screenlock.exe"
    MakeRegistrySetting "Software\Microsoft\Windows\CurrentVersion\Run", "Screen Lock", sPath
End Sub

Public Sub DelFromRegistry()
Dim hKey As Long

Qry = RegOpenKey(&H80000002, "Software\Microsoft\Windows\CurrentVersion\Run", hKey)
    RegDeleteValue hKey, "Screen Lock"
Qry = RegCloseKey(hKey)

End Sub

Private Sub MakeRegistrySetting(RegPath As String, Title As String, Data As String)
'This will make a registry setting
On Error GoTo error
a = MakeRegFile(&H80000002, RegPath$, Title$, Data$)
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Function MakeRegFile(ByVal hKey As Long, ByVal lpszSubKey As String, ByVal sSetValue As String, ByVal sValue As String) As Boolean
'For make startup and make registry setting:  Makes the registry setting
On Error GoTo error
Dim phkResult As Long
Dim lResult As Long
Dim SA As SECURITY_ATTRIBUTES
Dim lCreate As Long

RegCreateKeyEx hKey, lpszSubKey, 0, "", REG_OPTION_NON_VOLATILE, _
KEY_ALL_ACCESS, SA, phkResult, lCreate
lResult = RegSetValueEx(phkResult, sSetValue, 0, 1, sValue, _
CLng(Len(sValue) + 1))
RegCloseKey phkResult

MakeRegFile = (lResult = ERROR_SUCCESS)

Exit Function
error:
MakeRegFile = False
End Function

