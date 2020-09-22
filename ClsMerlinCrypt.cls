VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMerlinCrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Klasse:    ClsMerlinCrypt
'Author:    Merlin
'Funktion:  Cryptographie (Text- and Binär Mode, Randomize Chiffretext usw.)
'Import:    None

Public Enum eCryptArt
    Binary = 0
    Text = 1
End Enum

Public Function Decode(ByVal sText As String, Optional ByVal Flag As eCryptArt = Binary, Optional ByVal Password As String = "") As String
Dim lTimer As Long, lPassTimer As Long, lRndCut As Long, lTxtCut As Long
Dim sTimer As String, sTemp As String

lTimer = Int(Timer)
ConvertTimer lTimer, sTimer

Rnd -1
Randomize lTimer

For x = 1 To Len(sText)
    lRndCut = Int(Rnd * 256)
    lTxtCut = Asc(Mid(sText, x, 1))
    sTemp = sTemp & Chr(lRndCut Xor lTxtCut)
Next x

Timer_XOR sTimer

'timer to string
sText = Chr(Val("&h" & Left(sTimer, 2) & "&")) & Chr(Val("&h" & Right(sTimer, 2) & "&")) & sTemp
sTemp = ""

If Password = "" Then
    Rnd -1
    Randomize 101175
Else
    lPassTimer = 101175
    
    For x = 1 To Len(Password)
        lPassTimer = lPassTimer Xor Asc(Mid(Password, x, 1))
    Next x
    
    Rnd -1
    Randomize lPassTimer
End If




For x = 1 To Len(sText)
    lRndCut = Int(Rnd * 256)
    lTxtCut = Asc(Mid(sText, x, 1))
    sTemp = sTemp & Chr(lRndCut Xor lTxtCut)
Next x

If Flag = Binary Then
    Decode = sTemp
Else
    BinToTxt sTemp
    Decode = sTemp
End If

End Function

Private Sub TxtXor(ByRef sTxt As String)

    

End Sub

Private Sub BinToTxt(ByRef sString As String)
Dim sTemp As String


For x = 1 To Len(sString)
    y = Asc(Mid(sString, x, 1))
    
    If y <= 91 Then
        sTemp = sTemp & Chr(y + 33)
    
    ElseIf y >= 92 And y <= 183 Then
        sTemp = sTemp & Chr(92 + 33) & Chr(y - 92 + 33)
    
    ElseIf y >= 184 Then
        sTemp = sTemp & Chr(93 + 33) & Chr(y - 184 + 33)
    
    End If
    
Next x

sString = sTemp
End Sub

Private Sub TxtToBin(ByRef sString As String)
Dim sTemp As String


For x = 1 To Len(sString)
    y = Asc(Mid(sString, x, 1))
    
    Select Case y
            
        Case Is <= 124
            sTemp = sTemp & Chr(y - 33)
    
        Case Is = 125
            x = x + 1
            z = Asc(Mid(sString, x, 1))
            sTemp = sTemp & Chr(92 + (z - 33))
        
        Case Is = 126
            x = x + 1
            z = Asc(Mid(sString, x, 1))
            sTemp = sTemp & Chr(184 + (z - 33))
    
    End Select

Next x

sString = sTemp
End Sub

Public Function Encode(ByVal sText As String, Optional ByVal Flag As eCryptArt = Binary, Optional ByVal Password As String = "") As String
Dim lTimer As Long, lPassTimer As Long, lRndCut As Long, lTxtCut As Long
Dim sTimer As String, sTemp As String

If Flag = Binary Then
Else
    TxtToBin sText
End If



If Password = "" Then
    Rnd -1
    Randomize 101175
Else
    lPassTimer = 101175
    
    For x = 1 To Len(Password)
        lPassTimer = lPassTimer Xor Asc(Mid(Password, x, 1))
    Next x
    
    Rnd -1
    Randomize lPassTimer
End If




For x = 1 To Len(sText)
    lRndCut = Int(Rnd * 256)
    lTxtCut = Asc(Mid(sText, x, 1))
    sTemp = sTemp & Chr(lRndCut Xor lTxtCut)
Next x

sText = Mid(sTemp, 3)
sTimer = Left(sTemp, 2)
ConvertString lTimer, sTimer

sTemp = ""

Rnd -1
Randomize lTimer

For x = 1 To Len(sText)
    lRndCut = Int(Rnd * 256)
    lTxtCut = Asc(Mid(sText, x, 1))
    sTemp = sTemp & Chr(lRndCut Xor lTxtCut)
Next x

Encode = sTemp
End Function

Private Sub ConvertString(ByRef lTimer As Long, ByRef sTimer As String)
Dim lCut1 As Long, lCut2 As Long
Dim sCut1 As String, sCut2 As String

lCut1 = Asc(Left(sTimer, 1))
sCut1 = "00" & Hex(lCut1)
sCut1 = Right(sCut1, 2)

lCut2 = Asc(Right(sTimer, 1))
sCut2 = "00" & Hex(lCut2)
sCut2 = Right(sCut2, 2)

sTimer = sCut1 & sCut2
Timer_XOR sTimer

lTimer = Val("&h" & sTimer & "&")

End Sub


Private Sub ConvertTimer(ByRef lTimer As Long, ByRef sTimer As String)


    sTimer = Hex(lTimer)
    sTimer = "0000" & sTimer
    sTimer = Right(sTimer, 4)
    lTimer = Val("&h" & sTimer & "&")

End Sub


Private Sub Timer_XOR(ByRef sHex As String)
Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long

p1 = Val("&h" & Left(sHex, 1) & "h")
p2 = Val("&h" & Mid(sHex, 2, 1) & "h")
p3 = Val("&h" & Mid(sHex, 3, 1) & "h")
p4 = Val("&h" & Right(sHex, 1) & "h")

p1 = p1 Xor p4
p2 = p2 Xor p4
p3 = p3 Xor p4

sHex = Hex(p1) & Hex(p2) & Hex(p3) & Hex(p4)
End Sub

