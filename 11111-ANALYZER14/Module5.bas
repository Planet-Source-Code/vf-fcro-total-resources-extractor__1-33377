Attribute VB_Name = "ModuleINFO"
'VERSION
Public Type VS_FIXEDFILEINFO
    dwSignature As Long  'Contains the value 0xFEEFO4BD. This is used with the szKey member of VS_VERSION_INFO data when searching a file for the VS_FIXEDFILEINFO structure.
    dwStrucVersion As Long ' e.g. 0x00000042 = "0.42"
    dwFileVersionMS As Long ' e.g. 0x00030075 = "3.75"
    dwFileVersionLS As Long ' e.g. 0x00000031 = "0.31"
    dwProductVersionMS As Long ' e.g. 0x00030010 = "3.10"
    dwProductVersionLS As Long ' e.g. 0x00000031 = "0.31"
    dwFileFlagsMask As Long ' = 0x3F for version "0.42"
    dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long ' e.g. VFT_DRIVER
    dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long ' e.g. 0
    dwFileDateLS As Long ' e.g. 0
End Type

Public Type INF_INF '(Version Info)
 wLength As Integer
 wValueLength As Integer
 wType As Integer
End Type

Dim SINF2 As String
Dim Vlng As Long
Dim VVal As Integer
Dim check1 As Byte
Dim totalLG As Long
Dim TcontX As Long
Dim llng As Long
Dim SINF As String
Dim tmpINF As INF_INF
Public FFInfo As VS_FIXEDFILEINFO

Public Sub GetFileInfo(data() As Byte)
Dim countX As Long

CopyMemory tmpINF, data(0), Len(tmpINF)
countX = countX + Len(tmpINF)

Dim FINF As String
llng = lstrlenW(ByVal VarPtr(data(countX)))
FINF = Space(llng)
CopyMemory ByVal StrPtr(FINF), data(countX), llng * 2
countX = countX + llng * 2 + 2
If (countX Mod 4) <> 0 Then countX = countX + 2


OINF = "Length:" & tmpINF.wLength & " (" & Hex(tmpINF.wLength) & "h) Bytes" & vbCrLf
If tmpINF.wType = 1 Then
OINF = OINF & "Text Resource Version" & vbCrLf
Else
OINF = OINF & "Binary Resource Version" & vbCrLf
End If
OINF = OINF & FINF & vbCrLf

If Not (Not CBool(tmpINF.wValueLength)) Then
CopyMemory FFInfo, data(countX), Len(FFInfo)
countX = countX + Len(FFInfo)
OINF = OINF & "Signature:" & Hex(FFInfo.dwSignature) & "h" & vbCrLf
OINF = OINF & "File Version:" & FormatVER(FFInfo.dwFileVersionMS, FFInfo.dwFileVersionLS) & vbCrLf
OINF = OINF & "Product Version:" & FormatVER(FFInfo.dwProductVersionMS, FFInfo.dwProductVersionLS) & vbCrLf
OINF = OINF & "File OS:" & GetOS(FFInfo.dwFileOS) & vbCrLf
OINF = OINF & "File Type:" & GetFileType(FFInfo.dwFileType) & vbCrLf
End If
OINF = OINF & vbCrLf

Do While countX < ResTotLen
CopyMemory tmpINF, data(countX), Len(tmpINF)

totalLG = countX + tmpINF.wLength 'Ukupna duzina INFO-a

countX = countX + Len(tmpINF)

If tmpINF.wLength = 0 And tmpINF.wType = 0 And tmpINF.wValueLength = 0 Then GoTo eend

tcountx = countX
llng = lstrlenW(ByVal VarPtr(data(countX)))
SINF = Space(llng)
CopyMemory ByVal StrPtr(SINF), data(countX), llng * 2

OINF = OINF & SINF & ":" & vbCrLf
If SINF = "StringFileInfo" Then
countX = countX + llng * 2 + 2
If (countX Mod 4) <> 0 Then countX = countX + 2
countX = GetStrFilInf(data, countX, totalLG)

ElseIf SINF = "VarFileInfo" Then
countX = countX + llng * 2 + 2
If (countX Mod 4) <> 0 Then countX = countX + 2
countX = GetVarFilInf(data, countX)
End If

If (countX Mod 4) <> 0 Then countX = countX + 2
eend:
Loop
End Sub
Public Function GetVarFilInf(data() As Byte, ByVal countX As Long) As Long
CopyMemory tmpINF, data(countX), Len(tmpINF)
countX = countX + Len(tmpINF)
llng = lstrlenW(ByVal VarPtr(data(countX)))
SINF = Space(llng)
CopyMemory ByVal StrPtr(SINF), data(countX), llng * 2
OINF = OINF & SINF & ":"
countX = countX + llng * 2 + 2
If (countX Mod 4) <> 0 Then countX = countX + 2

Vlng = tmpINF.wValueLength / 2
For u = 1 To Vlng
CopyMemory VVal, data(countX), 2
OINF = OINF & " " & Hex(VVal) & "h"
countX = countX + 2
Next u
OINF = OINF & vbCrLf
If (countX Mod 4) <> 0 Then countX = countX + 2
GetVarFilInf = countX
OINF = OINF & vbCrLf
End Function


Public Function GetStrFilInf(data() As Byte, ByVal countX As Long, ByVal length As Long) As Long
CopyMemory tmpINF, data(countX), Len(tmpINF)
countX = countX + Len(tmpINF)
If tmpINF.wLength = 0 And tmpINF.wType = 0 And tmpINF.wValueLength = 0 Then GoTo dalje
llng = lstrlenW(ByVal VarPtr(data(countX)))
SINF = Space(llng)
CopyMemory ByVal StrPtr(SINF), data(countX), llng * 2
OINF = OINF & SINF & vbCrLf
countX = countX + llng * 2 + 2
If (countX Mod 4) <> 0 Then countX = countX + 2
dalje:
Do While countX < length
CopyMemory tmpINF, data(countX), Len(tmpINF)
tcountx = countX + tmpINF.wLength
countX = countX + Len(tmpINF)
llng = lstrlenW(ByVal VarPtr(data(countX)))
SINF = Space(llng)
CopyMemory ByVal StrPtr(SINF), data(countX), llng * 2
OINF = OINF & SINF
countX = countX + llng * 2 + 2
If (countX Mod 4) <> 0 Then countX = countX + 2
If countX = tcountx Then OINF = OINF & ":" & vbCrLf: GoTo nemadalje
llng = lstrlenW(ByVal VarPtr(data(countX)))
SINF = Space(llng)
CopyMemory ByVal StrPtr(SINF), data(countX), llng * 2
OINF = OINF & ":" & SINF & vbCrLf
nemadalje:
countX = tcountx
If (countX Mod 4) <> 0 Then countX = countX + 2
Loop
OINF = OINF & vbCrLf
GetStrFilInf = length
End Function
Public Function FormatVER(ByVal Hvalue As Long, ByVal Lvalue As Long) As String
FormatVER = (Hvalue And &HFFFF0000) / &H10000
FormatVER = FormatVER & "." & (Hvalue And &HFFFF&) & "."
FormatVER = FormatVER & (Lvalue And &HFFFF0000) / &H10000 & "."
FormatVER = FormatVER & (Lvalue And &HFFFF&)
End Function

Public Function GetOS(ByVal value As Long) As String
If value = 0 Then GetOS = "Unknow": Exit Function
If (value And &H10000) = &H10000 Then GetOS = GetOS & "Dos_"
If (value And &H1&) = &H1& Then GetOS = GetOS & "Windows16_"
If (value And &H4&) = &H4& Then GetOS = GetOS & "Windows32_"
If (value And &H40000) = &H40000 Then GetOS = GetOS & "NT_"
If value = &H20000 Then GetOS = "OS/2-16_"
If value = &H20002 Then GetOS = "OS/2-16_PM16_"
If value = &H30000 Then GetOS = "OS/2-32_"
If value = &H30002 Then GetOS = "OS/2-32_PM32_"
GetOS = Left(GetOS, Len(GetOS) - 1)
End Function

Public Function GetFileType(ByVal value As Long) As String
Select Case value
Case 1
GetFileType = "Application"
Case 2
GetFileType = "DLL (Dynamic Link Library)"
Case 3
GetFileType = "Driver"
Case 4
GetFileType = "Font"
Case 5
GetFileType = "Virtual Device"
Case 7
GetFileType = "SLL (Static Link Library)"
Case 0
GetFileType = "Unknow"
End Select
End Function


'Accelerator TABLE

Public Sub GetAccelInfo(data() As Byte)
Dim countX As Long
Dim VVal As Integer

OINF = "Resource ID:" & GRPinfo & vbCrLf
OINF = OINF & "Length:" & ResTotLen & " (" & Hex(ResTotLen) & "h) Bytes" & vbCrLf
OINF = OINF & "Total Accelerator Keys:" & (UBound(data) + 1) / 8 & vbCrLf & vbCrLf
For u = 0 To UBound(data) / 8
CopyMemory VVal, data(countX + 2), 2 'Uzmi Key
OINF = OINF & GetKEY(VVal)
OINF = OINF & ", System ID:" & CLng(data(countX + 4) + CLng(data(countX + 5)) * 256)
CopyMemory VVal, data(countX), 2 'Uzmi Flag
OINF = OINF & ", " & GetAccelFLAG(VVal) & vbCrLf
countX = countX + 8
Next u


End Sub

Public Function GetAccelFLAG(ByVal value As Integer) As String
If (value And &H1) = &H1 Then GetAccelFLAG = GetAccelFLAG & "VIRTKEY_"
If (value And &H2) = &H2 Then GetAccelFLAG = GetAccelFLAG & "NOINVERT_"
If (value And &H4) = &H4 Then GetAccelFLAG = GetAccelFLAG & "SHIFT_"
If (value And &H8) = &H8 Then GetAccelFLAG = GetAccelFLAG & "CTRL_"
If (value And &H10) = &H10 Then GetAccelFLAG = GetAccelFLAG & "ALT_"
If value = &H80 Then GetAccelFLAG = "_"
GetAccelFLAG = Left(GetAccelFLAG, Len(GetAccelFLAG) - 1)
End Function

Public Function GetKEY(ByVal value As Integer) As String
Select Case value
Case 1
GetKEY = "Left Mouse Button"
Case 2
GetKEY = "Right Mouse Button"
Case 3
GetKEY = "Cancel Key"
Case 4
GetKEY = "Middle Mouse Button"
Case 8
GetKEY = "BackSpace Key"
Case 9
GetKEY = "Tab Key"
Case 12
GetKEY = "Clear Key"
Case 13
GetKEY = "Return Key"
Case 16
GetKEY = "Shift Key"
Case 17
GetKEY = "Control Key"
Case 18
GetKEY = "Menu Key"
Case 19
GetKEY = "Pause Key"
Case 20
GetKEY = "Caps Lock Key"
Case 21
GetKEY = "Kana Key"
Case 23
GetKEY = "Junja Key"
Case 24
GetKEY = "Final Key"
Case 25
GetKEY = "Hanja Key"
Case 27
GetKEY = "Escape Key"
Case 28
GetKEY = "Convert Key"
Case 29
GetKEY = "Non Convert Key"
Case 30
GetKEY = "Accept Key"
Case 31
GetKEY = "Mode Change Key"
Case 32
GetKEY = "Space Key"
Case 33
GetKEY = "Prior Key"
Case 34
GetKEY = "Next Key"
Case 35
GetKEY = "End Key"
Case 36
GetKEY = "Home Key"
Case 37
GetKEY = "Left Key"
Case 38
GetKEY = "Up Key"
Case 39
GetKEY = "Right Key"
Case 40
GetKEY = "Down Key"
Case 41
GetKEY = "Select Key"
Case 42
GetKEY = "Print Key"
Case 43
GetKEY = "Execute Key"
Case 44
GetKEY = "Snapshot Key"
Case 45
GetKEY = "Insert Key"
Case 46
GetKEY = "Delete Key"
Case 47
GetKEY = "Help Key"
Case 48 To 57
GetKEY = Chr(CByte(value)) & " Key"
Case 65 To 90
GetKEY = Chr(CByte(value)) & " Key"
Case 91
GetKEY = "Lwin Key"
Case 92
GetKEY = "Rwin Key"
Case 93
GetKEY = "Apps Key"
Case 96 To 105
GetKEY = "Numpad" & (value - 96) & " Key"
Case 106
GetKEY = "Multiply Key"
Case 107
GetKEY = "Add Key"
Case 108
GetKEY = "Separator Key"
Case 109
GetKEY = "Substract Key"
Case 110
GetKEY = "Decimal Key"
Case 111
GetKEY = "Divide Key"
Case 112 To 135
GetKEY = "F" & (value - 111) & " Key"
Case 144
GetKEY = "Numlock Key"
Case 145
GetKEY = "Scroll Key"
Case 160
GetKEY = "LShift Key"
Case 161
GetKEY = "RShift Key"
Case 162
GetKEY = "LControl Key"
Case 163
GetKEY = "RControl Key"
Case 164
GetKEY = "LMenu Key"
Case 165
GetKEY = "RMenu Key"
Case Else
GetKEY = (value And &HFF) & " Key"
End Select

End Function


