Attribute VB_Name = "ModuleDUMP"
Public Function GetHexDump(data() As Byte, ByVal position As Long) As String
If position > UBound(data) Then Exit Function
GetHexDump = Space(76)
Dim mmax As Long
Dim plc As Integer
Mid(GetHexDump, 1, 1) = Hex(position And &HF0000000)
Mid(GetHexDump, 2, 1) = Hex(position And &HF000000)
Mid(GetHexDump, 3, 1) = Hex(position And &HF00000)
Mid(GetHexDump, 4, 1) = Hex(position And &HF0000)
Mid(GetHexDump, 5, 1) = Hex(position And &HF000&)
Mid(GetHexDump, 6, 1) = Hex(position And &HF00&)
Mid(GetHexDump, 7, 1) = Hex(position And &HF0&)
Mid(GetHexDump, 8, 1) = Hex(position And &HF&)

plc = 11
mmax = 16
If position + 16 > (UBound(data)) Then mmax = UBound(data) - position + 1
For u = position To position + mmax - 1
Mid(GetHexDump, plc, 1) = Hex(data(u) And &HF0)
Mid(GetHexDump, plc + 1, 1) = Hex(data(u) And &HF)
plc = plc + 3
Next u

For u = 1 To 16 - mmax
Mid(GetHexDump, plc, 1) = " "
Mid(GetHexDump, plc + 1, 1) = " "
plc = plc + 3
Next u

plc = plc + 2

For u = position To position + mmax - 1
If data(u) >= 32 And data(u) <= 126 Then
Mid(GetHexDump, plc, 1) = Chr$(data(u))
Else
Mid(GetHexDump, plc, 1) = Chr(Asc("."))
End If
plc = plc + 1
Next u

End Function

Public Sub PrintDump(ByVal TXT As TextBox, data() As Byte, ByVal position As Long)
TXT = ""
For u = 0 To 24
TXT = TXT & GetHexDump(data, (position + u) * 16) & vbCrLf
Next u
End Sub



