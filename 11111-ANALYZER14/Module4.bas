Attribute VB_Name = "Module4"
'MESSAGE TABLE!

Public Type MessageResBlock
 lowId As Long
 HighId As Long
 EntryPoint As Long
End Type
'Podaci o ulazima
Public MRB() As MessageResBlock

Public Function GetMessageEntries(data() As Byte) As RTSTRx()
On Error GoTo eRe
Dim TMPMSG() As RTSTRx
Dim lenEntries As Long
Dim countX As Long
Dim totalEntries As Long
Dim Xcnt As Long
Dim lLen As Integer: Dim lLen1 As Long
Dim Llent As Integer: Dim lLent1 As Long
Dim Flag As Integer 'Dali se radi o unicode (1) ili ansi (0)

'Uzmi broj Message entries-a...i sve podatke
CopyMemory lenEntries, data(0), 4
countX = 4
ReDim MRB(lenEntries - 1)
For u = 1 To lenEntries
CopyMemory MRB(u - 1), data(countX), Len(MRB(u - 1))
countX = countX + Len(MRB(u - 1))
totalEntries = totalEntries + (MRB(u - 1).HighId - MRB(u - 1).lowId) + 1
Next u

'Dimenzioniraj ukupan broj stringova
ReDim TMPMSG(totalEntries - 1)


'Pokupi stringove i identifikacije
For u = 1 To lenEntries

For uu = MRB(u - 1).lowId To MRB(u - 1).HighId
CopyMemory lLen, data(countX), 2
lLen1 = IntToLong(lLen)
CopyMemory Flag, data(countX + 2), 2
countX = countX + 4

If Not CBool(Flag) Then
'Kopiraj ANSI u UNICODE
lLent1 = lstrlen(ByVal VarPtr(data(countX)))
If lLent1 > ResTotLen Then GoTo eRe
TMPMSG(Xcnt).data = Space(lLent1)
CopyMemory ByVal TMPMSG(Xcnt).data, data(countX), lLent1
Else
'Kopiraj UNICODE sadržaj
lLent1 = lstrlenW(ByVal VarPtr(data(countX)))
If lLent1 > ResTotLen Then GoTo eRe
TMPMSG(Xcnt).data = Space(lLent1)
CopyMemory ByVal StrPtr(TMPMSG(Xcnt).data), data(countX), lLent1 * 2
End If
countX = countX + lLen1 - 4
TMPMSG(Xcnt).id = uu
Xcnt = Xcnt + 1
Next uu
Next u

GetMessageEntries = TMPMSG
Erase TMPMSG
Exit Function
eRe:
On Error GoTo 0
ERRORX = True
End Function
