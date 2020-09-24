Attribute VB_Name = "ModuleRES"
Public RESTYPEFILE As New Collection
Public RESNAMEFILE As New Collection
Public RESTYPENAMEFILE As New Collection
Public RESLANGIDFILE As New Collection

Public ResData() As Byte
Public RESXDATA As New Collection

Public MAXSINGLEICON As Integer
Public MAXSINGLECURSOR As Integer

Public HDLFL As Long
Public HMOD2 As Long

Public PATHPATH As String
Dim Uni As String
Dim CHECKINT As Integer
Dim CHECKVAL As Long

Public Sub ClearRESCOLL()
Set RESTYPEFILE = Nothing
Set RESNAMEFILE = Nothing
Set RESTYPENAMEFILE = Nothing
Set RESLANGIDFILE = Nothing
Set RESXDATA = Nothing
End Sub


Public Function EnumRESFile() As Boolean 'False ako je krivi file!
ClearRESCOLL
Dim CountY As Long
CountY = 1

Dim tempCNTX As Long
Dim TmpBXdata() As Byte

Dim lLen As Long
Dim countX As Long
Dim VALSTRUC As Variant
Dim ext As Long
Dim SRClen As Long
MAXSINGLEICON = 0
MAXSINGLECURSOR = 0

ext = UBound(ResData) + 1
VALSTRUC = Array("0", CStr(CLng(&H20)), CStr(CLng(&HFFFF&)), CStr(CLng(&HFFFF&)), "0", "0", "0", "0")
For u = 0 To 7
CopyMemory CHECKVAL, ResData(countX), 4
If CHECKVAL <> CLng(VALSTRUC(u)) Then Exit Function
countX = countX + 4
Next u
Erase VALSTRUC

Do While countX < ext
'Uzmi dužinu resourcea

CopyMemory CHECKVAL, ResData(countX), 4
SRClen = CHECKVAL
countX = countX + 8 'Preskoci duzinu strukture
'Uzmi TYPE
countX = GetResTypeName(countX, RESTYPEFILE, ResData, True)
'Uzmi NAME
countX = GetResTypeName(countX, RESNAMEFILE, ResData)

If RESTYPEFILE.item(CountY) = "3" Then
If CInt(RESNAMEFILE.item(CountY)) > MAXSINGLEICON Then MAXSINGLEICON = CInt(RESNAMEFILE.item(CountY))
End If

If RESTYPEFILE.item(CountY) = "1" Then
If CInt(RESNAMEFILE.item(CountY)) > MAXSINGLECURSOR Then MAXSINGLECURSOR = CInt(RESNAMEFILE.item(CountY))
End If

If (countX Mod 4) <> 0 Then countX = countX + 2

countX = countX + 6 'Preskoci DataVersion i MemoryFlag

'Uzmi LangID
CopyMemory CHECKINT, ResData(countX), 2
RESLANGIDFILE.Add CHECKINT
countX = countX + 10 'Preskoci Version i Characteristic

tempCNTX = countX
countX = countX + SRClen

Do While (countX Mod 4) <> 0
countX = countX + 1
Loop


ReDim TmpBXdata(countX - tempCNTX - 1)
CopyMemory TmpBXdata(0), ResData(tempCNTX), countX - tempCNTX
RESXDATA.Add TmpBXdata
Erase TmpBXdata


CountY = CountY + 1
Loop
Erase ResData
EnumRESFile = True
End Function
Public Sub MakeDllRes(listX As ListBox)
FreeLibrary HMOD2
GetEmptyDll PATHPATH
HDLFL = BeginUpdateResource(PATHPATH, 0)


For u = 0 To listX.ListCount - 1
PutResource RESTYPEFILE.item(u + 1), RESNAMEFILE.item(u + 1), RESLANGIDFILE.item(u + 1), ResData, u + 1
Next u

Call EndUpdateResource(HDLFL, 0)
HMOD2 = LoadLibrary(PATHPATH)
End Sub



Public Function GetResTypeName(ByVal countX As Long, ByVal TN As Collection, data() As Byte, Optional TYP As Boolean) As Long

CopyMemory CHECKINT, data(countX), 2
If CHECKINT = CInt(&HFFFF) Then
'Tada je broj
countX = countX + 2
CopyMemory CHECKINT, data(countX), 2
TN.Add CHECKINT
'Za TYPE
If TYP Then RESTYPENAMEFILE.Add NName(CLng(CHECKINT), CStr(CHECKINT))
countX = countX + 2
Else
lLen = lstrlenW(ByVal VarPtr(data(countX)))
Uni = Space(lLen)
CopyMemory ByVal StrPtr(Uni), ByVal VarPtr(data(countX)), lLen * 2
TN.Add Uni
If TYP Then RESTYPENAMEFILE.Add Uni
countX = countX + lLen * 2 + 2
End If
GetResTypeName = countX
End Function

Public Sub GetEmptyDll(ByVal filename As String)
Dim data() As Byte
data = LoadResData("DLL", "DLL")
If Dir(filename) <> "" Then Kill filename
Open filename For Binary As #1
Put #1, , data
Close #1
Erase data
End Sub

Public Sub GetDataFile(ByVal filename As String, data() As Byte)
Open filename For Binary As #1
ReDim data(LOF(1) - 1)
Get #1, , data
Close #1
End Sub
Public Sub PutResource(TypeX As String, NameX As String, Lid As Integer, data() As Byte, position As Long)
Dim TypeX1 As Long
Dim NameX1 As Long
Dim NameX2() As Byte
Dim TypeX2() As Byte
Dim ret() As Long
ret = NameType(NameX, TypeX, NameX2, TypeX2)
TypeX1 = ret(0)
NameX1 = ret(1)
Dim dXdata() As Byte
dXdata = RESXDATA.item(position)
Call UpdateResource(HDLFL, TypeX1, NameX1, Lid, ByVal VarPtr(dXdata(0)), UBound(dXdata) + 1)
Erase dXdata
End Sub
Public Sub DeleteResource(TypeX As String, NameX As String, Lid As Integer)
Dim TypeX1 As Long
Dim NameX1 As Long
Dim NameX2() As Byte
Dim TypeX2() As Byte
Dim ret() As Long
ret = NameType(NameX, TypeX, NameX2, TypeX2)
TypeX1 = ret(0)
NameX1 = ret(1)
Call UpdateResource(HDLFL, TypeX1, NameX1, Lid, ByVal 0&, 0)
End Sub

