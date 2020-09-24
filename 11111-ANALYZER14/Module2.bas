Attribute VB_Name = "Module2"
Public OverOrRen As Boolean

Public Rprop1 As String
Public Rprop2 As String
Public Rprop3 As String

Public OINF As String
Public ERRORX As Boolean
Public WorkingFile As String
Public WorkingResFile As String
Public Type RTSTRx
id As Long
data As String
End Type
Public RTTEXT() As RTSTRx
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Const IcoFilter = "Icon (*.ico)" & vbNullChar & "*.ico"
Public Const CurFilter = "Cursor (*.cur)" & vbNullChar & "*.cur"


Declare Function CreateStreamOnHGlobal Lib "ole32" _
                              (ByVal hGlobal As Long, _
                              ByVal fDeleteOnRelease As Long, _
                              ppstm As Any) As Long

Declare Function OleLoadPicture Lib "olepro32" _
                              (pStream As Any, _
                              ByVal lSize As Long, _
                              ByVal fRunmode As Long, _
                              riid As GUID, _
                              ppvObj As Any) As Long

Public Type GUID
  dwData1 As Long
  wData2 As Integer
  wData3 As Integer
  abData4(7) As Byte
End Type



Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As GUID) As Long
Const sIID_IPicture = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
Const GMEM_MOVEABLE = &H2
Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByValdwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long


'BITMAP & KURSOR & ICON -temporary data--privremeno
Public TMPDATAX() As Byte
'Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, lpInfoHeader As Long, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As Long, ByVal wUsage As Long) As Long
Public SNGDATAX() As Byte


Public PicWidth() As Long
Public PicHeight() As Long
Public STD1() As New StdPicture
'Picture



Public Type BITMAPFILEHEADER   ' 14 Bytes
     bfType As Integer
     bfSize As Long
     bfReserved1 As Integer
     bfReserved2 As Integer
     bfOffBits As Long
End Type


Public Type BITMAPINFOHEADER '  40 bytes
     biSize As Long
     biWidth As Long
     biHeight As Long
     biPlanes As Integer
     biBitCount As Integer
     biCompression As Long
     biSizeImage As Long
     biXPelsPerMeter As Long
     biYPelsPerMeter As Long
     biClrUsed As Long
     biClrImportant As Long
End Type
Public WavL As Long
Public RESTYPE As New Collection
Public RESNAME As New Collection
Public RESTYPENAME As New Collection
Public RESLANGID As New Collection


Public EXPLIST As New Collection

Declare Function PlaySound_Res Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszname As Long, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Public Const LB_SETTABSTOPS = &H192
Public Const LB_ITEMFROMPOINT = &H1A9

Public lIndex As Long 'koji je index klikut!

Declare Function CreateIconFromResource Lib "user32" (presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As Long) As Long
Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As Long, ByVal lpType As Long) As Long
Declare Function FindResourceEx Lib "kernel32" Alias "FindResourceExA" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal wLanguage As Integer) As Long
Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Declare Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumResourceLanguages Lib "kernel32" Alias "EnumResourceLanguagesA" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal hUpdate As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal wLanguage As Integer, lpData As Any, ByVal cbData As Long) As Long
Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" (ByVal hUpdate As Long, ByVal fDiscard As Long) As Long


Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Type PictureDescription
    cbSizeofStruct As Long
    PicType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type


Const DIFFERENCE = 11
Const RT_ACCELERATOR = 9&
Const RT_ANICURSOR = (21)
Const RT_ANIICON = (22)
Const RT_BITMAP = 2&
Const RT_CURSOR = 1&
Const RT_DIALOG = 5&
Const RT_DLGINCLUDE = (17)
Const RT_FONT = 8&
Const RT_FONTDIR = 7&
Const RT_ICON = 3&
Const RT_GROUP_CURSOR = (RT_CURSOR + DIFFERENCE)
Const RT_GROUP_ICON = (RT_ICON + DIFFERENCE)
Const RT_HTML = (23)
Const RT_MENU = 4&
Const RT_MESSAGETABLE = (11)
Const RT_PLUGPLAY = (19)
Const RT_RCDATA = 10&
Const RT_STRING = 6&
Const RT_VERSION = (16)
Const RT_VXD = (20)
Public Xcnt As Long

Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictureDescription, riid As GUID, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Declare Function CreateIconFromResourceEx Lib "user32" (presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal uFlags As Long) As Long
Const LR_LOADMAP3DCOLORS = &H1000
Const LR_LOADTRANSPARENT = &H20

Public Function GetIconToPicture(data() As Byte) As IPicture
Dim hMem  As Long
Dim lpMem  As Long
hMem = GlobalAlloc(GMEM_MOVEABLE, UBound(data) + 1)
lpMem = GlobalLock(hMem)
CopyMemory ByVal lpMem, data(0), UBound(data) + 1
Dim IID_IPicture As GUID
Call CLSIDFromString(StrPtr(sIID_IPicture), IID_IPicture)
Dim hIcon As Long
hIcon = CreateIconFromResourceEx(ByVal (lpMem + &H16), UBound(data) + 1, 1, &H30000, 0, 0, LR_LOADMAP3DCOLORS)
Call GlobalUnlock(hMem)
Dim tPicConv As PictureDescription
With tPicConv
.cbSizeofStruct = Len(tPicConv)
.PicType = vbPicTypeIcon
.hImage = hIcon
End With
OleCreatePictureIndirect tPicConv, IID_IPicture, True, GetIconToPicture
Call GlobalFree(hMem)
DestroyIcon hIcon
End Function
Function EnumRSLang(ByVal hModule As Long, ByVal lpsztype As Long, ByVal lpszname As Long, ByVal lpszID As Integer, ByVal lParam As Long) As Long
Dim TypeX As String
Dim NameX As String
LangID = lpszID
TypeX = GetStringFromPointer(lpsztype)
NameX = GetStringFromPointer(lpszname)
SetPropName lpsztype, TypeX
RESTYPE.Add TypeX
RESNAME.Add NameX
RESLANGID.Add lpszID
EnumRSLang = 1
End Function

Function EnumRSType(ByVal hModule As Long, ByVal lpsztype As Long, ByVal lParam As Long) As Long
Call EnumResourceNames(hModule, lpsztype, AddressOf EnumRSName, Xcnt)
EnumRSType = 1
End Function
Function EnumRSName(ByVal hModule As Long, ByVal lpsztype As Long, ByVal lpszname As Long, ByVal lParam As Long) As Long
Call EnumResourceLanguages(HMOD, lpsztype, lpszname, AddressOf EnumRSLang, 0)
EnumRSName = 1
End Function
Public Function LongToInt(ByVal value As Long) As Integer
CopyMemory LongToInt, ByVal VarPtr(value), 2
End Function
Public Function IntToLong(ByVal value As Integer) As Long
CopyMemory ByVal VarPtr(IntToLong), value, 2
End Function
Public Sub ClearCOLLECTION()
Set RESTYPE = Nothing
Set RESNAME = Nothing
Set RESTYPENAME = Nothing
Set RESLANGID = Nothing
End Sub
Function GetStringFromPointer(ByVal point As Long) As String
Dim lLen As Long
If (point > &HFFFF&) Or (point < 0) Then
lLen = lstrlen(point)
GetStringFromPointer = Space(lLen)
CopyMemory ByVal GetStringFromPointer, ByVal point, lLen
Else
GetStringFromPointer = CStr(point)
End If
End Function
Public Function SetStringNameEx(data() As Byte, ByVal entry As Long) As RTSTRx()
On Error GoTo eRe
Dim tmpBFR() As RTSTRx
ReDim tmpBFR(15)
Dim CountY As Integer
Dim countX As Long
Dim CHECKLNG As Integer
Dim CHECKLNG1 As Long
entry = (entry - 1) * 16
For u = 0 To 15
CopyMemory CHECKLNG, OtherData(countX), 2
CHECKLNG1 = IntToLong(CHECKLNG)
If CHECKLNG1 = 0 Then countX = countX + 2: GoTo dalje
If CHECKLNG1 > ResTotLen Then GoTo eRe
tmpBFR(CountY).data = Space(CHECKLNG1)
'Kopiraj UNICODE sadržaj
CopyMemory ByVal StrPtr(tmpBFR(CountY).data), OtherData(countX + 2), CHECKLNG1 * 2
countX = countX + 2 + CHECKLNG1 * 2
tmpBFR(CountY).id = entry
CountY = CountY + 1
dalje:
entry = entry + 1
Next u

ReDim Preserve tmpBFR(CountY - 1)
SetStringNameEx = tmpBFR
Erase tmpBFR
'Ova metoda je puno bolja jer sa LoadString ne znamo koje je velièine nadolazeci string,pa
'ne moramo puniti string proizvoljnom velicinom!
Exit Function
eRe:
On Error GoTo 0
ERRORX = True
End Function
Private Sub SetPropName(ByVal NameX As Long, ByVal TypeX As String)
RESTYPENAME.Add NName(NameX, TypeX)
End Sub

Public Function NName(ByVal NameX As Long, Optional TypeX As String) As String
Select Case NameX
Case RT_ACCELERATOR
NName = "Accelerator Table"
Case RT_ANICURSOR
NName = "Animated Cursor"
Case RT_ANIICON
NName = "Animated Icon"
Case RT_BITMAP
NName = "Bitmap"
Case RT_CURSOR
NName = "Single Cursor"
Case RT_DIALOG
NName = "Dialog Box"
Case RT_DLGINCLUDE
NName = "DlgBox definition"
Case RT_FONT
NName = "Font"
Case RT_FONTDIR
NName = "Font directory"
Case RT_ICON
NName = "Single Icon"
Case RT_GROUP_CURSOR
NName = "Group Cursor"
Case RT_GROUP_ICON
NName = "Group Icon"
Case RT_HTML
NName = "HTML document"
Case RT_MENU
NName = "Menu"
Case RT_MESSAGETABLE
NName = "Message Table"
Case RT_PLUGPLAY
NName = "Plug and Play"
Case RT_RCDATA
NName = "RC Data"
Case RT_VERSION
NName = "Version Info"
Case RT_VXD
NName = "VXD"
Case RT_STRING
NName = "String"
Case Else
If IsNumeric(TypeX) Then
NName = CStr(NameX)
Else
NName = TypeX
End If
End Select
End Function



Public Sub SaveIcon(ByVal hWnd As Long, ByVal filt As String, ByVal Text1 As String, ByVal icon As Long)
Dim spath As String
Dim Ddata() As Byte
aa = GetSaveFilePath(hWnd, filt, 0, filt, "", "", Text1, spath)
If aa = False Then Exit Sub
If Dir(spath) <> "" Then Kill spath
DoEvents

'Rutina za snimanje ikona!!!!!!!!!!!!!!!!!!!!!!
Dim CalcEntry As Long 'izracunaj pojedinacne EntryPointe ikona u grupi ikona
CalcEntry = Len(NEWH) + NEWH.rescountX * (Len(RSDIR(0)) + 2) 'izracunaj duzinu Headera..
'Izracunaj prvi ulaz!

Open spath For Binary As #1
Put #1, , NEWH 'Snimi header/brojac ikona
For x = 1 To NEWH.rescountX
Put #1, , RSDIR(x - 1) 'Postavi header ikone
Seek #1, Loc(1) - 1 'Vrati pointer snimanja unatrag...
Put #1, , CalcEntry 'Preko ID (integer) upiši entry point (long)
CalcEntry = CalcEntry + RSDIR(x - 1).bitesinres 'Dodaj slijedeci entry point
Next x
'Ucitaj RAW podatke clanova
For x = 1 To NEWH.rescountX
LoadIntoMemory CLng(RSDIR(x - 1).iconID), icon, Ddata 'Ikone
'Snimi RAW podatke clana
Put #1, , Ddata
Next x
' I to je to....Al me zajebavalo do boli!
Close #1
Erase Ddata
End Sub

Public Sub SaveCursor(ByVal hWnd As Long, ByVal filt As String, ByVal Text1 As String, ByVal cursor As Long)
Dim spath As String
Dim Ddata() As Byte
aa = GetSaveFilePath(hWnd, filt, 0, filt, "", "", Text1, spath)
If aa = False Then Exit Sub
If Dir(spath) <> "" Then Kill spath
DoEvents

'Rutina za snimanje ikona!!!!!!!!!!!!!!!!!!!!!!
Dim CalcEntry As Long 'izracunaj pojedinacne EntryPointe ikona u grupi kursora
CalcEntry = Len(NEWH) + NEWH.rescountX * (Len(CRSDIR(0)) + 2) 'izracunaj duzinu Headera..
'Izracunaj prvi ulaz!

Dim HOTSPOT As LocalHDR
Dim ReData() As Byte

Open spath For Binary As #1
Put #1, , NEWH 'Snimi header/brojac kursora

For x = 1 To NEWH.rescountX

'Ovdje se ne možemo poslužiti automatizmom zapisa jer je zapis u .CUR razlicit od onoga u .RES-u. (Bezveze!!koji je M$ racku)
Put #1, , CByte(CRSDIR(x - 1).curresD.width)
Put #1, , CByte(CRSDIR(x - 1).curresD.height)
Put #1, , CByte(0)
Put #1, , CByte(0)
Put #1, , CRSDIR(x - 1).hotXY 'Ali cemo se vratiti na ovu poziciju!!!
Put #1, , CRSDIR(x - 1).bitesinres - 4 'S obzirom da ne bilježimo HotSpot iz Res-a smanjujemo za njegovu strukturu (-4)
Put #1, , CalcEntry 'Preko ID (integer) upiši entry point (long)
CalcEntry = CalcEntry + CRSDIR(x - 1).bitesinres - 4 'Dodaj slijedeci entry point

LoadIntoMemory CLng(CRSDIR(x - 1).cursorID), cursor, Ddata 'nažalost moramo dvaput ucitavati u memoriju jer moramo procitati pravilan HotSpot iz Res-a
CopyMemory HOTSPOT, Ddata(0), Len(HOTSPOT) 'uzmi hotspot
Seek #1, Loc(1) - 11 'Vrati se na polozaj hotspota
Put #1, , HOTSPOT 'Upisi pravi hotspot!
Seek #1, LOF(1) + 1 'Vrati na pravi polozaj
Next x

'Dakle to je bio Header....
'Ucitaj RAW podatke clanova

For x = 1 To NEWH.rescountX
LoadIntoMemory CLng(CRSDIR(x - 1).cursorID), cursor, Ddata 'Ikone
'Snimi RAW podatke clana
ReDim ReData(UBound(Ddata) - 4) 'Redimenzioniraj jer ne bilježimo HOTSPOT
CopyMemory ReData(0), Ddata(4), UBound(ReData) + 1
Put #1, , ReData
Next x

'Moram priznati da sam se stvarno namucio sa kursorima i ikonama....
'Sto ce tek biti kada to sve bude trebalo konvertirati u RES????
Close #1
Erase Ddata
Erase ReData
End Sub

Public Function GetPicture(dataXY() As Byte) As IPicture
Dim hMem  As Long
Dim lpMem  As Long
Dim IID_IPicture As GUID
Dim istm As stdole.IUnknown
Dim ipic As IPicture
hMem = GlobalAlloc(GMEM_MOVEABLE, UBound(dataXY) + 1)
lpMem = GlobalLock(hMem)
CopyMemory ByVal lpMem, dataXY(0), UBound(dataXY) + 1
Call GlobalUnlock(hMem)
Call CreateStreamOnHGlobal(hMem, 1, istm)
Call CLSIDFromString(StrPtr(sIID_IPicture), IID_IPicture)
Call OleLoadPicture(ByVal ObjPtr(istm), UBound(dataXY) + 1, 0, IID_IPicture, GetPicture)
Call GlobalFree(hMem)
End Function
Public Sub SaveToRes(ByVal spath As String, ByVal listX As ListBox)
Open spath For Binary As #1


Dim TypeX1 As Long
Dim TypeX2() As Byte
Dim NameX1 As Long
Dim NameX2() As Byte
Dim TEMPNAME As String
Dim MEMCONT() As Byte

PutRESheader


For u = 0 To listX.ListCount - 1

Dim ret() As Long
ret = NameType(CStr(RESNAME.item(EXPLIST.item(u + 1))), CStr(RESTYPE.item(EXPLIST.item(u + 1))), NameX2, TypeX2)
TypeX1 = ret(0)
NameX1 = ret(1)
Erase ret


If TypeX1 = 14 Or TypeX1 = 12 Then
'Ako je ikona ili kursor----Jebem M$ Umjesto da ide prvo GROUP ikona i kursora pa tek SINGLE,,,oni to bilježe obrnuto!!!!

'Zapamti o kojem se tipu radi....
Dim oldTYPE As Long
If TypeX1 = 14 Then
oldTYPE = 3
ElseIf TypeX1 = 12 Then
oldTYPE = 1
End If

Dim MEM1() As Byte

PutHeadMem NameX1, NameX2, TypeX1, TypeX2, MEM1, True

'Uzmi broj clanova
Dim tmpHDR As NewHDR
Dim tmpRSDIR() As ResDIR
CopyMemory tmpHDR, MEM1(0), Len(tmpHDR)
ReDim tmpRSDIR(tmpHDR.rescountX - 1)
For x = 1 To tmpHDR.rescountX
'Popuni Clanove sa informacijama
CopyMemory tmpRSDIR(x - 1), MEM1(6 + (x - 1) * Len(RSDIR(0))), Len(RSDIR(0))
Next x

For x = 1 To tmpHDR.rescountX
'Zapisi Single Icone / Cursor
Dim ret2() As Long
ret2 = NameType(tmpRSDIR(x - 1).iconID, CStr(oldTYPE), NameX2, TypeX2)
TypeX1 = ret2(0)
NameX1 = ret2(1)
Erase ret2
PutHeadMem NameX1, NameX2, TypeX1, TypeX2, MEMCONT, , CInt(RESLANGID.item(EXPLIST.item(u + 1)))
Erase TypeX2
Erase NameX2
Erase MEMCONT
Next x

'I na kraju zapisi GROUP....Da,da,,zašto ne bi malo komplicirali i ljude zajebavali sa svojim cudnim formatima (M$)!!!!

ret2 = NameType(CStr(RESNAME.item(EXPLIST.item(u + 1))), CStr(RESTYPE.item(EXPLIST.item(u + 1))), NameX2, TypeX2)
TypeX1 = ret2(0)
NameX1 = ret2(1)
Erase ret2
PutHeadMem NameX1, NameX2, TypeX1, TypeX2, MEMCONT, , CInt(RESLANGID.item(EXPLIST.item(u + 1)))
Erase TypeX2
Erase NameX2
Erase MEMCONT



Else
'Zapisi sve ostalo!...Normalnim putem....
PutHeadMem NameX1, NameX2, TypeX1, TypeX2, MEMCONT, , CInt(RESLANGID.item(EXPLIST.item(u + 1)))
Erase TypeX2
Erase NameX2
Erase MEMCONT
End If

Next u



Close #1
End Sub
Public Sub PutRESheader()
'PRE-HEADER
Put #1, , CLng(0)
Put #1, , CLng(&H20)
Put #1, , CLng(&HFFFF&)
Put #1, , CLng(&HFFFF&)
Put #1, , CLng(0)
Put #1, , CLng(0)
Put #1, , CLng(0)
Put #1, , CLng(0)
'END OF PRE-HEADER
End Sub

Public Sub PutHeadMem(ByVal NameX1 As Long, NameX2() As Byte, ByVal TypeX1 As Long, TypeX2() As Byte, MEMCONT() As Byte, Optional OnlyLoadMem As Boolean, Optional LANGX As Integer, Optional OnlySaveName As Boolean)
Dim ResHedLen As Long 'Resource Header length
Dim nameQ As Boolean
Dim typeQ As Boolean
Dim HOBJ1 As Long
Dim TTB1 As Long
Dim TTF1 As Long
Dim SZ1 As Long
'Dword Alignment dimenzion
'Svaki resource (bez glavnog pre-headera) mora biti djeljiv sa 4?!



Dim Resst2 As Long
Dim Resst As Long

If OnlySaveName = True Then SZ1 = UBound(MEMCONT) + 1: GoTo rdalje

'HOBJ1 = FindResource(HMOD, NAMEX1, TYPEX1)
HOBJ1 = FindResourceEx(HMOD, TypeX1, NameX1, LANGX)
TTB1 = LoadResource(HMOD, HOBJ1)
TTF1 = LockResource(TTB1)
SZ1 = SizeofResource(HMOD, HOBJ1)
ReDim MEMCONT(SZ1 - 1)
CopyMemory MEMCONT(0), ByVal TTF1, SZ1

If OnlyLoadMem Then Exit Sub

rdalje:
ResHedLen = 24
If (NameX1 < 0) Or (NameX1 > &HFFFF&) Then
ResHedLen = ResHedLen + (lstrlen(VarPtr(NameX2(0))) + 1) * 2
nameQ = True
Else
ResHedLen = ResHedLen + 4
End If
If (TypeX1 < 0) Or (TypeX1 > &HFFFF&) Then
ResHedLen = ResHedLen + (lstrlen(VarPtr(TypeX2(0))) + 1) * 2
typeQ = True
Else
ResHedLen = ResHedLen + 4
End If
Put #1, , SZ1
Resst = ResHedLen Mod 4
If Resst <> 0 Then
ResHedLen = ResHedLen + Resst
End If
Put #1, , ResHedLen
If typeQ Then
Dim UNI1 As String
ReDim Preserve TypeX2(UBound(TypeX2) - 1)
UNI1 = StrConv(TypeX2, vbUnicode)
UNI1 = StrConv(UNI1, vbUnicode)
Put #1, , UNI1
Put #1, , CInt(0)
Else
Put #1, , CInt(&HFFFF)
Put #1, , CInt(TypeX1)
End If
If nameQ Then
Dim UNI2 As String
ReDim Preserve NameX2(UBound(NameX2) - 1)
UNI2 = StrConv(NameX2, vbUnicode)
UNI2 = StrConv(UNI2, vbUnicode)
Put #1, , UNI2
Put #1, , CInt(0)
Else
Put #1, , CInt(&HFFFF)
Put #1, , CInt(NameX1)
End If
If Resst <> 0 Then Put #1, , CInt(0)
Put #1, , CLng(0) 'Data Version
Put #1, , CInt(&H1030) 'Memory Flag
Put #1, , LANGX
Put #1, , CLng(0) 'Version
Put #1, , CLng(0) 'Characteristic
Put #1, , MEMCONT 'Put Memory Data
'Postavi da HEADER clana bude djeljiv sa 4
If ((ResHedLen + SZ1) Mod 4) <> 0 Then
Resst2 = ResHedLen + SZ1
Do While (Resst2 Mod 4) <> 0
Put #1, , CByte(0)
Resst2 = Resst2 + 1
Loop
End If
End Sub
Public Function NameType(ByVal TEMPNAME As String, ByVal TEMPTYPE As String, NameX2() As Byte, TypeX2() As Byte) As Long()
Dim tmpLNG() As Long
ReDim tmpLNG(1)
If Not IsNumeric(TEMPTYPE) Then
TypeX2 = StrConv(TEMPTYPE & Chr(CByte(0)), vbFromUnicode)
tmpLNG(0) = VarPtr(TypeX2(0))
Else
tmpLNG(0) = CLng(TEMPTYPE)
End If
If Not IsNumeric(TEMPNAME) Then
NameX2 = StrConv(TEMPNAME & Chr(CByte(0)), vbFromUnicode)
tmpLNG(1) = VarPtr(NameX2(0))
Else
tmpLNG(1) = CLng(TEMPNAME)
End If
NameType = tmpLNG
End Function


Public Function GetAppPath(ByVal strX As String) As String

If Right(strX, 1) <> "\" Then
GetAppPath = strX & "\"
Else
GetAppPath = strX
End If
End Function

