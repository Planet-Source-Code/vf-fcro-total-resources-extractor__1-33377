Attribute VB_Name = "Module1"
Public CASSE As Variant
Public CASSE2 As Variant
Public CNTG() As String
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public HMOD As Long
Public Const LOAD_LIBRARY_AS_DATAFILE = &H2
Public Const DONT_RESOLVE_DLL_REFERENCES = &H1
Public Const WM_INITDIALOG = &H110
Private Const WM_COMMAND = &H111


'ICON & CURSOR GROUP

Public GRPinfo As String
Public NEWH As NewHDR
Public RSDIR() As ResDIR 'Icon informacije
Public CRSDIR() As CResDir 'Cursor informacije

Public Type NewHDR
 reserved As Integer
 restypeX As Integer
 rescountX As Integer
End Type
'Icon
Public Type IconResDir
 width As Byte
 height As Byte
 colorCount As Byte
 reserved As Byte
End Type

Public Type ResDIR
 iconresD As IconResDir
 planes As Integer
 bitcount As Integer
 bitesinres As Long
 iconID As Integer
End Type

'CURSOR---Malo razlicito od IconResDir-a
Public Type CursorResDir
width As Integer
height As Integer
End Type

Public Type CResDir 'Slicno kao Icon samo uzima CursorResDir kao dio informacije
curresD As CursorResDir
hotXY As Long   'Hotspot kursora
bitesinres As Long
cursorID As Integer
End Type

Public Type LocalHDR
 Xspot As Integer
 Yspor As Integer
End Type

Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" _
(ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function mciGetErrorString Lib "winmm" Alias _
"mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, _
ByVal uLength As Long) As Long

Declare Function GetShortPathName Lib "kernel32" Alias _
"GetShortPathNameA" (ByVal lpszLongPath As String, _
ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public WMOD As Byte 'Koji je RESOURCE u pitanju?
Public HMOD1 As Long
Public MHDL As Long
Public HHDL As Long

Private Declare Function EndDialog Lib "user32" ( _
    ByVal hdlg As Long, _
    ByVal nResult As Long _
) As Long

Public OtherData() As Byte 'General LOADDATA
Public LangID As Integer
Public TypePtr As Long
Public TrueType() As Byte
Public TrueName As Long
Public TrueBuffer() As Byte
Public ResTotLen As Long

Public Const MF_OWNERDRAW = &H100&
Public Const MF_BYPOSITION = &H400&

Declare Function CreateDialogIndirectParam Lib "user32" Alias "CreateDialogIndirectParamA" (ByVal hInstance As Long, lpTemplate As Any, ByVal hwndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Long
Declare Function CreateDialogParam Lib "user32" Alias "CreateDialogParamA" (ByVal hInstance As Long, ByVal lpName As Long, ByVal hwndParent As Long, ByVal lpDialogFunc As Long, ByVal lParamInit As Long) As Long
Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As Any) As Long
Declare Function LoadMenuIndirect Lib "user32" Alias "LoadMenuIndirectA" (ByVal lpMenuTemplate As Long) As Long
Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As Long) As Long
Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetSystemMenu Lib "user32" _
(ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Type AVIFILEINFO
  dwMaxBytesPerSec As Long
  dwFlags As Long
  dwCaps As Long
  dwStreams As Long
  dwSuggestedBufferSize As Long
  dwWidth As Long
  dwHeight As Long
  dwScale As Long
  dwRate As Long
  dwLength As Long
  dwEditCount As Long
  szFileType As String * 64
End Type
Public Declare Function AVIFileGetStream Lib "avifil32" (ByVal pfile As Long, ByRef ppaviStream As Long, ByVal fccType As Long, ByVal lparam As Long) As Long
Public Declare Function AVIStreamLength Lib "avifil32" (ByVal pavi As Long) As Long
Public Declare Function AVIStreamRelease Lib "avifil32" (ByVal pavi As Long) As Long
Public Declare Sub AVIFileInit Lib "avifil32" ()
Public Declare Sub AVIFileExit Lib "avifil32" ()
Public Declare Function AVIFileOpen Lib "avifil32" (ByRef ppfile As Long, ByVal szFile As String, ByVal uMode As Long, ByVal pclsidHandler As Long) As Long  'HRESULT
Public Declare Function AVIFileRelease Lib "avifil32" ( _
  ByVal pfile As Long) As Long
Public Declare Function AVIGetFileInfo Lib "avifil32" Alias _
"AVIFileInfoA" _
  (ByVal pfile As Long, _
  pfi As AVIFILEINFO, _
  ByVal lSize As Long) As Long
Public Const AVIFILECAPS_NOCOMPRESSION = &H20
Public Const OF_SHARE_DENY_WRITE = &H20

Sub PlayAVIPictureBox(filename As String, ByVal Window As PictureBox)
Dim RetVal As Long
Dim CommandString As String
Dim ShortFileName As String * 260
Dim deviceIsOpen As Boolean
RetVal = GetShortPathName(filename, ShortFileName, Len(ShortFileName))
filename = Left$(ShortFileName, RetVal)
CommandString = "Open " & filename & " type AVIVideo alias AVIFile parent " _
& CStr(Window.hwnd) & " style " & CStr(WS_CHILD)
RetVal = mciSendString(CommandString, vbNullString, 0, 0&)
If RetVal Then GoTo error
deviceIsOpen = True
CommandString = "put AVIFile window at 0 0 " & CStr(Window.ScaleWidth / _
Screen.TwipsPerPixelX) & " " & CStr(Window.ScaleHeight / _
Screen.TwipsPerPixelY)
RetVal = mciSendString(CommandString, vbNullString, 0, 0&)
If RetVal <> 0 Then GoTo error
CommandString = "Play AVIFile wait"
RetVal = mciSendString(CommandString, vbNullString, 0, 0&)
If RetVal <> 0 Then GoTo error
CommandString = "Close AVIFile"
RetVal = mciSendString(CommandString, vbNullString, 0, 0&)
If RetVal <> 0 Then GoTo error
Exit Sub
error:
Dim ErrorString As String
ErrorString = Space$(256)
mciGetErrorString RetVal, ErrorString, Len(ErrorString)
ErrorString = Left$(ErrorString, InStr(ErrorString, vbNullChar) - 1)
If deviceIsOpen Then
CommandString = "Close AVIFile"
mciSendString CommandString, vbNullString, 0, 0&
End If
MsgBox ErrorString, vbCritical, "Error"
End Sub

Public Function LoadIntoMemory(ByVal NameX As Long, ByVal TypeX As Long, data() As Byte) As Boolean
Dim HOBJ As Long
Dim TTB As Long
Dim TTF As Long
Dim SZ As Long
HOBJ = FindResourceEx(HMOD, TypeX, NameX, LangID)
'HOBJ = FindResource(HMOD, nameX, typeX)
TTB = LoadResource(HMOD, HOBJ)
TTF = LockResource(TTB)
SZ = SizeofResource(HMOD, HOBJ)
If SZ = 0 Or HOBJ = 0 Then LoadIntoMemory = True: Exit Function
ReDim data(SZ - 1)
CopyMemory data(0), ByVal TTF, SZ
ResTotLen = SZ
End Function
Public Function LoadAndFiXBitmapHeader(ByVal NameX As Long, ByVal TypeX As Long) As Boolean
Dim HOBJ1 As Long
Dim TTB1 As Long
Dim TTF1 As Long
'HOBJ1 = FindResource(HMOD, TrueName, TypePtr)
HOBJ1 = FindResourceEx(HMOD, TypeX, NameX, LangID)
TTB1 = LoadResource(HMOD, HOBJ1)
TTF1 = LockResource(TTB1)
ResTotLen = SizeofResource(HMOD, HOBJ1)
If ResTotLen = 0 Or HOBJ1 = 0 Then LoadAndFiXBitmapHeader = True: Exit Function
ReDim OtherData(ResTotLen - 1 + 14)
'S obzirom da u resource bitmape-e ne postoji BITMAP FILE HEADER dodat cemo ga!
OtherData(0) = &H42
OtherData(1) = &H4D
'Upisi ukupnu dužinu BITMAPE FILE zajedno sa dužinom headera(14)
CopyMemory OtherData(2), ResTotLen + 14, 4
CopyMemory OtherData(14), ByVal TTF1, ResTotLen
'Izracunaj OFFSET ili gdje se nalaze RAW podaci u file-u!!!!
Dim OFFSET As Long

Dim CLRUSED As Long
CopyMemory CLRUSED, OtherData(50), 4

Dim COMPRSED As Long
CopyMemory COMPRSED, OtherData(30), 4


OFFSET = 56 - 2
Select Case OtherData(28)

Case 1
OFFSET = OFFSET + 8

Case 4
If Not CBool(CLRUSED) Then
OFFSET = OFFSET + 64
Else
OFFSET = OFFSET + CLRUSED * 4
End If

Case 8
If Not CBool(CLRUSED) Then
OFFSET = OFFSET + 1024
Else
OFFSET = OFFSET + CLRUSED * 4
End If

End Select
CopyMemory OtherData(10), OFFSET, 4
'Simulirali smo HEADER--->>>>BRAVO..
End Function
Public Sub SetWinPosByCursor(ByVal hwnd As Long, ByVal stat As Long)
Dim ppt As POINTAPI
GetCursorPos ppt
SetWindowPos hwnd, 0, ppt.x, ppt.y, 0, 0, stat
End Sub


Public Function dialogProc(ByVal hwnd As Long, ByVal umsg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long


End Function
