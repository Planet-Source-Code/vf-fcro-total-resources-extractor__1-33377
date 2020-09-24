Attribute VB_Name = "Module3"
'Window CLASS
Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2
Public Const CS_CLASSDC = &H40
Public Const CS_PARENTDC = &H80
Public Const CS_OWNDC = &H20

Public Const DLGWINDOWEXTRA = 30 'Bez ovoga ne radi CUSTOM klasa Dialoga!!!!!!!!


Public Const IDI_APPLICATION = 32512&
Public Const IDC_ARROW = 32512&

Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DefDlgProc Lib "user32" Alias "DefDlgProcA" (ByVal hdlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long


Public CLASSunk As New Collection


'Dialog STYLE:
Public Const DS_ABSALIGN = &H1&
Public Const DS_SYSMODAL = &H2&
Public Const DS_3DLOOK = &H4&
Public Const DS_FIXEDSYS = &H8&
Public Const DS_NOFAILCREATE = &H10&
Public Const DS_LOCALEDIT = &H20
Public Const DS_SETFONT = &H40
Public Const DS_MODALFRAME = &H80
Public Const DS_NOIDLEMSG = &H100
Public Const DS_SETFOREGROUND = &H200
Public Const DS_CONTROL = &H400&
Public Const DS_CENTER = &H800&
Public Const DS_CONTEXTHELP = &H2000&
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_THICKFRAME = &H40000
Public Const WS_SYSMENU = &H80000
Public Const WS_HSCROLL = &H100000
Public Const WS_VSCROLL = &H200000
Public Const WS_DLGFRAME = &H400000
Public Const WS_BORDER = &H800000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_CHILD = &H40000000
Public Const WS_POPUP = &H80000000

'Dialog EXTENDED STYLE
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_EX_MDICHILD = &H40&
Public Const WS_EX_TOOLWINDOW = &H80&
Public Const WS_EX_WINDOWEDGE = &H100&
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_CONTEXTHELP = &H400&
Public Const WS_EX_RIGHT = &H1000&
Public Const WS_EX_RTLREADING = &H2000&
Public Const WS_EX_LEFTSCROLLBAR = &H4000&
Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_EX_APPWINDOW = &H40000

Public Type Dlgtemplate
    style As Long
    ExStyle As Long
    cdit As Integer 'Broj itema
    x As Integer
    y As Integer
    cx As Integer
    cy As Integer
End Type

Public Type DlgItemTemplate
    style As Long
    ExStyle As Long
    x As Integer
    y As Integer
    cx As Integer
    cy As Integer
    id As Integer
End Type



'**********************************************************
Public Type DlgtemplateEx
    HELPID As Long
    ExStyle As Long
    style As Long
    cdit As Integer 'Broj itema
    x As Integer
    y As Integer
    cx As Integer
    cy As Integer
End Type

Public Type DlgItemTemplateEX
    HELPID As Long
    ExStyle As Long
    style As Long
    x As Integer
    y As Integer
    cx As Integer
    cy As Integer
    id As Integer
End Type


'Nakon toga dolazi Menu,Class,Caption,i nakon toga FONT..
'Ukoliko je DS_SETFONT setiran!!!!!!
'Pointsize As Integer
'Weight As Integer
'Italic as Byte
'CharSet As Byte
'Fontname Unicode
'**********************************************************

'Važno...Kada upotrebimo DlgItemTemplateEX strukturu za dobivanje podataka o kontrolama
'tada na kraju dolazi EXTRACOUNT za podatke koji se šalju u lparam,ali na DWORD Boundary-u..
'tj, ukoliko struktura ne završava djeljiva sa 4 tada dodajemo još 2 i tada dolazi struktura EXTRACOUNT
'Ukoliko je 00 00 00 00 tada nema extra podataka...ukoliko ima broj koji dobivamo je
'dužina u byte-ovima strukture koja dolazi...
'Upozorenje 2:na kraju strukture dodati 0 toliko da bude struktura djeljiva sa 4-->

Public Type DLGCONTROLSEX
DlgtemplateEx As DlgItemTemplateEX
classnameX As String
captionX As String
End Type

Public Type DLGCONTROLS
DlgTmplate As DlgItemTemplate
classnameX As String
captionX As String
End Type


Public Type FNTSTR
 pointsize As Integer  'Velicina fonta
 weight As Integer   'Dali je bold,extrabold,normal ....
 italic As Byte
 charset As Byte
End Type


'***********DIALOG MODULE
Public Type ADDITIONALINF
 MENU As String
 Class As String
 caption As String
 FONTSIZE As Integer
 FONTNAME As String
End Type

Public FONT1 As FNTSTR
Public DIALOGSTYLE As String
Public DIALOGEXSTYLE As String
Public DIALOGINF1 As ADDITIONALINF
Public DIALOG1 As Dlgtemplate
Public DIALOG2 As DlgtemplateEx
Public CONTROLS1() As DLGCONTROLS
Public CONTROLS2() As DLGCONTROLSEX
Dim countX As Long
Dim CHECKINF As Integer
Dim UNILEN As Long



Public Sub ShowDLGInfo()

OINF = "Dialog Resource ID:" & GRPinfo & vbCrLf
OINF = OINF & "Length:" & ResTotLen & " (" & Hex(ResTotLen) & "h) Bytes" & vbCrLf
OINF = OINF & "Class Name:" & DIALOGINF1.Class & vbCrLf
OINF = OINF & "Caption:" & DIALOGINF1.caption & vbCrLf
OINF = OINF & "X Position:" & DIALOG1.x & vbCrLf
OINF = OINF & "Y Position:" & DIALOG1.y & vbCrLf
OINF = OINF & "Width:" & DIALOG1.cx & vbCrLf
OINF = OINF & "Height:" & DIALOG1.cy & vbCrLf
OINF = OINF & "Number of Controls:" & DIALOG1.cdit & vbCrLf
OINF = OINF & "Style:" & DIALOGSTYLE & vbCrLf
OINF = OINF & "Extended Style:" & DIALOGEXSTYLE & vbCrLf
OINF = OINF & "Font:" & DIALOGINF1.FONTNAME & ",Font Size:" & DIALOGINF1.FONTSIZE & vbCrLf
OINF = OINF & "Menu:" & DIALOGINF1.MENU & vbCrLf & vbCrLf

For u = 0 To DIALOG1.cdit - 1

OINF = OINF & "Control " & (u + 1) & vbCrLf
OINF = OINF & "Class Name:" & CONTROLS1(u).classnameX & vbCrLf
OINF = OINF & "ID:" & CONTROLS1(u).DlgTmplate.id & vbCrLf
OINF = OINF & "Caption:" & CONTROLS1(u).captionX & vbCrLf
OINF = OINF & "X Position:" & CONTROLS1(u).DlgTmplate.x & vbCrLf
OINF = OINF & "Y Position:" & CONTROLS1(u).DlgTmplate.y & vbCrLf
OINF = OINF & "Width:" & CONTROLS1(u).DlgTmplate.cx & vbCrLf
OINF = OINF & "Height:" & CONTROLS1(u).DlgTmplate.cy & vbCrLf
OINF = OINF & "Style:" & CONTROLS1(u).DlgTmplate.style & vbCrLf
OINF = OINF & "Extended Style:" & CONTROLS1(u).DlgTmplate.ExStyle & vbCrLf & vbCrLf

Next u
End Sub
Public Sub ShowDLGInfoEx()
With INFO
OINF = "Dialog Resource ID:" & GRPinfo & vbCrLf
OINF = OINF & "File Length:" & ResTotLen & " (" & Hex(ResTotLen) & "h) Bytes" & vbCrLf
OINF = OINF & "Class Name:" & DIALOGINF1.Class & vbCrLf
OINF = OINF & "Caption:" & DIALOGINF1.caption & vbCrLf
OINF = OINF & "X Position:" & DIALOG2.x & vbCrLf
OINF = OINF & "Y Position:" & DIALOG2.y & vbCrLf
OINF = OINF & "Width:" & DIALOG2.cx & vbCrLf
OINF = OINF & "Height:" & DIALOG2.cy & vbCrLf
OINF = OINF & "Number of Controls:" & DIALOG2.cdit & vbCrLf
OINF = OINF & "Style:" & DIALOGSTYLE & vbCrLf
OINF = OINF & "Extended Style:" & DIALOGEXSTYLE & vbCrLf
OINF = OINF & "Font:" & DIALOGINF1.FONTNAME & ",Font Size:" & FONT1.pointsize & ",Font Statement:" & FONT1.pointsize & ",Font Charset:" & FONT1.charset & ",Font Italic:" & FONT1.italic & vbCrLf
OINF = OINF & "Menu:" & DIALOGINF1.MENU & vbCrLf
OINF = OINF & "Help ID:" & DIALOG2.HELPID & vbCrLf & vbCrLf

For u = 0 To DIALOG2.cdit - 1
OINF = OINF & "Control " & (u + 1) & vbCrLf
OINF = OINF & "Class Name:" & CONTROLS2(u).classnameX & vbCrLf
OINF = OINF & "ID:" & CONTROLS2(u).DlgtemplateEx.id & vbCrLf
OINF = OINF & "Caption:" & CONTROLS2(u).captionX & vbCrLf
OINF = OINF & "X Position:" & CONTROLS2(u).DlgtemplateEx.x & vbCrLf
OINF = OINF & "Y Position:" & CONTROLS2(u).DlgtemplateEx.y & vbCrLf
OINF = OINF & "Width:" & CONTROLS2(u).DlgtemplateEx.cx & vbCrLf
OINF = OINF & "Height:" & CONTROLS2(u).DlgtemplateEx.cy & vbCrLf
OINF = OINF & "Style:" & CONTROLS2(u).DlgtemplateEx.style & vbCrLf
OINF = OINF & "Extended Style:" & CONTROLS2(u).DlgtemplateEx.ExStyle & vbCrLf
OINF = OINF & "Help ID:" & CONTROLS2(u).DlgtemplateEx.HELPID & vbCrLf & vbCrLf

Next u

End With
End Sub

Public Sub GetDialogInformation()
Set CLASSunk = Nothing
Dim CHECKSGN As Integer
CopyMemory CHECKSGN, OtherData(2), 2
'Provjera dali se radi o normalnom ili extended dialogu?!
If CHECKSGN = CInt(&HFFFF) Then
'Extended
DlgtemplateEx
If DIALOG2.cdit <> 0 Then
GetControlsInfoEx DIALOG2.cdit, countX
End If
ShowDLGInfoEx
Else
'Normal
Dlgtemplate
If DIALOG1.cdit <> 0 Then
GetControlsInfo DIALOG1.cdit, countX
End If
ShowDLGInfo
End If

End Sub
Public Sub GetControlsInfoEx(ByVal numb As Integer, ByVal count As Long)
On Error GoTo eRe
ReDim CONTROLS2(numb - 1)
For u = 1 To numb
CopyMemory CONTROLS2(u - 1).DlgtemplateEx, OtherData(count), Len(CONTROLS2(u - 1).DlgtemplateEx)
count = count + Len(CONTROLS2(u - 1).DlgtemplateEx)
'Uzmi CLASSU I UŠTIMAJ ALIGNMENT:

If (count Mod 4) <> 0 Then
count = count + 2
End If

CopyMemory CHECKINF, OtherData(count), 2
If CHECKINF = CInt(&HFFFF) Then
CopyMemory CHECKINF, OtherData(count + 2), 2
CONTROLS2(u - 1).classnameX = SetDLGCTRLName(CHECKINF)
count = count + 4

Else
UNILEN = lstrlenW(ByVal VarPtr(OtherData(count)))
CONTROLS2(u - 1).classnameX = Space(UNILEN)
CopyMemory ByVal StrPtr(CONTROLS2(u - 1).classnameX), OtherData(count), UNILEN * 2
count = count + UNILEN * 2 + 2
End If

'Uzmi CAPTION i UŠTIMAJ ALIGNMENT:
CopyMemory CHECKINF, OtherData(count), 2
If CHECKINF = CInt(&HFFFF) Then
CopyMemory CHECKINF, OtherData(count + 2), 2
CONTROLS2(u - 1).captionX = CStr(CHECKINF)
count = count + 4
If (count Mod 4) <> 0 Then count = count + 2
'GoTo dalje:
Else
UNILEN = lstrlenW(ByVal VarPtr(OtherData(count)))
CONTROLS2(u - 1).captionX = Space(UNILEN)
CopyMemory ByVal StrPtr(CONTROLS2(u - 1).captionX), OtherData(count), UNILEN * 2
count = count + UNILEN * 2


If (count Mod 4) <> 0 Then count = count + 2
End If

Dim CHECKINF2 As Long
CopyMemory CHECKINF2, OtherData(count), 4
If Not CBool(CHECKINF2) Then
count = count + 4
Else
count = count + CHECKINF2
Do While (count Mod 4) <> 0
count = count + 1 'DWORD ALIGNMENT
Loop
End If

dalje:
CHECKunknowClass CONTROLS2(u - 1).classnameX
Next u
Exit Sub

eRe:
On Error GoTo 0
ERRORX = True
End Sub
Public Sub GetControlsInfo(ByVal numb As Integer, ByVal count As Long)
On Error GoTo eRe
ReDim CONTROLS1(numb - 1)
For u = 1 To numb
CopyMemory CONTROLS1(u - 1).DlgTmplate, OtherData(count), Len(CONTROLS1(u - 1).DlgTmplate)
count = count + Len(CONTROLS1(u - 1).DlgTmplate)
'Uzmi CLASSU I UŠTIMAJ ALIGNMENT:
CopyMemory CHECKINF, OtherData(count), 2
If CHECKINF = CInt(&HFFFF) Then
CopyMemory CHECKINF, OtherData(count + 2), 2
CONTROLS1(u - 1).classnameX = SetDLGCTRLName(CHECKINF)
count = count + 4

Else
UNILEN = lstrlenW(ByVal VarPtr(OtherData(count)))
CONTROLS1(u - 1).classnameX = Space(UNILEN)
CopyMemory ByVal StrPtr(CONTROLS1(u - 1).classnameX), OtherData(count), UNILEN * 2
count = count + UNILEN * 2 + 2
End If

'Uzmi CAPTION i UŠTIMAJ ALIGNMENT:
CopyMemory CHECKINF, OtherData(count), 2
If CHECKINF = CInt(&HFFFF) Then
CopyMemory CHECKINF, OtherData(count + 2), 2
CONTROLS1(u - 1).captionX = CStr(CHECKINF)
count = count + 4
If (count Mod 4) <> 0 Then count = count + 2
GoTo dalje:
Else
UNILEN = lstrlenW(ByVal VarPtr(OtherData(count)))
CONTROLS1(u - 1).captionX = Space(UNILEN)
CopyMemory ByVal StrPtr(CONTROLS1(u - 1).captionX), OtherData(count), UNILEN * 2
count = count + UNILEN * 2

'If UNILEN = 0 Then count = count + 2
If (count Mod 4) <> 0 Then count = count + 2
End If

Dim CHECKINF2 As Long
CopyMemory CHECKINF2, OtherData(count), 4
If Not CBool(CHECKINF2) Then
count = count + 4
Else
count = count + CHECKINF2
Do While (count Mod 4) <> 0
count = count + 1 'DWORD ALIGNMENT
Loop
End If
CHECKunknowClass CONTROLS1(u - 1).classnameX
dalje:
Next u
Exit Sub

eRe:
On Error GoTo 0
ERRORX = True
End Sub

Public Function SetDLGCTRLName(ByVal num As Integer) As String

Select Case num
Case &H80
SetDLGCTRLName = "BUTTON"
Case &H81
SetDLGCTRLName = "EDIT"
Case &H82
SetDLGCTRLName = "STATIC"
Case &H83
SetDLGCTRLName = "LISTBOX"
Case &H84
SetDLGCTRLName = "SCROLLBAR"
Case &H85
SetDLGCTRLName = "COMBOBOX"
Case Else
SetDLGCTRLName = CStr(num)
End Select

End Function

Public Sub DlgtemplateEx()
CopyMemory DIALOG2, OtherData(4), Len(DIALOG2)
CopyMemory CHECKINF, OtherData(22), 2
DIALOGSTYLE = GetStyle(DIALOG2.style)
DIALOGEXSTYLE = GetStyle(DIALOG2.ExStyle)

countX = Len(DIALOG2) + 4
countX = CheckMenu(countX)
countX = CheckClass(countX)
countX = CheckCaption(countX)

If (DIALOG2.style And (DS_SETFONT Or DS_FIXEDSYS)) = (DS_SETFONT Or DS_FIXEDSYS) Or _
(DIALOG2.style And DS_SETFONT) = DS_SETFONT Then
CopyMemory FONT1, OtherData(countX), Len(FONT1)
countX = countX + Len(FONT1)
countX = CheckFontName(countX)
Else
FONT1.charset = 0
FONT1.italic = 0
FONT1.pointsize = 0
FONT1.weight = 0
End If
'Kad ispisujemo podatke o fontu za extended dialog ime fonta se nalazi u DialogInf1.FONTNAME
'Svi ostali podaci se nalaze u font1 type-u...

If (countX Mod 4) <> 0 Then countX = countX + 2
If DIALOGINF1.Class <> "Default Dialog Box" Then
CHECKunknowClass DIALOGINF1.Class, True, True
End If
End Sub



Public Sub Dlgtemplate()
CopyMemory DIALOG1, OtherData(0), Len(DIALOG1)
DIALOGSTYLE = GetStyle(DIALOG1.style)
DIALOGEXSTYLE = GetExStyle(DIALOG1.ExStyle)

countX = Len(DIALOG1)
countX = CheckMenu(countX)
countX = CheckClass(countX)
countX = CheckCaption(countX)


If (DIALOG1.style And DS_SETFONT) = DS_SETFONT Then
'FONT***************DLGTEMPLATE
CopyMemory DIALOGINF1.FONTSIZE, OtherData(countX), 2
countX = countX + 2
countX = CheckFontName(countX)
Else
DIALOGINF1.FONTNAME = ""
DIALOGINF1.FONTSIZE = 0
End If
If (countX Mod 4) <> 0 Then countX = countX + 2
If DIALOGINF1.Class <> "Default Dialog Box" Then
CHECKunknowClass DIALOGINF1.Class, True, True
End If
End Sub
Public Function CheckMenu(ByVal count As Long) As Long
'Ukoliko ima informacija uzmi Menu informacije...
'MENU*****************
CopyMemory CHECKINF, OtherData(count), 2
If Not CBool(CHECKINF) Then
DIALOGINF1.MENU = "No Menu"
count = count + 2: GoTo eend:
End If
If CHECKINF = CInt(&HFFFF) Then
CopyMemory CHECKINF, OtherData(count + 2), 2
DIALOGINF1.MENU = CStr(CHECKINF)
count = count + 4
Else
'Ukoliko se radi o imenu,a ne broju on dolazi kao unicode
UNILEN = lstrlenW(ByVal VarPtr(OtherData(count)))
DIALOGINF1.MENU = Space(UNILEN)
CopyMemory ByVal StrPtr(DIALOGINF1.MENU), OtherData(count), UNILEN * 2
count = count + UNILEN * 2 + 2
End If
eend:
CheckMenu = count
End Function
Public Function CheckClass(ByVal count As Long) As Long
'CLASS***************
'Provjeri informacije o Klasi / Potpuno isto kao kod MENU--->
CopyMemory CHECKINF, OtherData(count), 2
If Not CBool(CHECKINF) Then
DIALOGINF1.Class = "Default Dialog Box"
count = count + 2: GoTo eend:
End If
If CHECKINF = CInt(&HFFFF) Then
CopyMemory CHECKINF, OtherData(count + 2), 2
DIALOGINF1.Class = CStr(CHECKINF)
count = count + 4
Else
UNILEN = lstrlenW(ByVal VarPtr(OtherData(count)))
DIALOGINF1.Class = Space(UNILEN)
CopyMemory ByVal StrPtr(DIALOGINF1.Class), OtherData(count), UNILEN * 2
count = count + UNILEN * 2 + 2
End If
eend:
CheckClass = count
End Function

Public Function CheckCaption(ByVal count As Long) As Long
'CAPTION**************
'Provjeri CAPTION dialog box-a
CopyMemory CHECKINF, OtherData(count), 2
If Not CBool(CHECKINF) Then
DIALOGINF1.caption = ""
count = count + 2: GoTo eend:
End If
If Not (Not CBool(CHECKINF)) Then
'Ukoliko nije 0--ludi izraz
UNILEN = lstrlenW(ByVal VarPtr(OtherData(count)))
DIALOGINF1.caption = Space(UNILEN)
CopyMemory ByVal StrPtr(DIALOGINF1.caption), OtherData(count), UNILEN * 2
count = count + UNILEN * 2 + 2
End If
eend:
CheckCaption = count
End Function

Public Function CheckFontName(ByVal count As Long) As Long
UNILEN = lstrlenW(ByVal VarPtr(OtherData(count)))
DIALOGINF1.FONTNAME = Space(UNILEN)
CopyMemory ByVal StrPtr(DIALOGINF1.FONTNAME), OtherData(count), UNILEN * 2
count = count + UNILEN * 2 + 2
CheckFontName = count
End Function


Public Function GetStyle(ByVal value As Long) As String
Dim sty As String
If (value And &H1&) = &H1& Then sty = sty & "DS_ABSALIGN | "
If (value And &H2&) = &H2& Then sty = sty & "DS_SYSMODAL | "
If (value And &H4&) = &H4& Then sty = sty & "DS_3DLOOK | "
If (value And &H8&) = &H8& Then sty = sty & "DS_FIXEDSYS | "
If (value And &H10&) = &H10& Then sty = sty & "DS_NOFAILCREATE | "
If (value And &H20&) = &H20& Then sty = sty & "DS_LOCALEDIT | "
If (value And &H40&) = &H40& Then sty = sty & "DS_SETFONT | "
If (value And &H80&) = &H80& Then sty = sty & "DS_MODALFRAME | "
If (value And &H100&) = &H100& Then sty = sty & "DS_NOIDLEMSG | "
If (value And &H200&) = &H200& Then sty = sty & "DS_SETFOREGROUND | "
If (value And &H400&) = &H400& Then sty = sty & "DS_CONTROL | "
If (value And &H800&) = &H800& Then sty = sty & "DS_CENTER | "
If (value And &H1000&) = &H1000& Then sty = sty & "DS_CENTERMOUSE | "
If (value And &H2000&) = &H2000& Then sty = sty & "DS_CONTEXTHELP | "
If (value And &H10000) = &H10000 Then sty = sty & "WS_MAXIMIZEDBOX | "
If (value And &H20000) = &H20000 Then sty = sty & "WS_MINIMIZEDBOX | "
If (value And &H40000) = &H40000 Then sty = sty & "WS_THICKFRAME | "
If (value And &H80000) = &H80000 Then sty = sty & "WS_SYSMENU | "
If (value And &H100000) = &H100000 Then sty = sty & "WS_HSCROLL | "
If (value And &H200000) = &H200000 Then sty = sty & "WS_VSCROLL | "
If (value And &H400000) = &H400000 Then sty = sty & "WS_DLGFRAME | "
If (value And &H800000) = &H800000 Then sty = sty & "WS_BORDER | "
If (value And &H1000000) = &H1000000 Then sty = sty & "WS_MAXIMIZE | "
If (value And &H2000000) = &H2000000 Then sty = sty & "WS_CLIPCHILDREN | "
If (value And &H4000000) = &H4000000 Then sty = sty & "WS_CLIPSIBLINGS | "
If (value And &H8000000) = &H8000000 Then sty = sty & "WS_DISABLED | "
If (value And &H10000000) = &H10000000 Then sty = sty & "WS_VISIBLE | "
If (value And &H20000000) = &H20000000 Then sty = sty & "WS_MINIMIZE | "
If (value And &H40000000) = &H40000000 Then sty = sty & "WS_CHILD | "
If (value And &H80000000) = &H80000000 Then sty = sty & "WS_POPUP | "
If Not CBool(Len(sty)) Then Exit Function
sty = Left(sty, Len(sty) - 2)
GetStyle = sty
End Function

Public Function GetExStyle(ByVal value As Long) As String
Dim exsty As String
If (value And &H1&) = &H1& Then exsty = exsty & "WS_EX_DLGMODALFRAME | "
If (value And &H4&) = &H4& Then exsty = exsty & "WS_EX_NOPARENTNOTIFY | "
If (value And &H8&) = &H8& Then exsty = exsty & "WS_EX_TOPMOST | "
If (value And &H10&) = &H10& Then exsty = exsty & "WS_EX_ACCEPTFILES | "
If (value And &H20&) = &H20& Then exsty = exsty & "WS_EX_TRANSPARENT | "
If (value And &H40&) = &H40& Then exsty = exsty & "WS_EX_MDICHILD | "
If (value And &H80&) = &H80& Then exsty = exsty & "WS_EX_TOOLWINDOW | "
If (value And &H100&) = &H100& Then exsty = exsty & "WS_EX_WINDOWEDGE | "
If (value And &H200&) = &H200& Then exsty = exsty & "WS_EX_CLIENTEDGE | "
If (value And &H400&) = &H400& Then exsty = exsty & "WS_EX_CONTEXTHELP | "
If (value And &H1000&) = &H1000& Then exsty = exsty & "WS_EX_RIGHT | "
If (value And &H2000&) = &H2000& Then exsty = exsty & "WS_EX_RTLREADING | "
If (value And &H4000&) = &H4000& Then exsty = exsty & "WS_EX_LEFTSCROLLBAR | "
If (value And &H10000) = &H10000 Then exsty = exsty & "WS_EX_CONTROLPARENT | "
If (value And &H20000) = &H20000 Then exsty = exsty & "WS_EX_STATICEDGE | "
If (value And &H40000) = &H40000 Then exsty = exsty & "WS_EX_APPWINDOW | "
If Not CBool(Len(exsty)) Then Exit Function
exsty = Left(exsty, Len(exsty) - 2)
GetExStyle = exsty
End Function


Public Sub CHECKunknowClass(ByVal name As String, Optional skipCheck As Boolean, Optional dlgORctrl As Boolean)
If IsNumeric(name) Then Exit Sub

Dim tempCLASSX As WNDCLASS
'Provjeri dali klasa postoji
If skipCheck Then GoTo rreg
If Not CBool(GetClassInfo(HMOD, name, tempCLASSX)) Then
rreg:
REGunknowClass name, dlgORctrl
End If
End Sub

Public Sub REGunknowClass(ByVal name As String, ByVal dlgORctrl As Boolean)
'Registriraj Custom Class da bi mogli prikazati dialog prozor,
'ukoliko to ne nepravimo pada inicijalizacija Dialoga i ništa od prikaza!!!!
Dim tempCLASSX As WNDCLASS

If Not dlgORctrl Then
'Ukoliko se radi o kontroli unutar Dialoga postavi ovo:
tempCLASSX.style = CS_HREDRAW Or CS_VREDRAW
tempCLASSX.lpfnwndproc = GetAddress(AddressOf WndProc)
Else
'Ukoliko se radi o samom dialogu tada postavi druge vrijednosti ili ce pasti inicijalizacija
tempCLASSX.style = 2056
tempCLASSX.cbWndExtra2 = DLGWINDOWEXTRA
tempCLASSX.lpfnwndproc = GetAddress(AddressOf DlgProc)
End If

tempCLASSX.hInstance = HMOD
tempCLASSX.hIcon = LoadIcon(HMOD, IDI_APPLICATION)
tempCLASSX.hCursor = LoadCursor(HMOD, IDC_ARROW)
tempCLASSX.hbrBackground = &H11
tempCLASSX.lpszClassName = name
Call RegisterClass(tempCLASSX)

CLASSunk.Add name 'Dodaj u listu
End Sub


Public Function UNREGunknowClass()
For u = 1 To CLASSunk.count
UnregisterClass CStr(CLASSunk.Item(u)), HMOD
Next u
Set CLASSunk = Nothing
End Function


Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Ovo je DEFAULTNA Window procedura za CLASSE koje smo pronašli a nisu registrirane

WndProc = DefWindowProc(hWnd&, uMsg&, wParam&, lParam&)
End Function

Public Function DlgProc(ByVal hdlg As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Ovo je DEFAULTNA Dialog procedura za CLASSE (Dialoga) koje smo pronašli a nisu registrirane

DlgProc = DefDlgProc(hdlg&, uMsg&, wParam&, lParam&)
End Function

