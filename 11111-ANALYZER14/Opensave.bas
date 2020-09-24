Attribute VB_Name = "OpenSave"
Option Explicit
Private Const HH_DISPLAY_TEXT_POPUP = &HE
Private Declare Function HtmlHelp Lib "HHCtrl.ocx" _
        Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, _
        ByVal uCommand As Long, dwData As Any) As Long

Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function AppendMenu Lib "user32" _
        Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, _
        ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, _
        ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, _
        ByVal nReserved As Long, ByVal hWnd As Long, lprc As RECT) As Long

Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

Public menuhwnd As Long
Private CRSR As POINTAPI
Private Const TPM_BOTTOMALIGN = &H20&
Private Const TPM_CENTERALIGN = &H4&
Private Const TPM_HORPOSANIMATION = &H400&
Private Const TPM_HORNEGANIMATION = &H800&
Private Const TPM_NOANIMATION = &H4000&
Private Const TPM_RIGHTALIGN = &H8&
Private Const TPM_RETURNCMD = &H100&
Private Const TPM_RIGHTBUTTON = &H2&
Private Const TPM_VCENTERALIGN = &H10&
Private Const TPM_VERTICAL = &H40&
Private Const TPM_VERPOSANIMATION = &H1000&
Private Const TPM_VERNEGANIMATION = &H2000&
Private Const TPM_LEFTALIGN = &H0&


Const BS_TOP = &H400&
Const BS_FLAT = &H8000&

Public xhwnd As Long
Public PXhwnd As Long
Public Xpos As Long
Public Ypos As Long
Private prevID As Long
Private WWdt As Long

Private OldProcedura As Long
Private CLL As Long
Private DIALOGHWND As Long

Const GWL_EXSTYLE = (-20)
Const GWL_ID = (-12)
Const GWL_STYLE = (-16)

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type POINTAPI
     x As Long
     y As Long
End Type

Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hwndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    style As Long
    lpszname As String
    lpszClass As String
    ExStyle As Long
End Type

Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    itemData As Long
End Type


Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXBORDER = 5
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYCAPTION = 4
Public Const SM_CYDLGFRAME = 8
Public Const SM_CYBORDER = 6
Public Const SM_CYMENU = 15
Public Const SM_CYMENUSIZE = 55
Public Const SM_CXMENUSIZE = 54



Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetDlgCtrlID Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDlgItem Lib "user32" (ByVal hdlg As Long, ByVal nIDDlgItem As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetTimer& Lib "user32" (ByVal hWnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc&)
Private Declare Function KillTimer& Lib "user32" (ByVal hWnd&, ByVal nIDEvent&)

Private TMR As Long
Private TMRHWND As Long
Private ParOld As Long


Const WM_COMMAND = &H111
Const WM_CONTEXTMENU = &H7B
Const WM_NOTIFY = &H4E
Const WM_PAINT = &HF
Const WM_DRAWITEM = &H2B
Const WM_SETTEXT = &HC
Const WM_SETREDRAW = &HB
Public Const WM_DESTROY = &H2
Const WM_CLOSE = &H10
' ============================================================================
' GetOpen/SaveFileName
Const WM_INITDIALOG = &H110

Public Const WM_CREATE = &H1
Const WM_PARENTNOTIFY = &H210
Const WM_NCCREATE = &H81

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10




Public Const MAX_PATH = 260

Public Type OPENFILENAME  '  ofn
  lStructSize As Long
  hWndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As OFN_Flags
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

' File Open/Save Dialog Flags
Public Enum OFN_Flags
  OFN_READONLY = &H1
  OFN_OVERWRITEPROMPT = &H2
  OFN_HIDEREADONLY = &H4
  OFN_NOCHANGEDIR = &H8
  OFN_SHOWHELP = &H10
  OFN_ENABLEHOOK = &H20
  OFN_ENABLETEMPLATE = &H40
  OFN_ENABLETEMPLATEHANDLE = &H80
  OFN_NOVALIDATE = &H100
  OFN_ALLOWMULTISELECT = &H200
  OFN_EXTENSIONDIFFERENT = &H400
  OFN_PATHMUSTEXIST = &H800
  OFN_FILEMUSTEXIST = &H1000
  OFN_CREATEPROMPT = &H2000
  OFN_SHAREAWARE = &H4000
  OFN_NOREADONLYRETURN = &H8000&
  OFN_NOTESTFILECREATE = &H10000
  OFN_NONETWORKBUTTON = &H20000
  OFN_NOLONGNAMES = &H40000               ' force no long names for 4.x modules
  OFN_EXPLORER = &H80000                       ' new look commdlg
  OFN_NODEREFERENCELINKS = &H100000
  OFN_LONGNAMES = &H200000                 ' force long names for 3.x modules
  ' ===============================
  ' Win98/NT5 only...
  OFN_ENABLEINCLUDENOTIFY = &H400000           ' send include message to callback
  OFN_ENABLESIZING = &H800000
  ' ===============================
End Enum

Private Type HH_POPUP
        cbStruct As Long
        hinst As Long
        idString As Long
        pszText As Long ' pointer na string
        pt As POINTAPI
        clrForeground As Long
        clrBackground As Long
        rcMargins As RECT
        pzsFont As Long  ' pointer na string
End Type

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" _
(ByVal hdlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As _
Long

Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
'Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Property Let SetWidth(ByVal newwidth As Long)
WWdt = newwidth
End Property

Public Sub InsertCtrl(ByVal Ctrlhwnd As Long, ByVal parhwnd As Long, ByVal x As Long, ByVal y As Long)
xhwnd = Ctrlhwnd
PXhwnd = parhwnd
Xpos = x + 4
Ypos = y + 23
End Sub

Public Function GetOpenFilePath(hWnd As Long, _
                                                      sFilter As String, _
                                                      iFilter As Integer, _
                                                      sFile As String, _
                                                      sInitDir As String, _
                                                      sTitle As String, _
                                                      sRtnPath As String) As Boolean
  Dim ofn As OPENFILENAME
  
  With ofn
    .lStructSize = Len(ofn)
    .hWndOwner = hWnd
    .lpstrFilter = sFilter & vbNullChar & vbNullChar
    .nFilterIndex = iFilter
    .lpstrFile = sFile & String$(MAX_PATH - Len(sFile), 0)
    .nMaxFile = MAX_PATH
    .lpstrInitialDir = sInitDir
    .lpstrTitle = sTitle & vbNullChar
    .Flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST Or OFN_EXPLORER Or OFN_ENABLEHOOK

    .lpfnHook = GetAddress(AddressOf HookX)

  End With
  

  
  If GetOpenFileName(ofn) Then
    iFilter = ofn.nFilterIndex
    sFile = Mid$(ofn.lpstrFile, ofn.nFileOffset + 1, InStr(ofn.lpstrFile, vbNullChar) - (ofn.nFileOffset + 1))
    sRtnPath = GetStrFromBufferA(ofn.lpstrFile)
    GetOpenFilePath = True
  End If

End Function

Public Function GetSaveFilePath(hWnd As Long, _
                                                      sFilter As String, _
                                                      iFilter As Integer, _
                                                      sDefExt As String, _
                                                      sFile As String, _
                                                      sInitDir As String, _
                                                      sTitle As String, _
                                                      sRtnPath As String) As Boolean
  Dim ofn As OPENFILENAME
  With ofn
    .lStructSize = Len(ofn)
    .hWndOwner = hWnd
    .lpstrFilter = sFilter & vbNullChar & vbNullChar
    .lpstrFile = sFile & String$(MAX_PATH - Len(sFile), 0)
    .lpstrDefExt = sDefExt
    .nMaxFile = MAX_PATH
    .lpstrInitialDir = sInitDir
    .lpstrTitle = sTitle & vbNullChar
    .Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_EXPLORER Or OFN_ENABLEHOOK
    .lpfnHook = GetAddress(AddressOf HookX)
    
  End With
  
  If GetSaveFileName(ofn) Then
    iFilter = ofn.nFilterIndex
    sFile = Mid$(ofn.lpstrFile, ofn.nFileOffset + 1, InStr(ofn.lpstrFile, vbNullChar) - (ofn.nFileOffset + 1))
    sRtnPath = GetStrFromBufferA(ofn.lpstrFile)
    GetSaveFilePath = True
  End If

End Function

' Returns the string before first null char (if any) in an ANSII string.

Public Function GetStrFromBufferA(szA As String) As String
  If InStr(szA, vbNullChar) Then
    GetStrFromBufferA = Left$(szA, InStr(szA, vbNullChar) - 1)
  Else

    GetStrFromBufferA = szA
  End If
End Function
Public Function GetAddress(ByVal address As Long) As Long
GetAddress = address
End Function

Public Function HookX(ByVal hdlg As Long, ByVal uiMsg As _
Long, ByVal wparam As Long, ByVal lparam As Long) As _
Long

Select Case uiMsg

Case WM_INITDIALOG
Dim przRECT As RECT
Dim ctrlX As Long
Dim parhwnd As Long

parhwnd = GetParent(hdlg) 'Uzmi Handle Dialoga
DIALOGHWND = parhwnd
ctrlX = GetDlgItem(parhwnd, &H1)  'Uzmi handle ID Item=1

'*****Poziv Timera
'TMRHWND = ctrlX
'TMR = SetTimer(TMRHWND, 2, 0, AddressOf Provjera)
'******************
CLL = ctrlX
OldProcedura = SetWindowLong(ctrlX, -4, AddressOf Provjera2)
SetWindowText CLL, "Get It!"

SetParent xhwnd, parhwnd 'Promjeni vlasnika
prevID = GetDlgCtrlID(xhwnd) 'Zapamti Stari ID
SetWindowLong xhwnd, GWL_ID, &H6000 'Promjeni ID PictureBoxa da ga se ne dupla ID(proizvoljna vrijednost)

'Dim tmphwnd As Long
'tmphwnd = GetDlgItem(Parhwnd, 1)
'SetWindowLong tmphwnd, GWL_ID, &H6001
'SetWindowText tmphwnd, "Otvori"
'Pokušaj promjene ID ---prozor ne reagira

'*************************
Dim XY As POINTAPI
Dim RC1 As RECT
GetWindowRect parhwnd, RC1
XY.x = RC1.Left
XY.y = RC1.Top
ScreenToClient parhwnd, XY
GetWindowRect xhwnd, RC1
MoveWindow xhwnd, XY.x + Xpos, XY.y + Ypos, RC1.Right - RC1.Left, RC1.Bottom - RC1.Top, 1
ShowWindow xhwnd, 1
'********Postavi Picture Box

'Dim ID As Long
'Dim tmphwnd As Long
'tmphwnd = GetDlgItem(Parhwnd, 1)
'ID = GetDlgCtrlID(Xhwnd)
'Ovim gore je utvrdeno da je PictureBox došao na DLGItem=1


SetDlgItemText parhwnd, &H2, "Forget It!"
SetDlgItemText parhwnd, &H443, "There:"
SetDlgItemText parhwnd, &H442, "Name It:"
SetDlgItemText parhwnd, &H441, "That Type:"


GetWindowRect parhwnd, RC1
MoveWindow parhwnd, 0, 0, RC1.Right - RC1.Left, WWdt, 1 'Promjeni velicinu prozora



'******************************
GetWindowRect parhwnd, przRECT
Dim x As Long
Dim y As Long
x = (Screen.width / 15 - (przRECT.Right - przRECT.Left)) / 2
y = (Screen.height / 15 - (przRECT.Bottom - przRECT.Top)) / 2
SetWindowPos parhwnd, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
'*******Centriraj prozor

Dim strX As String
Dim ANSIX() As Byte
menuhwnd = CreatePopupMenu()
strX = "Što je ovo dovraga?" & Chr(CByte(0))
ANSIX = StrConv(strX, vbFromUnicode)
Call AppendMenu(menuhwnd, 0&, 1&, VarPtr(ANSIX(0)))
Call DrawMenuBar(menuhwnd)


ParOld = SetWindowLong(parhwnd, -4, AddressOf ParentSubclass)


Case WM_DESTROY
ShowWindow xhwnd, 0 'Sakrij kontrolu
'KillTimer TMRHWND, TMR 'Zgazi timer
SetParent xhwnd, parhwnd 'Vrati kontrolu na pravog vlasnika
SetWindowLong xhwnd, GWL_ID, prevID 'Vrati stari ID
Call SetWindowLong(CLL, -4, OldProcedura)
Call DestroyMenu(menuhwnd)
SetWindowLong parhwnd, -4, ParOld


End Select
End Function
Public Function Provjera2(ByVal hWnd As Long, ByVal umsg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long


Select Case umsg
Case WM_SETTEXT
Dim TTX As String
Dim TX() As Byte
TTX = "Get It!" & Chr(CByte(0))
TX = StrConv(TTX, vbFromUnicode)
'CopyMemory ByVal lparam, TX(0), Len(TTX)
lparam = VarPtr(TX(0))

'******Može i ovako......................
'Dim TX() As Byte 'Ansi
'Dim TTX As String 'Unicode
'TTX = "Snimi" & Chr(CByte(0))
'ReDim TX(Len(TTX-1))
'CopyMemory TX(0), ByVal TTX, Len(TTX)
'CopyMemory ByVal lparam, TX(0), Len(TTX)
'***************************************
'********Po tvome sam primjeru skužio dosta stvar...A to je da na lparam dolazi pointer
'na ANSI string... i umjesto da šaljem SetWindowText jednostavno sam kopirao strukturu,ne UNICODE veæ ANSI...

'********Nema razloga da se koristi VBACCELERATOROV Subclassing...(Dobar je za kontrole,da se ne ruši VB)
End Select
Provjera2 = CallWindowProc(OldProcedura, hWnd, umsg, wparam, lparam)
End Function


Public Function ParentSubclass(ByVal hWnd As Long, ByVal umsg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long

Select Case umsg

Case WM_CONTEXTMENU

Call GetCursorPos(CRSR)
Dim RTC As RECT
Call TrackPopupMenu(menuhwnd, TPM_LEFTALIGN, CRSR.x, CRSR.y, 0&, hWnd, RTC)
Exit Function


'Case WM_COMMAND
'If wParam = 1 Then

'Call GetCursorPos(CRSR)

'Dim PTHELP As POINTAPI
'Dim RECTHELP As RECT

'With PTHELP
'.x = CRSR.x
'.y = CRSR.y
'End With

'With RECTHELP
'.Left = 10
'.Top = 10
'.Bottom = 10
'.Right = 10
'End With
'Dim inftext As String
'Dim infmem() As Byte
'inftext = "Nemam pojma što je to!" & Chr(CByte(0))
'infmem = StrConv(inftext, vbFromUnicode)
'Dim HHTEXT As HH_POPUP

'With HHTEXT
'.cbStruct = Len(HHTEXT)
'.hinst = 0
'.idString = 0
'.pszText = VarPtr(infmem(0))
'.clrBackground = -2
'.clrForeground = -1
'.rcMargins = RECTHELP
'.pt = PTHELP
'End With

'Call HtmlHelp(0&, 0&, HH_DISPLAY_TEXT_POPUP, HHTEXT)

'End If

End Select

ParentSubclass = CallWindowProc(ParOld, hWnd, umsg, wparam, lparam)
End Function





'Public Sub Provjera(ByVal hwnd&, ByVal umsg&, ByVal idEvent&, ByVal dwTime&)
'Konstantno mijenjaj Text na buttonu ako je promjenjen.
'Ovo je izbaèeno...
'Dim TXT1 As String
'Dim ltxt1 As Long
'TXT1 = Space(20)
'ltxt1 = GetWindowText(hwnd, TXT1, Len(TXT1))
'TXT1 = Left(TXT1, ltxt1)
'If TXT1 = "&Open" Then
'SetWindowText hwnd, "Otvori Folder"
'ElseIf TXT1 = "&Save" Then
'SetWindowText hwnd, "Snimi"
'End If

'End Sub

Public Sub CloseDLG()
PostMessage DIALOGHWND, WM_CLOSE, 0, 0
End Sub




