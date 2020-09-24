VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resources-FINAL 1.01B by Vanja Fuckar"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Picture         =   "Form1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "View Export List"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   1080
      Picture         =   "Form1.frx":0A14
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Add Selected to Export List"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Picture         =   "Form1.frx":0B5E
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Clear Selection"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Picture         =   "Form1.frx":0CA8
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Save As DLL"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Picture         =   "Form1.frx":1041
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Open Resource File"
      Top             =   120
      Width           =   495
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6270
      Left            =   6000
      MousePointer    =   99  'Custom
      MultiSelect     =   1  'Simple
      TabIndex        =   9
      ToolTipText     =   "Double Click To View Resource"
      Top             =   1200
      Width           =   5895
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   1560
      ScaleHeight     =   975
      ScaleWidth      =   5775
      TabIndex        =   6
      Top             =   7920
      Visible         =   0   'False
      Width           =   5775
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         Height          =   960
         Left            =   0
         Picture         =   "Form1.frx":13DE
         ScaleHeight     =   900
         ScaleWidth      =   1125
         TabIndex        =   7
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Extractor by Vanja Fuckar,EMAIL:INGA@VIP.HR"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Exit Program"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Picture         =   "Form1.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Open Executeable"
      Top             =   120
      Width           =   495
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6270
      Left            =   0
      MousePointer    =   99  'Custom
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      ToolTipText     =   "Double Click To View Resource"
      Top             =   1200
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Working With:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Index           =   7
      Left            =   6000
      TabIndex        =   16
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Working With:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   15
      Top             =   720
      Width           =   5895
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00ACC2A5&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   6000
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Type ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   6000
      TabIndex        =   12
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Name ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   8520
      TabIndex        =   11
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lang ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   11040
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00AB6B6F&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F1D2C9&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lang ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Name ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CB8472&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Type ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim aa As Long
Dim sFile As String
Dim spath As String
Dim Xpoint As Long
Dim Ypoint As Long

Private Sub Command1_Click()
ClearSEL List1
End Sub


Private Sub CheckClick(listX As ListBox)
If lIndex = -1 Then Exit Sub
On Error GoTo eRe
Dim rees As Boolean
Dim icorcr As Boolean
'Obriši sve što je u memoriji
EraseFromMem
OINF = ""
ERRORX = False 'Pobriši Error da ne bi duplirali!

Dim rtNN As String

Dim SINGLEMEM() As Byte

If WMOD = 1 Then
CASSE = RESTYPE.item(lIndex + 1)
ElseIf WMOD = 2 Then
CASSE = RESTYPEFILE.item(lIndex + 1)
End If

CASSE2 = ""
Dim TypeNnm As String
TypeNnm = CStr(CASSE) & Chr(CByte(0))
If Not IsNumeric(CASSE) Then
TrueType = StrConv(TypeNnm, vbFromUnicode)
TypePtr = VarPtr(TrueType(0))
Else
TypePtr = CLng(CASSE)
End If
Dim nnm As String

If WMOD = 1 Then
nnm = RESNAME.item(lIndex + 1)
ElseIf WMOD = 2 Then
nnm = RESNAMEFILE.item(lIndex + 1)
End If

GRPinfo = nnm

If IsNumeric(nnm) Then
TrueName = CLng(nnm)
Else
nnm = nnm & Chr(CByte(0))
TrueBuffer = StrConv(nnm, vbFromUnicode)
TrueName = VarPtr(TrueBuffer(0))
End If

If WMOD = 1 Then
LangID = RESLANGID.item(lIndex + 1) 'Uzmi LangID
ElseIf WMOD = 2 Then
LangID = RESLANGIDFILE.item(lIndex + 1)
End If

If WMOD = 1 Then
rtNN = RESTYPENAME.item(lIndex + 1)
ElseIf WMOD = 2 Then
rtNN = RESTYPENAMEFILE.item(lIndex + 1)
End If

If CASSE = "2" Then
'BITMAP
If LoadAndFiXBitmapHeader(TrueName, TypePtr) Then GoTo erXv1
GoSub LoadPIC
If Not CBool(STD1(0)) Then GoTo eRe
If LoadIntoMemory(TrueName, TypePtr, TMPDATAX) Then GoTo erXv1
INFO.Show

ElseIf CASSE = "9" Then
'ACCELERATOR TABLE
If LoadIntoMemory(TrueName, TypePtr, OtherData) Then GoTo erXv1
GetAccelInfo OtherData
INFO.Show

ElseIf CASSE = "6" Then
'STRING TABLE
If LoadIntoMemory(TrueName, TypePtr, OtherData) Then GoTo erXv1
RTTEXT = SetStringNameEx(OtherData, TrueName)
GoSub CheckERX
INFO.Show

ElseIf CASSE = "11" Then
'MESSAGE TABLE
If LoadIntoMemory(TrueName, TypePtr, OtherData) Then GoTo erXv1
RTTEXT = GetMessageEntries(OtherData)
GoSub CheckERX
INFO.Show


ElseIf CASSE = "16" Then
'VERSION
If LoadIntoMemory(TrueName, TypePtr, OtherData) Then GoTo erXv1
GetFileInfo OtherData
INFO.Show

ElseIf CASSE = "14" Then
'IKONE-GRUPA
If LoadIntoMemory(TrueName, TypePtr, OtherData) Then GoTo erXv1
CopyMemory NEWH, OtherData(0), Len(NEWH)

ReDim RSDIR(NEWH.rescountX - 1)
For x = 1 To NEWH.rescountX
'Popuni Clanove sa informacijama
CopyMemory RSDIR(x - 1), OtherData(6 + (x - 1) * Len(RSDIR(0))), Len(RSDIR(0))
Next x
ReDim STD1(NEWH.rescountX - 1)
ReDim PicWidth(NEWH.rescountX - 1)
ReDim PicHeight(NEWH.rescountX - 1)
For x = 1 To NEWH.rescountX
'Simuliraj svaku pojedinacnu ikonu!!!!!!!
ReDim TMPDATAX(22 - 1 + RSDIR(x - 1).bitesinres)
CopyMemory TMPDATAX(0), CInt(0), 2
CopyMemory TMPDATAX(2), CInt(1), 2
CopyMemory TMPDATAX(4), CInt(1), 2
'popravi Height jer zna biti krivi...
'RSDIR(x - 1).iconresD.height = RSDIR(x - 1).iconresD.width
CopyMemory TMPDATAX(6), RSDIR(x - 1), Len(RSDIR(0))
CopyMemory TMPDATAX(18), CLng(&H16), 4
If LoadIntoMemory(CLng(RSDIR(x - 1).iconID), CLng(3), SINGLEMEM) Then GoTo erXv1

CopyMemory TMPDATAX(22), SINGLEMEM(0), UBound(SINGLEMEM) + 1

Set STD1(x - 1) = GetIconToPicture(TMPDATAX)
If Not CBool(STD1(x - 1)) Then GoTo eRe
PicWidth(x - 1) = CLng(STD1(x - 1).width * (567 / 1000))
PicHeight(x - 1) = CLng(STD1(x - 1).height * (567 / 1000))
Erase TMPDATAX
Erase SINGLEMEM
Next x
INFO.Show

ElseIf CASSE = "12" Then
'KURSOR-GRUPA
On Error GoTo eRe
If LoadIntoMemory(TrueName, TypePtr, OtherData) Then GoTo erXv1
'Uzmi broj clanova
CopyMemory NEWH, OtherData(0), Len(NEWH)
ReDim CRSDIR(NEWH.rescountX - 1)
For x = 1 To NEWH.rescountX
'Popuni Clanove sa informacijama***potpuno isto kao kod ikona
CopyMemory CRSDIR(x - 1), OtherData(6 + (x - 1) * Len(CRSDIR(0))), Len(CRSDIR(0))
Next x
ReDim STD1(NEWH.rescountX - 1)
ReDim PicWidth(NEWH.rescountX - 1)
ReDim PicHeight(NEWH.rescountX - 1)
For x = 1 To NEWH.rescountX
'Simuliraj svaku pojedinacnu kursor!!!!!!!
ReDim TMPDATAX(22 - 1 + CRSDIR(x - 1).bitesinres)
CopyMemory TMPDATAX(0), CInt(0), 2
CopyMemory TMPDATAX(2), CInt(2), 2 'Radi se o kursoru!
CopyMemory TMPDATAX(4), CInt(1), 2
CRSDIR(x - 1).curresD.height = CRSDIR(x - 1).curresD.width 'Popravi Height jer nije ispravan
CopyMemory TMPDATAX(14), CRSDIR(x - 1).bitesinres, 4
CopyMemory TMPDATAX(18), CLng(&H16), 4
If LoadIntoMemory(CLng(CRSDIR(x - 1).cursorID), CLng(1), SINGLEMEM) Then GoTo erXv1
'Uzmi pravu dimenziju iz RAW data,jer znaju biti krivi podaci...
TMPDATAX(6) = SINGLEMEM(8)
TMPDATAX(7) = SINGLEMEM(8)
CRSDIR(x - 1).curresD.width = SINGLEMEM(8)
CRSDIR(x - 1).curresD.height = SINGLEMEM(8)
CopyMemory TMPDATAX(22), SINGLEMEM(4), UBound(SINGLEMEM) + 1 - 4 'Ne ulazi HOTSPOT
CopyMemory CRSDIR(x - 1).hotXY, SINGLEMEM(0), 4 'Kopiraj HOTSPOT
CopyMemory TMPDATAX(10), CRSDIR(x - 1).hotXY, 4

Set STD1(x - 1) = GetPicture(TMPDATAX)

If Not CBool(STD1(x - 1)) Then GoTo eRe
PicWidth(x - 1) = CLng(STD1(x - 1).width * (567 / 1000))
PicHeight(x - 1) = CLng(STD1(x - 1).height * (567 / 1000))
Erase TMPDATAX
Erase SINGLEMEM
Next x
INFO.Show

ElseIf CASSE = "5" Then
'DIALOG
SetWinPosByCursor Form2.hWnd, 0
'HHDL = CreateDialogParam(HMOD, TrueName, Form2.hWnd, AddressOf dialogProc, 0&)
If LoadIntoMemory(TrueName, TypePtr, OtherData) Then GoTo erXv1
GetDialogInformation
GoSub CheckERX
HHDL = CreateDialogIndirectParam(HMOD, ByVal VarPtr(OtherData(0)), Form2.hWnd, AddressOf dialogProc, 0&)
If Not CBool(HHDL) Then Unload INFO: MsgBox "Unable to Display!", vbCritical, "Error!": Exit Sub
Dim RCT1 As RECT
Dim RCT2 As RECT
Dim PT1 As POINTAPI
GetWindowRect HHDL, RCT1
GetWindowRect Form2.hWnd, RCT2
PT1.x = RCT2.Left
PT1.y = RCT2.Top
ScreenToClient Form2.hWnd, PT1
SetParent HHDL, Form2.hWnd
Dim MTY As Long
MTY = GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYBORDER) + GetSystemMetrics(SM_CYDLGFRAME)
Dim MTX As Long
MTX = GetSystemMetrics(SM_CXBORDER) + GetSystemMetrics(SM_CXDLGFRAME)
SetWindowPos HHDL, 0, PT1.x + MTX, PT1.y + MTY, 0, 0, 1
Form2.Top = 0
Form2.Left = 0
Form2.width = (RCT1.Right - RCT1.Left) * 15 + (MTX + GetSystemMetrics(SM_CXBORDER) + GetSystemMetrics(SM_CXDLGFRAME)) * 15
Form2.height = (RCT1.Bottom - RCT1.Top) * 15 + (MTY + GetSystemMetrics(SM_CYBORDER) + GetSystemMetrics(SM_CYDLGFRAME)) * 15


ElseIf CASSE = "23" Then GoTo HDoc

ElseIf CASSE = "4" Then
If LoadIntoMemory(TrueName, TypePtr, OtherData) Then GoTo erXv1
MHDL = LoadMenuIndirect(ByVal VarPtr(OtherData(0)))
'MHDL = LoadMenu(HMOD, TrueName)
If Not CBool(MHDL) Then Exit Sub
Dim USTRX As String
Dim chrllen As Long
Dim MnuCNT As Long
Dim IID As Long
MnuCNT = GetMenuItemCount(MHDL)
For u = 0 To MnuCNT - 1
USTRX = Space(255)
chrllen = GetMenuString(MHDL, u, USTRX, 255, MF_BYPOSITION)
USTRX = Left(USTRX, chrllen)
If USTRX = "" Then
IID = GetMenuItemID(MHDL, u)
Dim modd() As Byte
modd = StrConv("Hiden PopUp" & Chr(CByte(0)), vbFromUnicode)
Call ModifyMenu(MHDL, u, 0 Or MF_BYPOSITION, IID, VarPtr(modd(0)))
End If
Next u
INFO.Show


ElseIf CASSE = "1" Then
If LoadIntoMemory(TrueName, TypePtr, OtherData) Then GoTo erXv1
ReDim SNGDATAX(22 - 1 + ResTotLen)
SNGDATAX(0) = 0
SNGDATAX(2) = 2 'KURSOR=2
SNGDATAX(4) = 1 'Jedan kursor
SNGDATAX(6) = OtherData(8) 'Width
SNGDATAX(7) = OtherData(8) 'Height
CopyMemory SNGDATAX(10), OtherData(0), 4 'HotSpot
CopyMemory SNGDATAX(14), ResTotLen, 4 'Dužina
CopyMemory SNGDATAX(18), CLng(&H16), 4 'Entry Point
CopyMemory SNGDATAX(22), OtherData(4), ResTotLen - 4 'Kopiraj RAW Podatke
icorcr = True
GoTo IKON_CURSOR

ElseIf CASSE = "3" Then
If LoadIntoMemory(TrueName, TypePtr, OtherData) Then GoTo erXv1
ReDim SNGDATAX(22 - 1 + ResTotLen)
SNGDATAX(0) = 0
SNGDATAX(2) = 1 'IKONA=1
SNGDATAX(4) = 1 'Jedna Ikona
SNGDATAX(6) = OtherData(4) 'Width
SNGDATAX(7) = OtherData(4) 'Height
CopyMemory SNGDATAX(10), OtherData(12), 2 'Kopiraj BIT PLANES
CopyMemory SNGDATAX(12), OtherData(14), 2 'Kopiraj BIT COUNT
CopyMemory SNGDATAX(14), ResTotLen, 4 'Dužina
CopyMemory SNGDATAX(18), CLng(&H16), 4 'Entry Point
Dim BPTMP As Integer
CopyMemory BPTMP, SNGDATAX(12), 2
If BPTMP = 4 Then
'Postavi broj boja prema bitcount-u...
SNGDATAX(8) = CByte(16)
ElseIf BPTMP = 16 Then
SNGDATAX(8) = 0
ElseIf BPTMP = 1 Then
SNGDATAX(8) = 2
Else
SNGDATAX(8) = 0
End If
CopyMemory SNGDATAX(22), OtherData(0), ResTotLen
'Kopiraj RAW podatke
GoTo IKON_CURSOR


Else
HDoc:
If LoadIntoMemory(TrueName, TypePtr, OtherData) Then GoTo erXv1
Dim TestHEAD As String
Dim dlenX As Long
dlenX = 100
If UBound(OtherData) + 1 < 100 Then dlenX = UBound(OtherData) + 1
TestHEAD = Space(dlenX)
CopyMemory ByVal TestHEAD, ByVal VarPtr(OtherData(0)), dlenX

If Left(TestHEAD, 3) = "GIF" Then
If Dir(GetAppPath(App.Path) & "test.x.y") <> "" Then Kill GetAppPath(App.Path) & "test.x.y"
Dim TST As String
Open GetAppPath(App.Path) & "test.x.y" For Binary As #1
Put #1, , OtherData
Close #1
TST = StrConv(OtherData, vbUnicode)
CNTG = Split(TST, Chr(0) & Chr(33) & Chr(249), , vbBinaryCompare)
TST = Space(0)
CASSE2 = "Gif.Gif"
INFO.Show

ElseIf Left(TestHEAD, 4) = Chr(&HFF) & Chr(&HD8) & Chr(&HFF) & Chr(&HE0) Then
CASSE2 = "Jpg.Jpg"
GoSub LoadPIC
If Not CBool(STD1(0)) Then CASSE2 = ""
INFO.Show

ElseIf Left(TestHEAD, 2) = "BM" Then
GoSub LoadPIC
If Not CBool(STD1(0)) Then CASSE2 = ""
CASSE2 = "Bmp.Bmp"
INFO.Show


ElseIf Left(TestHEAD, 4) = Chr(&HD7) & Chr(&HCD) & Chr(&HC6) & Chr(&H9A) Then
CASSE2 = "Wmf.Wmf"
GoSub LoadPIC
If Not CBool(STD1(0)) Then CASSE2 = ""
INFO.Show

ElseIf Left(TestHEAD, 4) = "RIFF" Then
If Mid(TestHEAD, 9, 4) = "WAVE" Then

CASSE2 = "Wav.Wav"
INFO.Show
End If

If Mid(TestHEAD, 9, 3) = "AVI" Then
CASSE2 = "Avi.Avi"
INFO.Show
End If

ElseIf InStr(1, UCase(TestHEAD), "<HTML>", vbBinaryCompare) <> 0 Then
CASSE2 = "Html.Html": INFO.Show


Else
'Za sve ostalo
INFO.Show

End If
End If



Exit Sub

IKON_CURSOR:
ReDim STD1(0)
ReDim PicWidth(0)
ReDim PicHeight(0)
If icorcr Then
Set STD1(0) = GetPicture(SNGDATAX)
Else
Set STD1(0) = GetIconToPicture(SNGDATAX)
End If
If Not CBool(STD1(0)) Then GoTo eRe
PicWidth(0) = CLng(STD1(0).width * (567 / 1000))
PicHeight(0) = CLng(STD1(0).height * (567 / 1000))
INFO.Show
Exit Sub


LoadPIC:
ReDim STD1(0)
ReDim PicWidth(0)
ReDim PicHeight(0)
Set STD1(0) = GetPicture(OtherData)
PicWidth(0) = CLng(STD1(0).width * (567 / 1000))
PicHeight(0) = CLng(STD1(0).height * (567 / 1000))
Return

eRe:
On Error GoTo 0
ERRORX = True

CheckERX:
If ERRORX Then
MsgBox "Format Error! Switch to Binary", vbCritical, "Error"
CASSE = ""
If LoadIntoMemory(TrueName, TypePtr, OtherData) Then GoTo erXv1
INFO.Show
Exit Sub
End If
Return

erXv1:
MsgBox "Could not load into memory!", vbCritical, "Error"
End Sub







Private Sub Command2_Click()

aa = GetOpenFilePath(hWnd, "", 0, sFile, "", "Get OCX,DLL,EXE...", spath)
If aa = False Then Exit Sub
Unload Form2
Unload Form3

Set EXPLIST = Nothing
FreeLibrary HMOD1
DoEvents
EraseFromMem
List1.Clear
ClearCOLLECTION
'HMOD = LoadLibraryEx(spath, 0&, 2)
HMOD1 = LoadLibrary(spath)
If Not CBool(HMOD1) Then
MsgBox "File is not an Executable File!", vbCritical, "Error!"
Label2(6).caption = "Working With:"
WorkingFile = ""
Exit Sub
End If
HMOD = HMOD1
Label2(6).caption = "Working With:" & sFile
WorkingFile = spath
EnumRX

End Sub
Private Sub EnumRX()
ClearCOLLECTION
Call EnumResourceTypes(HMOD, AddressOf EnumRSType, 0)
Dim ccnt As Long
ccnt = RESNAME.count
If ccnt = 0 Then Exit Sub
For u = 1 To ccnt
List1.AddItem RESTYPENAME.item(u) & vbTab & RESNAME.item(u) & vbTab & RESLANGID.item(u)
Next u
End Sub
Private Sub Command4_Click()
EXPORTtoList
ClearSEL List1
End Sub

Private Sub Command6_Click()
Form5.Show 1
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Command8_Click()
aa = GetOpenFilePath(hWnd, "Resource File (*.res)" & vbNullChar & "*.res", 0, sFile, "", "Get Resource File", spath)
If aa = False Then Exit Sub
DoEvents
ClearRESCOLL
FreeLibrary HMOD2
GetDataFile spath, ResData

If Not CBool(EnumRESFile) Then
List2.Clear
MsgBox "Error in Resource File!", vbCritical, "Error"
Label2(7).caption = "Working With:"
Exit Sub
End If
Label2(7).caption = "Working With:" & sFile
WorkingResFile = spath
EnumRXX
End Sub

Private Sub EnumRXX()
List2.Clear
For u = 1 To RESTYPENAMEFILE.count
List2.AddItem RESTYPENAMEFILE.item(u) & vbTab & RESNAMEFILE.item(u) & vbTab & RESLANGIDFILE.item(u)
Next u
End Sub



Private Sub Command9_Click()
If List2.ListCount = 0 Then Exit Sub
aa = GetSaveFilePath(hWnd, "DLL (*.dll)" & vbNullChar & "*.dll", 0, "", "", "", "Save As Dll", spath)
If aa = False Then Exit Sub
If Dir(spath) <> "" Then Kill spath
MakeDllRes List2
CopyFile GetAppPath(App.Path) & "precomp.work", spath, 0
End Sub

Private Sub Form_Load()
PATHPATH = GetAppPath(App.Path) & "precomp.work"
Top = (Screen.height - height) / 2
Left = (Screen.width - width) / 2
Dim tabs(1) As Long
tabs(0) = 115
tabs(1) = 0
SendMessage List1.hWnd, LB_SETTABSTOPS, 1, tabs(0)
SendMessage List2.hWnd, LB_SETTABSTOPS, 1, tabs(0)
Erase tabs

InsertCtrl Picture2.hWnd, hWnd, 10, 235
SetWidth = 330

End Sub

Private Sub Form_Unload(Cancel As Integer)
ClearCOLLECTION
FreeLibrary HMOD1
FreeLibrary HMOD2
Unload INFO
Erase TrueType
Erase TrueBuffer
Erase OtherData
ClearRESCOLL
Erase ResData
If Dir(PATHPATH) <> "" Then Kill PATHPATH
End Sub

Private Sub List1_Click()
HMOD = HMOD1
WMOD = 1
LindexW List1
End Sub

Private Sub List1_DblClick()
HMOD = HMOD1
WMOD = 1
LindexW List1
CheckClick List1
End Sub
Private Sub LindexW(listX As ListBox)
Dim lXPoint As Long
Dim lYPoint As Long
lXPoint = CLng(Xpoint / Screen.TwipsPerPixelX)
lYPoint = CLng(Ypoint / Screen.TwipsPerPixelY)
'lindex je trenutni index
lIndex = SendMessage(listX.hWnd, LB_ITEMFROMPOINT, 0, ByVal ((lYPoint * 65536) + lXPoint))
End Sub


Private Function EXPORTtoList()
Dim chk As Byte
If List1.SelCount = 0 Then MsgBox "Nothing to Export!", vbInformation, "Information": Exit Function
For u = 0 To List1.ListCount - 1
If List1.Selected(u) Then
If CStr(RESTYPE.item(u + 1)) = "3" Or CStr(RESTYPE.item(u + 1)) = "1" Then
chk = chk Or 2
Else
If CheckName(u + 1) Then
chk = chk Or 1
Else
EXPLIST.Add u + 1
End If
End If
End If
Next u
If (chk And 1) = 1 Then MsgBox "Some Resources is already in Exportlist!", vbInformation, "Information"
If (chk And 2) = 2 Then MsgBox "Only Group of Icon or Cursor could be able to export!", vbInformation, "Information"
End Function

Private Function CheckName(ByVal position As Long) As Boolean
For u = 1 To EXPLIST.count
If CLng(EXPLIST.item(u)) = position Then CheckName = True: Exit Function
Next u
End Function


Private Function EraseFromMem()
Unload INFO
Unload Form4
Unload Form2
Unload Form3
Unload HDUMP
Unload Form7
Unload Form6
Erase TrueType
Erase TrueBuffer
Erase OtherData
Erase RSDIR
Erase CRSDIR
Erase CONTROLS1
Erase CONTROLS2
End Function

Private Sub ClearSEL(listX As ListBox)
For u = 0 To listX.ListCount - 1
If listX.Selected(u) = True Then listX.Selected(u) = False
Next u
End Sub
Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then List1_DblClick
End Sub
Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Xpoint = x
Ypoint = y
End Sub
Private Sub List2_Click()
HMOD = HMOD2
WMOD = 2
LindexW List2
End Sub
Private Sub List2_DblClick()
WMOD = 2
LindexW List2
MakeDllRes List2
HMOD = HMOD2
CheckClick List2
End Sub
Private Sub List2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then List2_DblClick
End Sub
Private Sub List2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Xpoint = x
Ypoint = y
End Sub
