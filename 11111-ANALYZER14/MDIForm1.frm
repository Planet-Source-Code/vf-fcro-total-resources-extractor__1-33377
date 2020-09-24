VERSION 5.00
Begin VB.MDIForm INFO 
   BackColor       =   &H00808080&
   ClientHeight    =   7320
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8940
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00AB6B6F&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   8910
      TabIndex        =   0
      Top             =   0
      Width           =   8940
      Begin VB.CommandButton Command5 
         Caption         =   "E&xit"
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
         Left            =   7920
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command4 
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
         Left            =   1080
         Picture         =   "MDIForm1.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "View Binary"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Command3 
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
         Picture         =   "MDIForm1.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "View Information"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Enabled         =   0   'False
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
         Left            =   1560
         Picture         =   "MDIForm1.frx":0B5E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Play Multimedia"
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
         Left            =   120
         Picture         =   "MDIForm1.frx":0CA8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Save"
         Top             =   120
         Width           =   495
      End
   End
End
Attribute VB_Name = "INFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim unknowX As Byte
Private Sub Command1_Click()

If unknowX = 3 Then
GoTo anyT
End If

Select Case CASSE
Case "1"
SSAVE "Cursor (*.cur)" & vbNullChar & "*.cur", "Save As Icon", True

Case "3"
SSAVE "Icon (*.ico)" & vbNullChar & "*.ico", "Save As Icon", True

Case "2" 'BITMAP
SSAVE "Bitmap (*.bmp)" & vbNullChar & "*.bmp", "Save As Bitmap"
'Važno za dodati:
'BITMAP u resorce-u ide bez BITMAPFILEHEADER-a
'BITMAP PREHADER: 42 4d XX XX XX XX  00 00 00 00 YY YY YY YY (14 BYTE)
'XXXX-Length of BMP file,YYYY-Offset RAW Podataka
'Kada uzimamo iz memorije i želimo snimiti moramo dodati taj BitmapFileHeader....!!!!!!!
'Kada prebacujemo u res oduzimamo taj header...Kristalno jasno!!!
Case "14"
SaveIcon Me.hWnd, IcoFilter, "Save As Icon", 3
Case "12"
SaveCursor Me.hWnd, CurFilter, "Save As Cursor", 1
Case Is = "4", "5", "6", "16", "9"
anyT:
SSAVE "", "Save As Any"
End Select

Select Case CASSE2
Case "Gif.Gif"
SSAVE "Gif File (*.gif)" & vbNullChar & "*.gif", "Save As Gif"
Case "Avi.Avi"
SSAVE "Avi File (*.gif)" & vbNullChar & "*.avi", "Save As Avi"
Case "Html.Html"
SSAVE "Html (*.html)" & vbNullChar & "*.html" & vbNullChar & "Htm (*.htm)" & vbNullChar & "*.htm" & vbNullChar & "Dhtml (*.dhtml)" & vbNullChar & "*.dhtml", "Save As Internet File"
Case "Jpg.Jpg"
SSAVE "Jpg File (*.jpg)" & vbNullChar & "*.jpg", "Save As Jpg"
Case "Wmf.Wmf"
SSAVE "Wmf File (*.wmf)" & vbNullChar & "*.wmf", "Save As Wmf"
Case "Wav.Wav"
SSAVE "Wav File (*.wav)" & vbNullChar & "*.wav", "Save As Wav"
Case "Bmp.Bmp"
SSAVE "Bitmap (*.bmp)" & vbNullChar & "*.bmp", "Save As Bitmap"

End Select

End Sub
Private Sub SSAVE(ByVal gfilter As String, ByVal ttext As String, Optional what As Boolean)
Dim sPath As String
If Not CBool(GetSaveFilePath(hWnd, gfilter, 0, gfilter, "", "", ttext, sPath)) Then Exit Sub
If Dir(sPath) <> "" Then Kill sPath
DoEvents
Open sPath For Binary As #1
If what = True Then
Put #1, , SNGDATAX
Else
Put #1, , OtherData
End If
Close #1
End Sub

Private Sub Command2_Click()
If CASSE2 = "Avi.Avi" Then
PlayAVIPictureBox GetAppPath(App.Path) & "temp.temp.avi", Form6.Picture1(0)
ElseIf CASSE2 = "Wav.Wav" Then
PlaySound_Res ByVal VarPtr(OtherData(0)), 0, &H4 Or &H1
End If
End Sub

Private Sub Command3_Click()
Form3.Show
End Sub

Private Sub Command4_Click()
HDUMP.Show
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub MDIForm_Load()

unknowX = 0
Top = (Screen.height - height) / 2
Left = (Screen.width - width) / 2

'Koju cemo ucitati?
Select Case CASSE
Case Is = "2", "12", "14", "1", "3"
Load Form6
Case Is = "5"
Load Form2
Case Is = "6", "11"
Load Form7
End Select

'Za nestandardne,ali prepoznate formate
Select Case CASSE2
Case "Gif.Gif"
Load Form4
Case Is = "Avi.Avi", "Jpg.Jpg", "Wmf.Wmf", "Bmp.Bmp"
Load Form6
Case "Html.Html"
Load Form4


End Select

Select Case CASSE

Case "1"
ShowSGNInfo
caption = "Single Cursor"


Case "2"
ShowPicInfo
OINF = OINF & "Length:" & (UBound(TMPDATAX) + 1) & " (" & Hex(UBound(TMPDATAX) + 1) & "h) Bytes" & vbCrLf
caption = "Bitmap"


Case "3"
ShowSGNInfo2
caption = "Single Icon"


Case "5"
caption = "Dialog Box"


Case "6"
caption = "Text Table"
ShowSTR

Case "11"
caption = "Message Table"
ShowTBL


Case "4"
caption = "Menu"
ShowUNK

Case "12"
ShowGroupInfo2
caption = "Cursor Group"


Case "14"
ShowGroupInfo
caption = "Icon Group"

Case "16"
caption = "Version Information"

Case 9
caption = "Accelerator Table"


Case Else
unknowX = unknowX Or 1

End Select


Select Case CASSE2
Case "Gif.Gif"
caption = "Gif"


Case "Avi.Avi"
caption = "Avi"
Command2.Enabled = True


Case "Html.Html"
caption = "Internet Document"

Case "Jpg.Jpg"
caption = "Jpg"
ShowPX

Case "Wmf.Wmf"
caption = "Wmf"
ShowPX

Case "Wav.Wav"
caption = "Wav"
Command2.Enabled = True
PlaySound_Res ByVal VarPtr(OtherData(0)), 0, &H4 Or &H1
ShowUNK

Case "Bmp.Bmp"
caption = "Bitmap"
OINF = OINF & "Length:" & (UBound(OtherData) + 1) & " (" & Hex(UBound(OtherData) + 1) & "h) Bytes" & vbCrLf
ShowPicInfo

Case Else
unknowX = unknowX Or 2
End Select

If unknowX = 3 Then
caption = "Unknow Data"
ShowUNK
End If


End Sub

Private Sub ShowPX()
OINF = "Resource ID:" & GRPinfo & vbCrLf
OINF = OINF & "Length:" & ResTotLen & " (" & Hex(ResTotLen) & "h) Bytes" & vbCrLf
OINF = OINF & "Width:" & CLng(PicWidth(0) / 15) & vbCrLf
OINF = OINF & "Height:" & CLng(PicHeight(0) / 15) & vbCrLf
End Sub

Private Sub ShowUNK()
OINF = "Resource ID:" & GRPinfo & vbCrLf
OINF = OINF & "Length:" & ResTotLen & " (" & Hex(ResTotLen) & "h) Bytes"
End Sub
Private Sub ShowTBL()
OINF = "Resource ID:" & GRPinfo & vbCrLf
OINF = OINF & "Number of Group Entries:" & (UBound(MRB) + 1) & vbCrLf
OINF = OINF & "Total Number of Entries:" & (UBound(RTTEXT) + 1) & vbCrLf
OINF = OINF & "Length:" & ResTotLen & " (" & Hex(ResTotLen) & "h) Bytes" & vbCrLf & vbCrLf

For u = 0 To UBound(MRB)
OINF = OINF & "Group " & MRB(u).lowId & " To " & MRB(u).HighId & vbCrLf
OINF = OINF & "Number of Entries:" & (MRB(u).HighId - MRB(u).lowId + 1) & vbCrLf & vbCrLf
Next u

End Sub

Private Sub ShowSTR()
OINF = "Resource ID:" & GRPinfo & vbCrLf
OINF = OINF & "Number of Entries:" & (UBound(RTTEXT) + 1) & vbCrLf
OINF = OINF & "Length:" & ResTotLen & " (" & Hex(ResTotLen) & "h) Bytes"
End Sub

Private Sub MDIForm_Resize()
Call SetMenu(INFO.hWnd, MHDL)
DrawMenuBar INFO.hWnd
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If CASSE = "14" Then Erase RSDIR: Erase PicWidth: Erase PicHeight
If CASSE = "12" Then Erase CRSDIR: Erase PicWidth: Erase PicHeight
If CASSE = "1" Or CASSE = "3" Then Erase SNGDATAX
If CASSE = "5" Then UNREGunknowClass
Unload Form6
Unload Form2
Unload Form4
Unload Form3
Unload Form7
Unload HDUMP
DestroyMenu HHDL
OINF = ""
End Sub
Private Sub ShowSGNInfo()
OINF = "Single Icon ID:" & GRPinfo & vbCrLf
OINF = OINF & "File Length:" & ResTotLen & " (" & Hex(ResTotLen) & "h) Bytes" & vbCrLf
OINF = OINF & "Width:" & SNGDATAX(6) & vbCrLf
OINF = OINF & "Height:" & SNGDATAX(7) & vbCrLf
OINF = OINF & "Hot Spot X:" & SNGDATAX(10) & vbCrLf
OINF = OINF & "Hot Spot Y:" & SNGDATAX(12) & vbCrLf
End Sub

Private Sub ShowSGNInfo2()
OINF = "Single Cursor ID:" & GRPinfo & vbCrLf
OINF = OINF & "Length:" & ResTotLen & " (" & Hex(ResTotLen) & "h) Bytes" & vbCrLf
OINF = OINF & "Width:" & SNGDATAX(6) & vbCrLf
OINF = OINF & "Height:" & SNGDATAX(7) & vbCrLf
Dim boja As Integer
boja = CInt(SNGDATAX(8))
If Not CBool(SNGDATAX(8)) Then boja = 256 'If number of colors =0 then number of colors=256
OINF = OINF & "Colors:" & boja & vbCrLf & vbCrLf
End Sub

Private Sub ShowGroupInfo2()
'KURSORI
OINF = "Cursor Group ID:" & GRPinfo & vbCrLf
OINF = OINF & "Number of Cursors:" & NEWH.rescountX & vbCrLf

OINF = OINF & "Length:" & ResTotLen & " (" & Hex(ResTotLen) & "h) Bytes" & vbCrLf & vbCrLf
For x = 1 To NEWH.rescountX
OINF = OINF & "Single Cursor ID:" & CRSDIR(x - 1).cursorID & vbCrLf
OINF = OINF & "Width:" & CRSDIR(x - 1).curresD.width & vbCrLf
OINF = OINF & "Height:" & CRSDIR(x - 1).curresD.height & vbCrLf
OINF = OINF & "Hot Spot X:" & ((CRSDIR(x - 1).hotXY And &HFFFF0000) / &H10000) & vbCrLf
OINF = OINF & "Hot Spot Y:" & (CRSDIR(x - 1).hotXY And &HFFFF&) & vbCrLf & vbCrLf
Next x
End Sub

Private Sub ShowGroupInfo()
'IKONE
OINF = "Icon Group ID:" & GRPinfo & vbCrLf
OINF = OINF & "Number of Icons:" & NEWH.rescountX & vbCrLf
OINF = OINF & "Length:" & ResTotLen & " (" & Hex(ResTotLen) & "h) Bytes" & vbCrLf & vbCrLf
For x = 1 To NEWH.rescountX
OINF = OINF & "Single Icon ID:" & RSDIR(x - 1).iconID & vbCrLf
OINF = OINF & "Width:" & RSDIR(x - 1).iconresD.width & vbCrLf
OINF = OINF & "Height:" & RSDIR(x - 1).iconresD.height & vbCrLf
Dim boja As Integer
boja = CInt(RSDIR(x - 1).iconresD.colorCount)
If Not CBool(boja) Then boja = 256 'If number of colors =0 then number of colors=256
OINF = OINF & "Colors:" & boja & vbCrLf & vbCrLf
Next x
End Sub


Private Sub ShowPicInfo()
'BITMAPE
Dim Pinfo As BITMAPINFOHEADER
CopyMemory Pinfo, OtherData(14), Len(Pinfo) 'Uzmi informacije u Picture!
OINF = "Bitmap ID:" & GRPinfo & vbCrLf & vbCrLf
OINF = OINF & "Width:" & Pinfo.biWidth & vbCrLf
OINF = OINF & "Height:" & Pinfo.biHeight & vbCrLf
Select Case Pinfo.biBitCount
Case 1
OINF = OINF & "Monochrome Picture" & vbCrLf
Case 4
OINF = OINF & "16 Colors" & vbCrLf
Case 8
OINF = OINF & "256 Colors" & vbCrLf
Case 16
OINF = OINF & "16-bit RGB Colors" & vbCrLf
Case 24
OINF = OINF & "24-bit RGB Colors" & vbCrLf
Case 32
OINF = OINF & "32-bit RGB Colors" & vbCrLf
End Select
Select Case Pinfo.biCompression
Case Is = 0, 3
OINF = OINF & "Not Compressed" & vbCrLf
Case 1
OINF = OINF & "RLE8 Compression" & vbCrLf
Case 2
OINF = OINF & "RLE4 Compression" & vbCrLf
End Select
End Sub




