VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   0
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   1695
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
Select Case CASSE2
Case "Avi.Avi"
DoEvents
PlayAVIPictureBox GetAppPath(App.Path) & "temp.temp.avi", Form6.Picture1(0)
End Select

End Sub

Private Sub Form_Load()
Form6.Left = 0


Select Case CASSE2
Case "Avi.Avi"
On Error GoTo dalje:
Open GetAppPath(App.Path) & "temp.temp.avi" For Binary As #1
Put #1, , OtherData
Close #1
Dim pAVIStream As Long
Dim pAviFile As Long
Dim numFrames As Long
Dim afi As AVIFILEINFO
Dim ret As Long
Call AVIFileInit
If 0 <> AVIFileOpen(pAviFile, GetAppPath(App.Path) & "temp.temp.avi", OF_SHARE_DENY_WRITE, 0&) Then
MsgBox "Not an valid AVI format.", vbCritical, "Error"
Kill GetAppPath(App.Path) & "temp.temp.avi"
Exit Sub
End If
Call AVIGetFileInfo(pAviFile, afi, Len(afi))

OINF = "Avi Resource ID:" & GRPinfo & vbCrLf
OINF = OINF & "Width:" & afi.dwWidth & vbCrLf
OINF = OINF & "Height:" & afi.dwHeight & vbCrLf
res = AVIFileGetStream(pAviFile, pAVIStream, 1935960438, 0)
numFrames = AVIStreamLength(pAVIStream)
OINF = OINF & "Stream Frames:" & numFrames & vbCrLf
OINF = OINF & "Samples Per Sec:" & CLng(afi.dwRate / afi.dwScale) & vbCrLf
OINF = OINF & "Play Time:" & CLng(numFrames / CLng(afi.dwRate / afi.dwScale)) & "." & (numFrames Mod CLng(afi.dwRate / afi.dwScale)) & " Sec" & vbCrLf
If ((afi.dwCaps And AVIFILECAPS_NOCOMPRESSION) = AVIFILECAPS_NOCOMPRESSION) Then
OINF = OINF & "Not Compressed" & vbCrLf
Else
OINF = OINF & "Compressed" & vbCrLf
End If
OINF = OINF & "Length:" & ResTotLen & " (" & Hex(ResTotLen) & "h) Bytes" & vbCrLf


Picture1(0).height = afi.dwHeight * 15
Picture1(0).width = afi.dwWidth * 15
width = Picture1(0).width
height = Picture1(0).height
Call AVIStreamRelease(pAVIStream)
Call AVIFileRelease(pAviFile)
Call AVIFileExit
Exit Sub
dalje:
On Error GoTo 0
Kill GetAppPath(App.Path) & "temp.temp.avi"
Close #1

Case Is = "Jpg.Jpg", "Wmf.Wmf", "Bmp.Bmp"
GoTo GPIC
End Select



Select Case CASSE
Case Is = "2", "1", "3"
GPIC:
Set Picture1(0).Picture = STD1(0)
Erase STD1
Picture1(0).width = PicWidth(0)
Picture1(0).height = PicHeight(0)
width = PicWidth(0)
height = PicHeight(0)

Case Is = "14", "12"
Dim allWidth As Long
Dim allHeight As Long
Set Picture1(0).Picture = STD1(0)
Picture1(0).width = PicWidth(0)
Picture1(0).height = PicHeight(0)
allWidth = PicWidth(0)
allHeight = PicHeight(0)
For x = 2 To NEWH.rescountX
Load Picture1(x - 1)
Picture1(x - 1).Left = Picture1(x - 2).Left + Picture1(x - 2).width + 15 * 5
Picture1(x - 1).Visible = True
If PicHeight(x - 1) > allHeight Then allHeight = PicHeight(x - 1)
allWidth = allWidth + PicWidth(x - 1) + 15 * 5
Set Picture1(x - 1).Picture = STD1(x - 1)
Picture1(x - 1).width = PicWidth(x - 1)
Picture1(x - 1).height = PicHeight(x - 1)
Next x
width = allWidth
height = allHeight
Erase STD1
End Select



End Sub

Private Sub Form_Unload(Cancel As Integer)
If CASSE2 = "Avi.Avi" Then
If Dir(GetAppPath(App.Path) & "temp.temp.avi") <> "" Then
Kill GetAppPath(App.Path) & "temp.temp.avi"
End If
End If
End Sub
