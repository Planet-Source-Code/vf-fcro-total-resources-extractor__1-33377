VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form4 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "GIF"
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1995
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   1995
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser doc1 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   855
      ExtentX         =   1508
      ExtentY         =   1085
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents anim1 As VBControlExtender
Attribute anim1.VB_VarHelpID = -1
Private Sub Form_Activate()
If CASSE2 = "Gif.Gif" Then
Me.width = anim1.width '+ 15 * 2 * GetSystemMetrics(SM_CXBORDER)
Me.height = anim1.height ' + 15 * GetSystemMetrics(SM_CYCAPTION) + 15 * 2 * GetSystemMetrics(SM_CYBORDER)
ElseIf CASSE2 = "Html.Html" Then
doc1.Top = 0
doc1.Left = 0
doc1.width = INFO.ScaleWidth
doc1.height = INFO.ScaleHeight
width = doc1.width
height = doc1.height
End If
End Sub

Private Sub Form_Load()
If CASSE2 = "Gif.Gif" Then
Set anim1 = Controls.Add("Gif89.Gif89.1", "anim1", Form4)
anim1.Move 1, 1, 10, 10
anim1.Visible = True
doc1.Visible = False
anim1.object.AutoSize = True
anim1.object.Play
anim1.object.filename = GetAppPath(App.Path) & "test.x.y"
Kill GetAppPath(App.Path) & "test.x.y"

OINF = "Gif Resource ID:" & GRPinfo & vbCrLf
OINF = OINF & "Width:" & CLng(anim1.width / 15) & vbCrLf
OINF = OINF & "Height:" & CLng(anim1.height / 15) & vbCrLf
OINF = OINF & "Number of Gifs:"
If Not CBool(UBound(CNTG)) Or UBound(CNTG) = 1 Then
OINF = OINF & "1" & vbCrLf
Else
OINF = OINF & UBound(CNTG) - 1 & vbCrLf
End If
OINF = OINF & "Length:" & ResTotLen & " (" & Hex(ResTotLen) & "h) Bytes"

Erase CNTG
ElseIf CASSE2 = "Html.Html" Then
OINF = "HTML Resource ID:" & GRPinfo & vbCrLf
OINF = OINF & "Length:" & ResTotLen & " (" & Hex(ResTotLen) & "h) Bytes" & vbCrLf
Open GetAppPath(App.Path) & "test.test.html" For Binary As #1
Put #1, , OtherData
Close #1
doc1.Navigate GetAppPath(App.Path) & "test.test.html"

End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
If CASSE2 = "Gif.Gif" Then
anim1.object.Stop
ElseIf CASSE2 = "Html.Html" Then
Kill GetAppPath(App.Path) & "test.test.html"
End If
End Sub
