VERSION 5.00
Begin VB.Form HDUMP 
   Caption         =   "Binary Preview"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   Icon            =   "FormD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Left            =   3960
      TabIndex        =   2
      Top             =   6000
      Width           =   1575
   End
   Begin VB.VScrollBar vs1 
      Height          =   5775
      Left            =   9360
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox TextX 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "HDUMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents VSCROLL As CLongScroll
Attribute VSCROLL.VB_VarHelpID = -1

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

Top = (Screen.height - height) / 2
Left = (Screen.width - width) / 2

vs1.Visible = True

Set VSCROLL = New CLongScroll
Set VSCROLL.Client = vs1


If UBound(OtherData) < 16 * 25 Then
vs1.Visible = False

If CASSE = "2" Then
PrintDump TextX, TMPDATAX, 0
Else
PrintDump TextX, OtherData, 0
End If
Else
With VSCROLL
      .Min = 1
      .Max = CLng(UBound(OtherData) / 16 + 0.1) - 25 + 2
      .SmallChange = 1
      .LargeChange = 5
      .value = 0
End With
End If



End Sub

Private Sub VSCROLL_Change()
If CASSE = "2" Then
PrintDump TextX, TMPDATAX, VSCROLL.value - 1
Else
PrintDump TextX, OtherData, VSCROLL.value - 1
End If
End Sub
