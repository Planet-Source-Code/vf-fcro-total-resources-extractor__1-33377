VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7305
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Entry / Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
width = INFO.ScaleWidth
height = INFO.ScaleHeight

Text1.width = width
Text1.height = height - 300

Label1(1).width = Text1.width - Label1(1).Left

Dim tabs(2) As Long
tabs(0) = 60
tabs(1) = 0

Const EM_SETTABSTOPS = &HCB
SendMessage Text1.hWnd, EM_SETTABSTOPS, 1, tabs(0)
Erase tabs
Text1 = ""

'STRING TABLA i MESSAGE TABLA
For u = 0 To UBound(RTTEXT)
'Protekcija da nebi pisao u novi red!
Text1 = Text1 & RTTEXT(u).id & vbTab & Replace(RTTEXT(u).data, Chr$(&HD) & Chr$(&HA), "  ", , , vbBinaryCompare) & vbCrLf
Next u

End Sub

Private Sub Form_Unload(Cancel As Integer)
Erase RTTEXT
Erase MRB
End Sub
