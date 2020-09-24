VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export List"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Delete All"
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
      Index           =   3
      Left            =   2160
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
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
      Index           =   2
      Left            =   6120
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete From List"
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
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save As Res"
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
      Index           =   0
      Left            =   4080
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7575
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Total Resources of Export List:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)


Select Case Index
Case 1
If List1.ListIndex = -1 Then Exit Sub
EXPLIST.Remove (List1.ListIndex + 1)
List1.RemoveItem (List1.ListIndex)
List1.Refresh

Case 0
If List1.ListCount = 0 Then Exit Sub
Const RSFilter = "Resource File (*.res)" & vbNullChar & "*.res"
Dim spath As String
aa = GetSaveFilePath(hWnd, RSFilter, 0, RSFilter, "", "", "Save RESOURCE FILE", spath)
If aa = False Then Exit Sub
If Dir(spath) <> "" Then Kill spath
DoEvents
SaveToRes spath, List1




Case 2
Unload Me
Form1.SetFocus

Case 3
Set EXPLIST = Nothing
List1.Clear


End Select
End Sub
Private Sub Form_Load()
Top = (Screen.height - height) / 2
Left = (Screen.width - width) / 2
Dim tabs(1) As Long
tabs(0) = 150
tabs(1) = 150
SendMessage List1.hWnd, LB_SETTABSTOPS, 1, tabs(0)
List1.Clear
For u = 1 To EXPLIST.count
List1.AddItem RESTYPENAME.item(EXPLIST.item(u)) & vbTab & RESNAME.item(EXPLIST.item(u)) & vbTab & RESLANGID.item(EXPLIST.item(u))
Next u
Label2 = EXPLIST.count
End Sub

