VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource Extractor"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2445
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   2445
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
ShowWindow HHDL, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.SetFocus
End Sub
