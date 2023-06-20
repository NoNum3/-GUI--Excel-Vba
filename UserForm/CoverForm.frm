VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CoverForm 
   Caption         =   "UserForm2"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "CoverForm.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "CoverForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
Application.WindowState = xlMaximized
With Application
Me.Top = .Top
Me.Left = .Left
Me.Height = .Height
Me.Width = .Width

End With
Me.Enabled = False ' Disable UserForm2
End Sub

Private Sub UserForm_Initialize()
    HideTitleBar Me
End Sub
