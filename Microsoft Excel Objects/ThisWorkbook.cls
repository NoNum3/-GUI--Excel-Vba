VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Application.Visible = False ' Hide the Excel application
    
    Dim appHeight As Long ' Variable to store the application height
    Dim appWidth As Long ' Variable to store the application width
    
    appHeight = Application.Height ' Get the application height
    appWidth = Application.Width ' Get the application width
    
    With CoverForm
        .Show ' Show the black cover UserForm
        .Top = 0 ' Set the top position to 0
        .Left = 0 ' Set the left position to 0
        .Width = appWidth ' Set the width to match the application width
        .Height = appHeight ' Set the height to match the application height
    End With
    
    With UserForm1
        .StartUpPosition = 0 ' Set the UserForm startup position to manual
        .Left = (appWidth - .Width) / 2 ' Center the UserForm horizontally
        .Top = (appHeight - .Height) / 2 ' Center the UserForm vertically
    End With
    
    CoverForm.Show vbModeless ' Show CoverForm as modeless
    UserForm1.Show vbModeless ' Show UserForm1 as modeless
    
    Application.Visible = True ' Show the Excel application
End Sub
