Attribute VB_Name = "Module2"
Sub SortData()
Application.ScreenUpdating = False
Sheet1.Select
End Sub

Sub HideTitleBar(frm As Object)

Dim Style As Long, Menu As Long, hWndForm As Long
hWndForm = FindWindow("ThunderDFrame", frm.Caption)
Style = GetWindowLong(hWndForm, &HFFF0)
Style = Style And Not &HC00000
SetWindowLong hWndForm, &HFFF0, Style
DrawMenuBar hWndForm

End Sub
