VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10692
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   18900
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMB_CTBC_Click()
Me.TXT_BANK.Value = "CTCBTWTPXXX"
End Sub

Private Sub CMB_GENDER_Change()
If Me.TXT_FIRSTNAME.Value = "" _
Or Me.TXT_LASTNAME.Value = "" _
Or Me.TXT_BIRTHDAY.Value = "" _
Or Me.CMB_GENDER.Value = "" _
Or Me.CMB_MARRIED.Value = "" _
Or Me.TXT_EMAIL.Value = "" Then
Me.CMD_NEXT1.Enabled = False
Me.S1.Caption = 1
Else
Me.CMD_NEXT1.Enabled = True
Me.S1.Caption = "OK"
End If

End Sub

Private Sub CMB_MARRIED_Change()
If Me.TXT_FIRSTNAME.Value = "" _
Or Me.TXT_LASTNAME.Value = "" _
Or Me.TXT_BIRTHDAY.Value = "" _
Or Me.CMB_GENDER.Value = "" _
Or Me.CMB_MARRIED.Value = "" _
Or Me.TXT_EMAIL.Value = "" Then
Me.CMD_NEXT1.Enabled = False
Me.S1.Caption = 1
Else
Me.CMD_NEXT1.Enabled = True
Me.S1.Caption = "OK"
End If

End Sub

Private Sub CMB_STATUS_Change()
If Me.TXT_COMPNAME.Value = "" _
Or Me.TXT_COMPADDR.Value = "" _
Or Me.TXT_POSITION.Value = "" _
Or Me.CMB_STATUS.Value = "" _
Or Me.TXT_COMPPHONE.Value = "" _
Or Me.TXT_SALARY.Value = "" Then
Me.CMD_NEXT2.Enabled = False
Me.S2.Caption = 2
Else
Me.CMD_NEXT2.Enabled = True
Me.S2.Caption = "OK"
End If

End Sub

Private Sub CMD_CANCEL_Click()
Me.TXT_FIRSTNAME.Value = ""
Me.TXT_LASTNAME.Value = ""
Me.TXT_BIRTHDAY.Value = ""
Me.CMB_GENDER.Value = ""
Me.CMB_MARRIED.Value = ""
Me.TXT_EMAIL.Value = ""
Me.TXT_COMPNAME.Value = ""
Me.TXT_COMPADDR.Value = ""
Me.TXT_POSITION.Value = ""
Me.CMB_STATUS.Value = ""
Me.TXT_COMPPHONE.Value = ""
Me.TXT_SALARY.Value = ""
Me.TXT_ADDRESS.Value = ""
Me.TXT_POSTCODE.Value = ""
Me.TXT_CITY.Value = ""
Me.TXT_PHONE.Value = ""
Me.TXT_COUNTRY.Value = ""
Me.TXT_BANK.Value = ""
Me.TXT_CARDNUM.Value = ""
Me.TXT_EXPDATE.Value = ""
Me.TXT_CVC.Value = ""
Me.CMB_GENDER.Clear
Me.CMB_MARRIED.Clear
Me.CMB_STATUS.Clear

Call UserForm_Initialize

End Sub

Private Sub CMD_CANCEL1_Click()
Me.TXT_FIRSTNAME.Value = ""
Me.TXT_LASTNAME.Value = ""
Me.TXT_BIRTHDAY.Value = ""
Me.CMB_GENDER.Value = ""
Me.CMB_MARRIED.Value = ""
Me.TXT_EMAIL.Value = ""
Me.TXT_COMPNAME.Value = ""
Me.TXT_COMPADDR.Value = ""
Me.TXT_POSITION.Value = ""
Me.CMB_STATUS.Value = ""
Me.TXT_COMPPHONE.Value = ""
Me.TXT_SALARY.Value = ""
Me.TXT_ADDRESS.Value = ""
Me.TXT_POSTCODE.Value = ""
Me.TXT_CITY.Value = ""
Me.TXT_PHONE.Value = ""
Me.TXT_COUNTRY.Value = ""
Me.TXT_BANK.Value = ""
Me.TXT_CARDNUM.Value = ""
Me.TXT_EXPDATE.Value = ""
Me.TXT_CVC.Value = ""
Me.CMB_GENDER.Clear
Me.CMB_MARRIED.Clear
Me.CMB_STATUS.Clear

Call UserForm_Initialize

End Sub

Private Sub CMD_CANCEL2_Click()
Me.TXT_FIRSTNAME.Value = ""
Me.TXT_LASTNAME.Value = ""
Me.TXT_BIRTHDAY.Value = ""
Me.CMB_GENDER.Value = ""
Me.CMB_MARRIED.Value = ""
Me.TXT_EMAIL.Value = ""
Me.TXT_COMPNAME.Value = ""
Me.TXT_COMPADDR.Value = ""
Me.TXT_POSITION.Value = ""
Me.CMB_STATUS.Value = ""
Me.TXT_COMPPHONE.Value = ""
Me.TXT_SALARY.Value = ""
Me.TXT_ADDRESS.Value = ""
Me.TXT_POSTCODE.Value = ""
Me.TXT_CITY.Value = ""
Me.TXT_PHONE.Value = ""
Me.TXT_COUNTRY.Value = ""
Me.TXT_BANK.Value = ""
Me.TXT_CARDNUM.Value = ""
Me.TXT_EXPDATE.Value = ""
Me.TXT_CVC.Value = ""
Me.CMB_GENDER.Clear
Me.CMB_MARRIED.Clear
Me.CMB_STATUS.Clear

Call UserForm_Initialize

End Sub

Private Sub CMD_CANCEL3_Click()
Me.TXT_FIRSTNAME.Value = ""
Me.TXT_LASTNAME.Value = ""
Me.TXT_BIRTHDAY.Value = ""
Me.CMB_GENDER.Value = ""
Me.CMB_MARRIED.Value = ""
Me.TXT_EMAIL.Value = ""
Me.TXT_COMPNAME.Value = ""
Me.TXT_COMPADDR.Value = ""
Me.TXT_POSITION.Value = ""
Me.CMB_STATUS.Value = ""
Me.TXT_COMPPHONE.Value = ""
Me.TXT_SALARY.Value = ""
Me.TXT_ADDRESS.Value = ""
Me.TXT_POSTCODE.Value = ""
Me.TXT_CITY.Value = ""
Me.TXT_PHONE.Value = ""
Me.TXT_COUNTRY.Value = ""
Me.TXT_BANK.Value = ""
Me.TXT_CARDNUM.Value = ""
Me.TXT_EXPDATE.Value = ""
Me.TXT_CVC.Value = ""
Me.CMB_GENDER.Clear
Me.CMB_MARRIED.Clear
Me.CMB_STATUS.Clear

Call UserForm_Initialize

End Sub

Private Sub CMD_CATHAY_Click()
Me.TXT_BANK.Value = "CATHUS6LXXX"
End Sub

Private Sub CMD_DATABASE_Click()
Me.MultiPage1.Value = 4
Me.STEP4.BackColor = RGB(16, 89, 199)
End Sub

Private Sub CMD_DELETE_Click()

If Me.TXT_FIRSTNAME.Value = "" Then
    Call MsgBox("Choose a data from the table", vbInformation, "Delete Data")
Else
    Select Case MsgBox("Confirm" _
& vbCrLf & "Are you sure?" _
, vbYesNo Or vbQuestion Or vbDefaultButton1, "Delete Data")
Case vbNo
Exit Sub
Case vbYes
End Select
'Decide Where to delete data, remove the data and clear form
Set DLTDATA = Sheet1.Range("A7:A500000").Find(What:=Me.TXT_FIRSTNAME.Value, LookIn:=xlValues)
DLTDATA.Offset(0, 0).ClearContents
DLTDATA.Offset(0, 1).ClearContents
DLTDATA.Offset(0, 2).ClearContents
DLTDATA.Offset(0, 3).ClearContents
DLTDATA.Offset(0, 4).ClearContents
DLTDATA.Offset(0, 5).ClearContents
DLTDATA.Offset(0, 6).ClearContents
DLTDATA.Offset(0, 7).ClearContents
DLTDATA.Offset(0, 8).ClearContents
DLTDATA.Offset(0, 9).ClearContents
DLTDATA.Offset(0, 10).ClearContents
DLTDATA.Offset(0, 11).ClearContents
DLTDATA.Offset(0, 12).ClearContents
DLTDATA.Offset(0, 13).ClearContents
DLTDATA.Offset(0, 14).ClearContents
DLTDATA.Offset(0, 15).ClearContents
DLTDATA.Offset(0, 16).ClearContents
DLTDATA.Offset(0, 17).ClearContents
DLTDATA.Offset(0, 18).ClearContents
DLTDATA.Offset(0, 19).ClearContents
DLTDATA.Offset(0, 20).ClearContents

Call MsgBox("Data Has been deleted", vbInformation, "Removed")

Me.TXT_FIRSTNAME.Value = ""
Me.TXT_LASTNAME.Value = ""
Me.TXT_BIRTHDAY.Value = ""
Me.CMB_GENDER.Value = ""
Me.CMB_MARRIED.Value = ""
Me.TXT_EMAIL.Value = ""
Me.TXT_COMPNAME.Value = ""
Me.TXT_COMPADDR.Value = ""
Me.TXT_POSITION.Value = ""
Me.CMB_STATUS.Value = ""
Me.TXT_COMPPHONE.Value = ""
Me.TXT_SALARY.Value = ""
Me.TXT_ADDRESS.Value = ""
Me.TXT_POSTCODE.Value = ""
Me.TXT_CITY.Value = ""
Me.TXT_PHONE.Value = ""
Me.TXT_COUNTRY.Value = ""
Me.TXT_BANK.Value = ""
Me.TXT_CARDNUM.Value = ""
Me.TXT_EXPDATE.Value = ""
Me.TXT_CVC.Value = ""
Me.CMB_GENDER.Clear
Me.CMB_MARRIED.Clear
Me.CMB_STATUS.Clear

Selection.EntireRow.Delete

Call UserForm_Initialize
Call SortData
End If

End Sub

Private Sub CMD_ESUN_Click()
Me.TXT_BANK.Value = "ESUNTWTP"
End Sub

Private Sub CMD_FUBON_Click()
Me.TXT_BANK.Value = "TPBKTWTP"
End Sub

Private Sub CMD_NEXT1_Click()
Me.MultiPage1.Value = 1
Me.STEP2.BackColor = RGB(16, 89, 199)
Me.S1.Caption = "OK"

End Sub

Private Sub CMD_NEXT2_Click()
Me.MultiPage1.Value = 2
Me.STEP3.BackColor = RGB(16, 89, 199)
Me.S2.Caption = "OK"

End Sub

Private Sub CMD_NEXT3_Click()
Me.MultiPage1.Value = 3
Me.STEP4.BackColor = RGB(16, 89, 199)
Me.S3.Caption = "OK"

End Sub

Private Sub CMD_PERV3_Click()
Me.MultiPage1.Value = 2
End Sub

Private Sub CMD_PREV1_Click()
Me.MultiPage1.Value = 0
End Sub

Private Sub CMD_PREV2_Click()
Me.MultiPage1.Value = 1
End Sub



Private Sub CMD_PREV4_Click()
Me.MultiPage1.Value = 3
End Sub

Private Sub CMD_RESET_Click()
On Error Resume Next
Me.TXT_SEARCH.Value = ""
Me.DATA_TABLE.RowSource = Sheet1.Range("DATASOURCE").Address(External:=True)

End Sub

Private Sub CMD_RESETFORM_Click()
Me.TXT_FIRSTNAME.Value = ""
Me.TXT_LASTNAME.Value = ""
Me.TXT_BIRTHDAY.Value = ""
Me.CMB_GENDER.Value = ""
Me.CMB_MARRIED.Value = ""
Me.TXT_EMAIL.Value = ""
Me.TXT_COMPNAME.Value = ""
Me.TXT_COMPADDR.Value = ""
Me.TXT_POSITION.Value = ""
Me.CMB_STATUS.Value = ""
Me.TXT_COMPPHONE.Value = ""
Me.TXT_SALARY.Value = ""
Me.TXT_ADDRESS.Value = ""
Me.TXT_POSTCODE.Value = ""
Me.TXT_CITY.Value = ""
Me.TXT_PHONE.Value = ""
Me.TXT_COUNTRY.Value = ""
Me.TXT_BANK.Value = ""
Me.TXT_CARDNUM.Value = ""
Me.TXT_EXPDATE.Value = ""
Me.TXT_CVC.Value = ""
Me.CMB_GENDER.Clear
Me.CMB_MARRIED.Clear
Me.CMB_STATUS.Clear

Call UserForm_Initialize


End Sub

Private Sub CMD_SUBMIT_Click()
Dim DBCUSTOMER As Object
Set DBCUSTOMER = Sheet1.Range("A20000").End(xlUp)
Select Case MsgBox("The Costumer Data Will be Saved" _
& vbCrLf & "Are You Sure?" _
, vbYesNo Or vbQuestion Or vbDefaultButton1, "Save")
Case vbNo
Exit Sub
Case vbYes
End Select
DBCUSTOMER.Offset(1, 0).Value = Me.TXT_FIRSTNAME.Value
DBCUSTOMER.Offset(1, 1).Value = Me.TXT_LASTNAME.Value
DBCUSTOMER.Offset(1, 2).Value = Me.TXT_BIRTHDAY.Value
DBCUSTOMER.Offset(1, 3).Value = Me.CMB_GENDER.Value
DBCUSTOMER.Offset(1, 4).Value = Me.CMB_MARRIED.Value
DBCUSTOMER.Offset(1, 5).Value = Me.TXT_EMAIL.Value
DBCUSTOMER.Offset(1, 6).Value = Me.TXT_COMPNAME.Value
DBCUSTOMER.Offset(1, 7).Value = Me.TXT_COMPADDR.Value
DBCUSTOMER.Offset(1, 8).Value = Me.TXT_POSITION.Value
DBCUSTOMER.Offset(1, 9).Value = Me.CMB_STATUS.Value
DBCUSTOMER.Offset(1, 10).Value = Me.TXT_COMPPHONE.Value
DBCUSTOMER.Offset(1, 11).Value = Me.TXT_SALARY.Value
DBCUSTOMER.Offset(1, 12).Value = Me.TXT_ADDRESS.Value
DBCUSTOMER.Offset(1, 13).Value = Me.TXT_POSTCODE.Value
DBCUSTOMER.Offset(1, 14).Value = Me.TXT_CITY.Value
DBCUSTOMER.Offset(1, 15).Value = Me.TXT_PHONE.Value
DBCUSTOMER.Offset(1, 16).Value = Me.TXT_COUNTRY.Value
DBCUSTOMER.Offset(1, 17).Value = Me.TXT_BANK.Value
DBCUSTOMER.Offset(1, 18).Value = Me.TXT_CARDNUM.Value
DBCUSTOMER.Offset(1, 19).Value = Me.TXT_EXPDATE.Value
DBCUSTOMER.Offset(1, 20).Value = Me.TXT_CVC.Value

On Error Resume Next
Me.DATA_TABLE.RowSource = Sheet1.Range("DATASOURCE").Address(External:=True)

Call MsgBox("Customer Data has been Added", vbInformation, "Save Data Customer")
Me.TXT_FIRSTNAME.Value = ""
Me.TXT_LASTNAME.Value = ""
Me.TXT_BIRTHDAY.Value = ""
Me.CMB_GENDER.Value = ""
Me.CMB_MARRIED.Value = ""
Me.TXT_EMAIL.Value = ""
Me.TXT_COMPNAME.Value = ""
Me.TXT_COMPADDR.Value = ""
Me.TXT_POSITION.Value = ""
Me.CMB_STATUS.Value = ""
Me.TXT_COMPPHONE.Value = ""
Me.TXT_SALARY.Value = ""
Me.TXT_ADDRESS.Value = ""
Me.TXT_POSTCODE.Value = ""
Me.TXT_CITY.Value = ""
Me.TXT_PHONE.Value = ""
Me.TXT_COUNTRY.Value = ""
Me.TXT_BANK.Value = ""
Me.TXT_CARDNUM.Value = ""
Me.TXT_EXPDATE.Value = ""
Me.TXT_CVC.Value = ""
Me.CMB_GENDER.Clear
Me.CMB_MARRIED.Clear
Me.CMB_STATUS.Clear
Call UserForm_Initialize

End Sub

Private Sub CMD_UPDATE_Click()
Dim ROWC, CHGSOURCE As String
Dim CHGDATA As String
CHGDATA = Me.TXT_FIRSTNAME.Value

If Me.TXT_FIRSTNAME.Text = "" Then
Call MsgBox("Choose a data", vbInformation, "Pick One")
Else
ROWC = ActiveCell.Row
Cells(ROWC, 1) = Me.TXT_FIRSTNAME.Value
Cells(ROWC, 2) = Me.TXT_LASTNAME.Value
Cells(ROWC, 3) = Me.TXT_BIRTHDAY.Value
Cells(ROWC, 4) = Me.CMB_GENDER.Value
Cells(ROWC, 5) = Me.CMB_MARRIED.Value
Cells(ROWC, 6) = Me.TXT_EMAIL.Value
Cells(ROWC, 7) = Me.TXT_COMPNAME.Value
Cells(ROWC, 8) = Me.TXT_COMPADDR.Value
Cells(ROWC, 9) = Me.TXT_POSITION.Value
Cells(ROWC, 10) = Me.CMB_STATUS.Value
Cells(ROWC, 11) = Me.TXT_COMPPHONE.Value
Cells(ROWC, 12) = Me.TXT_SALARY.Value
Cells(ROWC, 13) = Me.TXT_ADDRESS.Value
Cells(ROWC, 14) = Me.TXT_POSTCODE.Value
Cells(ROWC, 15) = Me.TXT_CITY.Value
Cells(ROWC, 16) = Me.TXT_PHONE.Value
Cells(ROWC, 17) = Me.TXT_COUNTRY.Value
Cells(ROWC, 18) = Me.TXT_BANK.Value
Cells(ROWC, 19) = Me.TXT_CARDNUM.Value
Cells(ROWC, 20) = Me.TXT_EXPDATE.Value
Cells(ROWC, 21) = Me.TXT_CVC.Value

On Error Resume Next
Me.DATA_TABLE.RowSource = Sheet1.Range("DATASOURCE").Address(External:=True)
Call MsgBox("Data has been updated", vbInformation, "Update Data")

Me.TXT_FIRSTNAME.Value = ""
Me.TXT_LASTNAME.Value = ""
Me.TXT_BIRTHDAY.Value = ""
Me.CMB_GENDER.Value = ""
Me.CMB_MARRIED.Value = ""
Me.TXT_EMAIL.Value = ""
Me.TXT_COMPNAME.Value = ""
Me.TXT_COMPADDR.Value = ""
Me.TXT_POSITION.Value = ""
Me.CMB_STATUS.Value = ""
Me.TXT_COMPPHONE.Value = ""
Me.TXT_SALARY.Value = ""
Me.TXT_ADDRESS.Value = ""
Me.TXT_POSTCODE.Value = ""
Me.TXT_CITY.Value = ""
Me.TXT_PHONE.Value = ""
Me.TXT_COUNTRY.Value = ""
Me.TXT_BANK.Value = ""
Me.TXT_CARDNUM.Value = ""
Me.TXT_EXPDATE.Value = ""
Me.TXT_CVC.Value = ""
Me.CMB_GENDER.Clear
Me.CMB_MARRIED.Clear
Me.CMB_STATUS.Clear

Call UserForm_Initialize
Me.CMD_SUBMIT.Visible = True
End If

End Sub


Private Sub DATA_TABLE_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo EXCELVBA
Me.TXT_FIRSTNAME.Value = Me.DATA_TABLE.Value
Me.TXT_LASTNAME.Value = Me.DATA_TABLE.Column(1)
Me.TXT_BIRTHDAY.Value = Me.DATA_TABLE.Column(2)
Me.CMB_GENDER.Value = Me.DATA_TABLE.Column(3)
Me.CMB_MARRIED.Value = Me.DATA_TABLE.Column(4)
Me.TXT_EMAIL.Value = Me.DATA_TABLE.Column(5)
Me.TXT_COMPNAME.Value = Me.DATA_TABLE.Column(6)
Me.TXT_COMPADDR.Value = Me.DATA_TABLE.Column(7)
Me.TXT_POSITION.Value = Me.DATA_TABLE.Column(8)
Me.CMB_STATUS.Value = Me.DATA_TABLE.Column(9)
Me.TXT_COMPPHONE.Value = Me.DATA_TABLE.Column(10)
Me.TXT_SALARY.Value = Me.DATA_TABLE.Column(11)
Me.TXT_ADDRESS.Value = Me.DATA_TABLE.Column(12)
Me.TXT_POSTCODE.Value = Me.DATA_TABLE.Column(13)
Me.TXT_CITY.Value = Me.DATA_TABLE.Column(14)
Me.TXT_PHONE.Value = Me.DATA_TABLE.Column(15)
Me.TXT_COUNTRY.Value = Me.DATA_TABLE.Column(16)
Me.TXT_BANK.Value = Me.DATA_TABLE.Column(17)
Me.TXT_CARDNUM.Value = Me.DATA_TABLE.Column(18)
Me.TXT_EXPDATE.Value = Me.DATA_TABLE.Column(19)
Me.TXT_CVC.Value = Me.DATA_TABLE.Column(20)
Me.CMD_SUBMIT.Enabled = False
Sheet1.Select
CHGSOURCE = Sheets("DATACUSTOMER").Cells(Rows.Count, "A").End(xlUp).Row
Sheets("DATACUSTOMER").Range("A7:A" & CHGSOURCE).Find(What:=Me.TXT_FIRSTNAME.Value, LookIn:=xlValues, LookAt:=xlWhole).Activate
ACTIVE_CELL = ActiveCell.Row
Sheets("DATACUSTOMER").Range("A" & ACTIVE_CELL & ":U" & ACTIVE_CELL).Select

Me.MultiPage1.Value = 0
Me.STEP1.BackColor = RGB(16, 89, 199)
Me.STEP2.BackColor = RGB(16, 89, 199)
Me.STEP3.BackColor = RGB(16, 89, 199)
Me.STEP4.BackColor = RGB(16, 89, 199)
Me.S1.Caption = "OK"
Me.S2.Caption = "OK"
Me.S3.Caption = "OK"
Me.S4.Caption = "OK"

Exit Sub
EXCELVBA:
Call MsgBox("Double Click The Data", vbInformation, "Choose The Data")

End Sub

Private Sub Close_BTN_Click()
Unload Me
End Sub

Private Sub Excel_Return_Click()
    Me.MultiPage1.Value = 0
End Sub

Private Sub TXT_ADDRESS_Change()
If Me.TXT_ADDRESS.Value = "" _
Or Me.TXT_POSTCODE.Value = "" _
Or Me.TXT_CITY.Value = "" _
Or Me.TXT_PHONE.Value = "" _
Or Me.TXT_COUNTRY.Value = "" Then
Me.CMD_NEXT3.Enabled = False
Me.S3.Caption = 3
Else
Me.CMD_NEXT3.Enabled = True
Me.S3.Caption = "OK"
End If

End Sub

Private Sub TXT_BANK_Change()
If Me.TXT_BANK.Value = "" _
Or Me.TXT_CARDNUM.Value = "" _
Or Me.TXT_EXPDATE.Value = "" _
Or Me.TXT_CVC.Value = "" Then
Me.CMD_SUBMIT.Enabled = False
Me.S4.Caption = 4
Else
Me.CMD_SUBMIT.Enabled = True
Me.S4.Caption = "OK"
End If

End Sub

Private Sub TXT_BIRTHDAY_Change()
If Me.TXT_FIRSTNAME.Value = "" _
Or Me.TXT_LASTNAME.Value = "" _
Or Me.TXT_BIRTHDAY.Value = "" _
Or Me.CMB_GENDER.Value = "" _
Or Me.CMB_MARRIED.Value = "" _
Or Me.TXT_EMAIL.Value = "" Then
Me.CMD_NEXT1.Enabled = False
Me.S1.Caption = 1
Else
Me.CMD_NEXT1.Enabled = True
Me.S1.Caption = "OK"
End If

End Sub

Private Sub TXT_EMAIL_Click()
If Me.TXT_FIRSTNAME.Value = "" _
Or Me.TXT_LASTNAME.Value = "" _
Or Me.TXT_BIRTHDAY.Value = "" _
Or Me.CMB_GENDER.Value = "" _
Or Me.CMB_MARRIED.Value = "" _
Or Me.TXT_EMAIL.Value = "" Then
Me.CMD_NEXT1.Enabled = False
Me.S1.Caption = 1
Else
Me.CMD_NEXT1.Enabled = True
Me.S1.Caption = "OK"
End If
End Sub

Private Sub TXT_CARDNUM_Change()
If Me.TXT_BANK.Value = "" _
Or Me.TXT_CARDNUM.Value = "" _
Or Me.TXT_EXPDATE.Value = "" _
Or Me.TXT_CVC.Value = "" Then
Me.CMD_SUBMIT.Enabled = False
Me.S4.Caption = 4
Else
Me.CMD_SUBMIT.Enabled = True
Me.S4.Caption = "OK"
End If

End Sub

Private Sub TXT_CITY_Change()
If Me.TXT_ADDRESS.Value = "" _
Or Me.TXT_POSTCODE.Value = "" _
Or Me.TXT_CITY.Value = "" _
Or Me.TXT_PHONE.Value = "" _
Or Me.TXT_COUNTRY.Value = "" Then
Me.CMD_NEXT3.Enabled = False
Me.S3.Caption = 3
Else
Me.CMD_NEXT3.Enabled = True
Me.S3.Caption = "OK"
End If

End Sub

Private Sub TXT_COMPADDR_Change()
If Me.TXT_COMPNAME.Value = "" _
Or Me.TXT_COMPADDR.Value = "" _
Or Me.TXT_POSITION.Value = "" _
Or Me.CMB_STATUS.Value = "" _
Or Me.TXT_COMPPHONE.Value = "" _
Or Me.TXT_SALARY.Value = "" Then
Me.CMD_NEXT2.Enabled = False
Me.S2.Caption = 2
Else
Me.CMD_NEXT2.Enabled = True
Me.S2.Caption = "OK"
End If

End Sub

Private Sub TXT_COMPNAME_Change()
If Me.TXT_COMPNAME.Value = "" _
Or Me.TXT_COMPADDR.Value = "" _
Or Me.TXT_POSITION.Value = "" _
Or Me.CMB_STATUS.Value = "" _
Or Me.TXT_COMPPHONE.Value = "" _
Or Me.TXT_SALARY.Value = "" Then
Me.CMD_NEXT2.Enabled = False
Me.S2.Caption = 2
Else
Me.CMD_NEXT2.Enabled = True
Me.S2.Caption = "OK"
End If

End Sub

Private Sub TXT_COMPPHONE_Change()
If Me.TXT_COMPNAME.Value = "" _
Or Me.TXT_COMPADDR.Value = "" _
Or Me.TXT_POSITION.Value = "" _
Or Me.CMB_STATUS.Value = "" _
Or Me.TXT_COMPPHONE.Value = "" _
Or Me.TXT_SALARY.Value = "" Then
Me.CMD_NEXT2.Enabled = False
Me.S2.Caption = 2
Else
Me.CMD_NEXT2.Enabled = True
Me.S2.Caption = "OK"
End If

End Sub

Private Sub TXT_COUNTRY_Change()
If Me.TXT_ADDRESS.Value = "" _
Or Me.TXT_POSTCODE.Value = "" _
Or Me.TXT_CITY.Value = "" _
Or Me.TXT_PHONE.Value = "" _
Or Me.TXT_COUNTRY.Value = "" Then
Me.CMD_NEXT3.Enabled = False
Me.S3.Caption = 3
Else
Me.CMD_NEXT3.Enabled = True
Me.S3.Caption = "OK"
End If

End Sub

Private Sub TXT_CVC_Change()
If Me.TXT_BANK.Value = "" _
Or Me.TXT_CARDNUM.Value = "" _
Or Me.TXT_EXPDATE.Value = "" _
Or Me.TXT_CVC.Value = "" Then
Me.CMD_SUBMIT.Enabled = False
Me.S4.Caption = 4
Else
Me.CMD_SUBMIT.Enabled = True
Me.S4.Caption = "OK"
End If

End Sub

Private Sub TXT_EMAIL_Change()
If Me.TXT_FIRSTNAME.Value = "" _
Or Me.TXT_LASTNAME.Value = "" _
Or Me.TXT_BIRTHDAY.Value = "" _
Or Me.CMB_GENDER.Value = "" _
Or Me.CMB_MARRIED.Value = "" _
Or Me.TXT_EMAIL.Value = "" Then
Me.CMD_NEXT1.Enabled = False
Me.S1.Caption = 1
Else
Me.CMD_NEXT1.Enabled = True
Me.S1.Caption = "OK"
End If

End Sub

Private Sub TXT_EXPDATE_Change()
If Me.TXT_BANK.Value = "" _
Or Me.TXT_CARDNUM.Value = "" _
Or Me.TXT_EXPDATE.Value = "" _
Or Me.TXT_CVC.Value = "" Then
Me.CMD_SUBMIT.Enabled = False
Me.S4.Caption = 4
Else
Me.CMD_SUBMIT.Enabled = True
Me.S4.Caption = "OK"
End If

End Sub

Private Sub TXT_FIRSTNAME_Change()
If Me.TXT_FIRSTNAME.Value = "" _
Or Me.TXT_LASTNAME.Value = "" _
Or Me.TXT_BIRTHDAY.Value = "" _
Or Me.CMB_GENDER.Value = "" _
Or Me.CMB_MARRIED.Value = "" _
Or Me.TXT_EMAIL.Value = "" Then
Me.CMD_NEXT1.Enabled = False
Me.S1.Caption = 1
Else
Me.CMD_NEXT1.Enabled = True
Me.S1.Caption = "OK"
End If
End Sub

Private Sub TXT_LASTNAME_Change()
If Me.TXT_FIRSTNAME.Value = "" _
Or Me.TXT_LASTNAME.Value = "" _
Or Me.TXT_BIRTHDAY.Value = "" _
Or Me.CMB_GENDER.Value = "" _
Or Me.CMB_MARRIED.Value = "" _
Or Me.TXT_EMAIL.Value = "" Then
Me.CMD_NEXT1.Enabled = False
Me.S1.Caption = 1
Else
Me.CMD_NEXT1.Enabled = True
Me.S1.Caption = "OK"
End If
End Sub

Private Sub TXT_PHONE_Change()
If Me.TXT_ADDRESS.Value = "" _
Or Me.TXT_POSTCODE.Value = "" _
Or Me.TXT_CITY.Value = "" _
Or Me.TXT_PHONE.Value = "" _
Or Me.TXT_COUNTRY.Value = "" Then
Me.CMD_NEXT3.Enabled = False
Me.S3.Caption = 3
Else
Me.CMD_NEXT3.Enabled = True
Me.S3.Caption = "OK"
End If

End Sub

Private Sub TXT_POSITION_Change()
If Me.TXT_COMPNAME.Value = "" _
Or Me.TXT_COMPADDR.Value = "" _
Or Me.TXT_POSITION.Value = "" _
Or Me.CMB_STATUS.Value = "" _
Or Me.TXT_COMPPHONE.Value = "" _
Or Me.TXT_SALARY.Value = "" Then
Me.CMD_NEXT2.Enabled = False
Me.S2.Caption = 2
Else
Me.CMD_NEXT2.Enabled = True
Me.S2.Caption = "OK"
End If

End Sub

Private Sub TXT_POSTCODE_Change()
If Me.TXT_ADDRESS.Value = "" _
Or Me.TXT_POSTCODE.Value = "" _
Or Me.TXT_CITY.Value = "" _
Or Me.TXT_PHONE.Value = "" _
Or Me.TXT_COUNTRY.Value = "" Then
Me.CMD_NEXT3.Enabled = False
Me.S3.Caption = 3
Else
Me.CMD_NEXT3.Enabled = True
Me.S3.Caption = "OK"
End If

End Sub

Private Sub TXT_SALARY_Change()
If Me.TXT_COMPNAME.Value = "" _
Or Me.TXT_COMPADDR.Value = "" _
Or Me.TXT_POSITION.Value = "" _
Or Me.CMB_STATUS.Value = "" _
Or Me.TXT_COMPPHONE.Value = "" _
Or Me.TXT_SALARY.Value = "" Then
Me.CMD_NEXT2.Enabled = False
Me.S2.Caption = 2
Else
Me.CMD_NEXT2.Enabled = True
Me.S2.Caption = "OK"
End If

End Sub

Private Sub TXT_SEARCH_Change()

On Error GoTo SMTHWRONG
Set Cari_Data = Sheet1
Cari_Data.Range("W7").Value = "*" & Me.TXT_SEARCH.Value & "*"
Cari_Data.Range("A6").CurrentRegion.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:= _
Sheet1.Range("W6:W7"), CopyToRange:=Sheet2.Range("A1:U1"), Unique:=False
Me.DATA_TABLE.RowSource = Sheet2.Range("SEARCH_RESULT").Address(External:=True)
Exit Sub
SMTHWRONG:
Call MsgBox("Sorry, Data Not Found", vbInformation, "Find Data")

End Sub

Private Sub UserForm_Initialize()
HideTitleBar Me
Me.RIGHT_FRAME.Height = Me.Height
Me.MultiPage1.Value = 0
Me.S1.Caption = 1
Me.S2.Caption = 2
Me.S3.Caption = 3
Me.S4.Caption = 4

Me.TXT_BIRTHDAY.Value = Format(Me.TXT_BIRTHDAY.Value, "DD MM YYYY")

Me.RIGHT_FRAME.BackColor = RGB(14, 13, 38)
Me.STEP1.BackColor = RGB(16, 89, 199)
Me.STEP2.BackColor = RGB(196, 10, 88)
Me.STEP3.BackColor = RGB(196, 10, 88)
Me.STEP4.BackColor = RGB(196, 10, 88)

Me.CMD_NEXT1.Enabled = False
Me.CMD_NEXT2.Enabled = False
Me.CMD_NEXT3.Enabled = False
Me.CMD_SUBMIT.Enabled = False
Me.TXT_BANK.Enabled = False
With CMB_GENDER
.AddItem "Male"
.AddItem "Female"
End With

With CMB_MARRIED
.AddItem "Single"
.AddItem "Married"
End With

With CMB_STATUS
.AddItem "Full-Time"
.AddItem "Contract"
End With

On Error Resume Next
Me.DATA_TABLE.RowSource = Sheet1.Range("DATASOURCE").Address(External:=True)

End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        If Button = 1 Then
        m_sngDownX = X
        m_sngDownY = Y
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button And 1 Then
        Me.Left = Me.Left + (X - m_sngDownX)
        Me.Top = Me.Top + (Y - m_sngDownY)
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' User clicked the close button on the UserForm1 window
        CoverForm.Hide ' Hide the black cover UserForm
    End If
End Sub

Private Sub UserForm_Terminate()
    CoverForm.Hide ' Hide the black cover UserForm
End Sub


