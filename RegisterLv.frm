VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegisterLv 
   Caption         =   "Registracija"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4605
   OleObjectBlob   =   "RegisterLv.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegisterLv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_back_Click()
RegisterLv.Hide
LogInLv.Show

End Sub

Private Sub CommandButton1_Click()
RegisterLv.Hide
RegisterEn.Show
End Sub


Private Sub btn_register_Click()
    RegisterLv.Hide
    WorkFolderVar = Application.ActiveWorkbook.Path
    Dim currentbook As Workbook
    Dim source As Worksheet
    Set currentbook = Workbooks.Open(WorkFolderVar & "\Lietotaji.xlsm")
    Set source = currentbook.Sheets(1)
    source.Range("f4").Select
    Do While ActiveCell.Value <> ""
    If ActiveCell.Value <> "" Then
        ActiveCell.Offset(1, 0).Select
    End If
    Loop
    If RegisterLv.txtbox_username = "" Or RegisterLv.txtbox_pass = "" Then
    currentbook.Close
    ErrorRegLv.Show
    Else
    ActiveCell.Value = txtbox_username.Value
    ActiveCell.Offset(0, -1).Value = ActiveCell.Offset(-1, -1).Value + 1
    ActiveCell.Offset(0, 1).Value = txtbox_name.Value
    ActiveCell.Offset(0, 2).Value = txtbox_lname.Value
    ActiveCell.Offset(0, 3).Value = txtbox_email.Value
    ActiveCell.Offset(0, 4).Value = txtbox_num.Value
    ActiveCell.Offset(0, 5).Value = txtbox_birth.Value
    ActiveCell.Offset(0, 6).Value = txtbox_pass.Value
    currentbook.Save
    currentbook.Close
    LogInLv.Show
    End If
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub lbl_lang_EN_Click()

'-------------Objekti kuriem maina Caption------------

RegisterLv.lbl_birth.Caption = "Date of birth"
RegisterLv.lbl_email.Caption = "E-mail"
RegisterLv.lbl_lastname.Caption = "Last name"
RegisterLv.lbl_name.Caption = "Name"
RegisterLv.lbl_num.Caption = "Phone number"
RegisterLv.lbl_pass.Caption = "Password"
RegisterLv.lbl_username.Caption = "Username"

RegisterLv.btn_back.Caption = "Back"
RegisterLv.btn_register.Caption = "Register"

RegisterLv.Caption = "Register"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
RegisterLv.lbl_lang_EN.BackColor = "&H8000000D"
RegisterLv.lbl_lang_LV.BackColor = "&H80000010"
RegisterLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_LV_Click()

'-------------Objekti kuriem maina Caption------------

RegisterLv.lbl_birth.Caption = "Dzim. datums"
RegisterLv.lbl_email.Caption = "E-pasts"
RegisterLv.lbl_lastname.Caption = "Uzvards"
RegisterLv.lbl_name.Caption = "Vards"
RegisterLv.lbl_num.Caption = "Tel.Nr."
RegisterLv.lbl_pass.Caption = "Parole"
RegisterLv.lbl_username.Caption = "Lietotajvards"

RegisterLv.btn_back.Caption = "Atpakal"
RegisterLv.btn_register.Caption = "Registreties"

RegisterLv.Caption = "Registracija"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
RegisterLv.lbl_lang_EN.BackColor = "&H80000010"
RegisterLv.lbl_lang_LV.BackColor = "&H8000000D"
RegisterLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_RU_Click()

'-------------Objekti kuriem maina Caption------------

RegisterLv.lbl_birth.Caption = "???"
RegisterLv.lbl_email.Caption = "???"
RegisterLv.lbl_lastname.Caption = "???"
RegisterLv.lbl_name.Caption = "???"
RegisterLv.lbl_num.Caption = "???"
RegisterLv.lbl_pass.Caption = "???"
RegisterLv.lbl_username.Caption = "???"

RegisterLv.btn_back.Caption = "???"
RegisterLv.btn_register.Caption = "???"

RegisterLv.Caption = "???"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
RegisterLv.lbl_lang_EN.BackColor = "&H80000010"
RegisterLv.lbl_lang_LV.BackColor = "&H80000010"
RegisterLv.lbl_lang_RU.BackColor = "&H8000000D"
'------------------Bloka beigas----------------------

End Sub
