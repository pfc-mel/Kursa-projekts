VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ErrorLV 
   Caption         =   "ErrorLv"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ErrorLV.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ErrorLV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_ok_Click()
ErrorLV.Hide
LogInLv.Show
End Sub


Private Sub lbl_lang_EN_Click()

'-------------Objekti kuriem maina Caption------------

ErrorLV.lbl_error_text.Caption = "Error, wrong username or password"
ErrorLV.btn_ok.Caption = "Ok"

ErrorLV.Caption = "Error"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
ErrorLV.lbl_lang_EN.BackColor = "&H8000000D"
ErrorLV.lbl_lang_LV.BackColor = "&H80000010"
ErrorLV.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_LV_Click()

'-------------Objekti kuriem maina Caption------------

ErrorLV.lbl_error_text.Caption = "Kluda, nepareizs lietotajvards un/vai parole"
ErrorLV.btn_ok.Caption = "Ok"

ErrorLV.Caption = "Kluda"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
ErrorLV.lbl_lang_EN.BackColor = "&H80000010"
ErrorLV.lbl_lang_LV.BackColor = "&H8000000D"
ErrorLV.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_RU_Click()

'-------------Objekti kuriem maina Caption------------

ErrorLV.lbl_error_text.Caption = "???"
ErrorLV.btn_ok.Caption = "???"

ErrorLV.Caption = "???"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
ErrorLV.lbl_lang_EN.BackColor = "&H80000010"
ErrorLV.lbl_lang_LV.BackColor = "&H80000010"
ErrorLV.lbl_lang_RU.BackColor = "&H8000000D"
'------------------Bloka beigas----------------------

End Sub
