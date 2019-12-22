VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ErrorRegLv 
   Caption         =   "Kluuda"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ErrorRegLv.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ErrorRegLv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
ErrorRegLv.Hide
RegisterLv.Show
End Sub

Private Sub lbl_lang_EN_Click()

'-------------Objekti kuriem maina Caption------------

ErrorRegLv.lbl_error_text.Caption = "Error, username and password are required fields!"
ErrorRegLv.btn_ok.Caption = "Ok"

ErrorRegLv.Caption = "Error"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
ErrorRegLv.lbl_lang_EN.BackColor = "&H8000000D"
ErrorRegLv.lbl_lang_LV.BackColor = "&H80000010"
ErrorRegLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_LV_Click()

'-------------Objekti kuriem maina Caption------------

ErrorRegLv.lbl_error_text.Caption = "Kluda , lietotajvards un parole ir obligati janorada!!"
ErrorRegLv.btn_ok.Caption = "Ok"

ErrorRegLv.Caption = "Kluda"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
ErrorRegLv.lbl_lang_EN.BackColor = "&H80000010"
ErrorRegLv.lbl_lang_LV.BackColor = "&H8000000D"
ErrorRegLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_RU_Click()

'-------------Objekti kuriem maina Caption------------

ErrorRegLv.lbl_error_text.Caption = "???"
ErrorRegLv.btn_ok.Caption = "???"

ErrorRegLv.Caption = "???"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
ErrorRegLv.lbl_lang_EN.BackColor = "&H80000010"
ErrorRegLv.lbl_lang_LV.BackColor = "&H80000010"
ErrorRegLv.lbl_lang_RU.BackColor = "&H8000000D"
'------------------Bloka beigas----------------------

End Sub
