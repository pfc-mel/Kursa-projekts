VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ThanksLv 
   Caption         =   "Paldies"
   ClientHeight    =   1905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ThanksLv.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ThanksLv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl_lang_EN_Click()

'-------------Objekti kuriem maina Caption------------

ThanksLv.lbl_orderNO.Caption = "Your order No is"
ThanksLv.lbl_payment.Caption = "You can pay in shop when items are there"
ThanksLv.lbl_sms.Caption = "You will receive a message (SMS)."
ThanksLv.lbl_thanks.Caption = "Thank you for ordering!"

ThanksLv.Caption = "Thank you"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
ThanksLv.lbl_lang_EN.BackColor = "&H8000000D"
ThanksLv.lbl_lang_LV.BackColor = "&H80000010"
ThanksLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_LV_Click()

'-------------Objekti kuriem maina Caption------------

ThanksLv.lbl_orderNO.Caption = "Jusu pasutijuma numurs ir"
ThanksLv.lbl_payment.Caption = "Jus varesiet veikt apmaksu veikala, kad prece bus uz vietas"
ThanksLv.lbl_sms.Caption = "Jus sanemsiet SMS"
ThanksLv.lbl_thanks.Caption = "Paldies par pasutijumu!"

ThanksLv.Caption = "Paldies"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
ThanksLv.lbl_lang_EN.BackColor = "&H80000010"
ThanksLv.lbl_lang_LV.BackColor = "&H8000000D"
ThanksLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_RU_Click()

'-------------Objekti kuriem maina Caption------------

ThanksLv.lbl_orderNO.Caption = "???"
ThanksLv.lbl_payment.Caption = "???"
ThanksLv.lbl_sms.Caption = "???"
ThanksLv.lbl_thanks.Caption = "???"

ThanksLv.Caption = "???"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
ThanksLv.lbl_lang_EN.BackColor = "&H80000010"
ThanksLv.lbl_lang_LV.BackColor = "&H80000010"
ThanksLv.lbl_lang_RU.BackColor = "&H8000000D"
'------------------Bloka beigas----------------------

End Sub
