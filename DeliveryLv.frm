VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeliveryLv 
   Caption         =   "Delivery"
   ClientHeight    =   2460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6150
   OleObjectBlob   =   "DeliveryLv.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeliveryLv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl_lang_EN_Click()

'-------------Objekti kuriem maina Caption------------

DeliveryLv.btn_cont.Caption = "Continue"

DeliveryLv.lbl_address.Caption = "Address"
DeliveryLv.lbl_delivery.Caption = "Delivery"
DeliveryLv.lbl_time.Caption = "Most suitable time"

DeliveryLv.Caption = "Delivery"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
DeliveryLv.lbl_lang_EN.BackColor = "&H8000000D"
DeliveryLv.lbl_lang_LV.BackColor = "&H80000010"
DeliveryLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_LV_Click()

'-------------Objekti kuriem maina Caption------------

DeliveryLv.btn_cont.Caption = "Talak"

DeliveryLv.lbl_address.Caption = "Adrese"
DeliveryLv.lbl_delivery.Caption = "Piegades noformesana"
DeliveryLv.lbl_time.Caption = "Piemerotakais laiks"

DeliveryLv.Caption = "Piegade"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
DeliveryLv.lbl_lang_EN.BackColor = "&H80000010"
DeliveryLv.lbl_lang_LV.BackColor = "&H8000000D"
DeliveryLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_RU_Click()

'-------------Objekti kuriem maina Caption------------

DeliveryLv.btn_cont.Caption = "???"

DeliveryLv.lbl_address.Caption = "???"
DeliveryLv.lbl_delivery.Caption = "???"
DeliveryLv.lbl_time.Caption = "???"

DeliveryLv.Caption = "???"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
DeliveryLv.lbl_lang_EN.BackColor = "&H80000010"
DeliveryLv.lbl_lang_LV.BackColor = "&H80000010"
DeliveryLv.lbl_lang_RU.BackColor = "&H8000000D"
'------------------Bloka beigas----------------------

End Sub
