VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OrderLv 
   Caption         =   "Order"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8220.001
   OleObjectBlob   =   "OrderLv.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OrderLv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl_lang_EN_Click()

'-------------Objekti kuriem maina Caption------------

OrderLv.lbl_create_order.Caption = "Create order"
OrderLv.lbl_deliv.Caption = "Do you need delivery?"
OrderLv.lbl_Item_list.Caption = "Item list"
OrderLv.lbl_no.Caption = "No"
OrderLv.lbl_pay_method.Caption = "Paying method"
OrderLv.lbl_yes.Caption = "Yes"

OrderLv.Caption = "Order"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
OrderLv.lbl_lang_EN.BackColor = "&H8000000D"
OrderLv.lbl_lang_LV.BackColor = "&H80000010"
OrderLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_LV_Click()

'-------------Objekti kuriem maina Caption------------

OrderLv.lbl_create_order.Caption = "Pasutijuma Noformesana"
OrderLv.lbl_deliv.Caption = "Vai bus vajadziga piegade?"
OrderLv.lbl_Item_list.Caption = "Precu saraksts"
OrderLv.lbl_no.Caption = "Ne"
OrderLv.lbl_pay_method.Caption = "Maks. Veids"
OrderLv.lbl_yes.Caption = "Ja"

OrderLv.Caption = "Pasutijums"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
OrderLv.lbl_lang_EN.BackColor = "&H80000010"
OrderLv.lbl_lang_LV.BackColor = "&H8000000D"
OrderLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_RU_Click()

'-------------Objekti kuriem maina Caption------------

OrderLv.lbl_create_order.Caption = "???"
OrderLv.lbl_deliv.Caption = "???"
OrderLv.lbl_Item_list.Caption = "???"
OrderLv.lbl_no.Caption = "???"
OrderLv.lbl_pay_method.Caption = "???"
OrderLv.lbl_yes.Caption = "???"

OrderLv.Caption = "???"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
OrderLv.lbl_lang_EN.BackColor = "&H80000010"
OrderLv.lbl_lang_LV.BackColor = "&H80000010"
OrderLv.lbl_lang_RU.BackColor = "&H8000000D"
'------------------Bloka beigas----------------------

End Sub
