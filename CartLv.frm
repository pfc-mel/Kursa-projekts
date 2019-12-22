VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CartLv 
   Caption         =   "Grozs"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   OleObjectBlob   =   "CartLv.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CartLv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl_lang_EN_Click()
'-------------Objekti kuriem maina Caption------------
CartLv.lbl_itemNo.Caption = "Item No"
CartLv.lbl_itemsSelect.Caption = "Selected items"
CartLv.btn_back.Caption = "Back"
CartLv.btn_DelteItem.Caption = "Delete"
CartLv.btn_order.Caption = "Finish order"
CartLv.Caption = "Cart"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
CartLv.lbl_lang_EN.BackColor = "&H8000000D"
CartLv.lbl_lang_LV.BackColor = "&H80000010"
CartLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_LV_Click()
'-------------Objekti kuriem maina Caption------------
CartLv.lbl_itemNo.Caption = "Preces nr.p.k."
CartLv.lbl_itemsSelect.Caption = "Izveletas preces"
CartLv.btn_back.Caption = "Atpakal"
CartLv.btn_DelteItem.Caption = "Dzest"
CartLv.btn_order.Caption = "Noformet pasutijumu"
CartLv.Caption = "Grozs"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
CartLv.lbl_lang_EN.BackColor = "&H80000010"
CartLv.lbl_lang_LV.BackColor = "&H8000000D"
CartLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub


Private Sub lbl_lang_RU_Click()
'-------------Objekti kuriem maina Caption------------
CartLv.lbl_itemNo.Caption = "???"
CartLv.lbl_itemsSelect.Caption = "???"
CartLv.btn_back.Caption = "???"
CartLv.btn_DelteItem.Caption = "???"
CartLv.btn_order.Caption = "???"
CartLv.Caption = "???"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
CartLv.lbl_lang_EN.BackColor = "&H80000010"
CartLv.lbl_lang_LV.BackColor = "&H80000010"
CartLv.lbl_lang_RU.BackColor = "&H8000000D"
'------------------Bloka beigas----------------------


End Sub

Private Sub UserForm_Click()

End Sub
