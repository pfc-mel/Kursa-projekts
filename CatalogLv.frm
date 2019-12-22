VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CatalogLv 
   Caption         =   "Katalogs"
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13830
   OleObjectBlob   =   "CatalogLv.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CatalogLv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_logout_Click()
CatalogLv.Hide
LogInLv.Show
End Sub

Private Sub lbl_lang_EN_Click()

'-------------Objekti kuriem maina Caption------------

CatalogLv.lbl_computer_no.Caption = "Computer No"
CatalogLv.lbl_SortBy.Caption = "Sort By"

CatalogLv.btn_add_to_Cart.Caption = "Add to cart"
CatalogLv.btn_Cart.Caption = "Cart"
CatalogLv.btn_logout.Caption = "Log out"

CatalogLv.Caption = "Catalog"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
CatalogLv.lbl_lang_EN.BackColor = "&H8000000D"
CatalogLv.lbl_lang_LV.BackColor = "&H80000010"
CatalogLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_LV_Click()

'-------------Objekti kuriem maina Caption------------

CatalogLv.lbl_computer_no.Caption = "Datora Nr.p.k."
CatalogLv.lbl_SortBy.Caption = "Atlasit pec"

CatalogLv.btn_add_to_Cart.Caption = "Pievienot grozam"
CatalogLv.btn_Cart.Caption = "Grozs"
CatalogLv.btn_logout.Caption = "Iziet"

CatalogLv.Caption = "Katalogs"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
CatalogLv.lbl_lang_EN.BackColor = "&H80000010"
CatalogLv.lbl_lang_LV.BackColor = "&H8000000D"
CatalogLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_RU_Click()

'-------------Objekti kuriem maina Caption------------

CatalogLv.lbl_computer_no.Caption = "???"
CatalogLv.lbl_SortBy.Caption = "???"

CatalogLv.btn_add_to_Cart.Caption = "???"
CatalogLv.btn_Cart.Caption = "???"
CatalogLv.btn_logout.Caption = "???"

CatalogLv.Caption = "???"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
CatalogLv.lbl_lang_EN.BackColor = "&H80000010"
CatalogLv.lbl_lang_LV.BackColor = "&H80000010"
CatalogLv.lbl_lang_RU.BackColor = "&H8000000D"
'------------------Bloka beigas----------------------

End Sub

Private Sub UserForm_Click()

End Sub
