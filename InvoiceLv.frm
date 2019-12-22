VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InvoiceLv 
   Caption         =   "Invoice"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "InvoiceLv.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InvoiceLv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl_lang_EN_Click()

'-------------Objekti kuriem maina Caption------------

InvoiceLv.lbl_com_addr.Caption = "Registered address: Riga, Jelgavas 25a"
InvoiceLv.lbl_company.Caption = "Datorveikals Ltd."
InvoiceLv.lbl_title.Caption = "Invoice - Bill"

InvoiceLv.btn_ok.Caption = "Ok"

InvoiceLv.Caption = "Invoice"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
InvoiceLv.lbl_lang_EN.BackColor = "&H8000000D"
InvoiceLv.lbl_lang_LV.BackColor = "&H80000010"
InvoiceLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_LV_Click()

'-------------Objekti kuriem maina Caption------------

InvoiceLv.lbl_com_addr.Caption = "Jur.adr : Riga, Jelgavas 25a"
InvoiceLv.lbl_company.Caption = "SIA Datorveikals"
InvoiceLv.lbl_title.Caption = "Pavadzime - Rekins"

InvoiceLv.btn_ok.Caption = "Ok"

InvoiceLv.Caption = "Invoice"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
InvoiceLv.lbl_lang_EN.BackColor = "&H80000010"
InvoiceLv.lbl_lang_LV.BackColor = "&H8000000D"
InvoiceLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_RU_Click()

'-------------Objekti kuriem maina Caption------------

InvoiceLv.lbl_com_addr.Caption = "???"
InvoiceLv.lbl_company.Caption = "???"
InvoiceLv.lbl_title.Caption = "???"

InvoiceLv.btn_ok.Caption = "???"

InvoiceLv.Caption = "Invoice"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
InvoiceLv.lbl_lang_EN.BackColor = "&H80000010"
InvoiceLv.lbl_lang_LV.BackColor = "&H80000010"
InvoiceLv.lbl_lang_RU.BackColor = "&H8000000D"
'------------------Bloka beigas----------------------

End Sub


