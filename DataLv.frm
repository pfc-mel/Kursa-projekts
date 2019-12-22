VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataLv 
   Caption         =   "Data"
   ClientHeight    =   8745.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11490
   OleObjectBlob   =   "DataLv.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataLv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl_lang_EN_Click()

'-------------Objekti kuriem maina Caption------------

DataLv.btn_add.Caption = "Add"
DataLv.btn_computers.Caption = "Computers"
DataLv.btn_exit.Caption = "Log out"
DataLv.btn_remove.Caption = "Remove"
DataLv.btn_update.Caption = "Update"
DataLv.btn_users.Caption = "Users"

DataLv.Caption = "Data"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
DataLv.lbl_lang_EN.BackColor = "&H8000000D"
DataLv.lbl_lang_LV.BackColor = "&H80000010"
DataLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_LV_Click()

'-------------Objekti kuriem maina Caption------------

DataLv.btn_add.Caption = "Pievienot"
DataLv.btn_computers.Caption = "Datori"
DataLv.btn_exit.Caption = "Iziet"
DataLv.btn_remove.Caption = "Nonemt"
DataLv.btn_update.Caption = "Atjaunot"
DataLv.btn_users.Caption = "Lietotaji"

DataLv.Caption = "Data"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
DataLv.lbl_lang_EN.BackColor = "&H80000010"
DataLv.lbl_lang_LV.BackColor = "&H8000000D"
DataLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_RU_Click()

'-------------Objekti kuriem maina Caption------------

DataLv.btn_add.Caption = "???"
DataLv.btn_computers.Caption = "???"
DataLv.btn_exit.Caption = "???"
DataLv.btn_remove.Caption = "???"
DataLv.btn_update.Caption = "???"
DataLv.btn_users.Caption = "???"

DataLv.Caption = "???"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
DataLv.lbl_lang_EN.BackColor = "&H80000010"
DataLv.lbl_lang_LV.BackColor = "&H80000010"
DataLv.lbl_lang_RU.BackColor = "&H8000000D"
'------------------Bloka beigas----------------------

End Sub



Private Sub ExitButton_Click()
DataLv.Hide
LogInLv.Show
End Sub



Private Sub UsersButton_Click()
WorkFolderVar = Application.ActiveWorkbook.Path
    Dim currentbook As Workbook
    Dim source As Worksheet
    Set currentbook = Workbooks.Open(WorkFolderVar & "\Lietotaji.xlsm")
    Set source = currentbook.Sheets(1)
With DataLv.ListBox1
.TextAlign = 2
.ColumnHeads = True
.ColumnCount = 8
 .ColumnWidths = "50;50;50;50;50;50;50;50;"
ListBox1.List = source.Range("E4:L9999").Value
End With
currentbook.Save
currentbook.Close
End Sub


