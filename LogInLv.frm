VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LogInLv 
   Caption         =   "LogIn"
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4590
   OleObjectBlob   =   "LogInLv.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LogInLv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_register_Click()
LogInLv.Hide
RegisterLv.Show
End Sub

Private Sub lbl_lang_EN_Click()

'-------------Objekti kuriem maina Caption------------

LogInLv.lbl_password.Caption = "Password"
LogInLv.lbl_username.Caption = "Username"

LogInLv.btn_login.Caption = "Log In"
LogInLv.btn_register.Caption = "Register"

LogInLv.Caption = "Log In"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
LogInLv.lbl_lang_EN.BackColor = "&H8000000D"
LogInLv.lbl_lang_LV.BackColor = "&H80000010"
LogInLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_LV_Click()

'-------------Objekti kuriem maina Caption------------

LogInLv.lbl_password.Caption = "Parole"
LogInLv.lbl_username.Caption = "Lietotajvards"

LogInLv.btn_login.Caption = "Pieslegties"
LogInLv.btn_register.Caption = "Registreties"

LogInLv.Caption = "Pieslegties"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
LogInLv.lbl_lang_EN.BackColor = "&H80000010"
LogInLv.lbl_lang_LV.BackColor = "&H8000000D"
LogInLv.lbl_lang_RU.BackColor = "&H80000010"
'------------------Bloka beigas----------------------

End Sub

Private Sub lbl_lang_RU_Click()

'-------------Objekti kuriem maina Caption------------

LogInLv.lbl_password.Caption = "???"
LogInLv.lbl_username.Caption = "???"

LogInLv.btn_login.Caption = "???"
LogInLv.btn_register.Caption = "???"

LogInLv.Caption = "???"
'------------------Bloka beigas----------------------

'----------------Label fona krasas-------------------
LogInLv.lbl_lang_EN.BackColor = "&H80000010"
LogInLv.lbl_lang_LV.BackColor = "&H80000010"
LogInLv.lbl_lang_RU.BackColor = "&H8000000D"
'------------------Bloka beigas----------------------

End Sub
Private Sub btn_login_Click()
    WorkFolderVar = Application.ActiveWorkbook.Path
    Dim currentbook As Workbook
    Dim source As Worksheet
    Set currentbook = Workbooks.Open(WorkFolderVar & "\Lietotaji.xlsm")
    Set source = currentbook.Sheets(1)
    LogInLv.Hide
    If LogInLv.txtBox_password = "admin" And LogInLv.txtbox_login_username = "admin" Then
    currentbook.Close
    DataLv.Show
    ElseIf LogInLv.txtBox_password = "" Or LogInLv.txtbox_login_username = "" Then
    currentbook.Close
    ErrorLV.Show
    ElseIf LogInLv.txtbox_login_username = "admin" And LogInLv.txtBox_password <> "admin" Then
    currentbook.Close
    ErrorLV.Show
    Else
    source.Range("f4").Select
    Do While ActiveCell.Value <> LogInLv.txtbox_login_username
    ActiveCell.Offset(1, 0).Select
    Loop
    

 
   Dim cLetter As String
   Dim cPassword As String
   Dim i As Integer

   If Not IsNull(ActiveCell.Offset(0, 6).Value) Then
     
      
      cEncryptedPassword = Trim(ActiveCell.Offset(0, 6).Value)
      
      cPassword = ""
      For i = 1 To Len(cEncryptedPassword)
          
          cLetter = Mid(cEncryptedPassword, i, 1)
          cPassword = cPassword + Chr(Asc(cLetter))
      Next i
   End If
   
   ActiveCell.Offset(0, 6).Value = cPassword
   MsgBox cPassword
    If (LogInLv.txtbox_login_username = ActiveCell.Value) And (LogInLv.txtBox_password = ActiveCell.Offset(0, 6).Value) Then
    currentbook.Save
    currentbook.Close
    CatalogLv.Show
    Else: currentbook.Save
    currentbook.Close
    ErrorLV.Show
    End If
      
    End If
End Sub



Private Sub UserForm_Click()

End Sub
