Attribute VB_Name = "modDataSources"
Option Explicit

'Data Tables Structure

Public Enum UsersTableFields
    COL_index = 1
    COL_userID
    COL_userStatus
    COL_userType
    COL_userName
    COL_password
End Enum

'Data Sources Table Name

Public Const USERS_TABLE_NAME As String = "Table_Users"
Public Const MAIN_TABLE_NAME As String = "Table_Main"

'following functions returns array objects that will be used as dataSource for the comboboxes.
'If the combobox uses dynamic data then ofcourse it can be managed from the database application
'but generally, some lists never evolve with the time and if situation occures then
'we can eventually give update as well.

Public Function arrListOfUnitOfMeasure() As Variant
   arrListOfUnitOfMeasure = Array("KAR", "RO", "ST", "KG", "LM", "M2")
End Function

Public Function arrListofCurrencies() As Variant
    arrListofCurrencies = Array("EUR", "USD", "GBP", "PLN")
End Function

Public Function arrListofTypesOfUser() As Variant
    arrListofTypesOfUser = Array("CLIENT", "APPROVER")
End Function

Public Function arrListofStatusOfUser() As Variant
    arrListofStatusOfUser = Array("ACTIVE", "INACTIVE")
End Function
