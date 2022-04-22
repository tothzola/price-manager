Attribute VB_Name = "modDataSources"
Option Explicit

'Data Tables Structure

Public Enum PasswordManagerFields
    PM_CURRENT_PASSWORD = 1
    PM_NEW_PASSWORD
    PM_CONFIRM_NEW_PASSWORD
End Enum

Public Enum UsersTableFields
    COL_INDEX = 1
    COL_userID
    COL_userStatus
    COL_userType
    COL_userName
    COL_password
End Enum

Public Enum MainTableFields
    COL_MAIN_INDEX = 1
    COL_MAIN_recordID
    COL_MAIN_recordStatus
    COL_MAIN_statusChangeDate
    COL_MAIN_customerID
    COL_MAIN_materialID
    COL_MAIN_price
    COL_MAIN_currency
    COL_MAIN_unitOfPrice
    COL_MAIN_unitOfMeasure
    COL_MAIN_validFromDate
    COL_MAIN_validToDate
End Enum

'Data Sources Table Name

Public Const USERS_TABLE_NAME As String = "Table_Users"
Public Const MAIN_TABLE_NAME As String = "Table_Main"

'Elements of The Table

Public Function arrListOfFields_PASSWORD_MANAGER() As Variant
    arrListOfFields_PASSWORD_MANAGER = Array("Current Password", _
                                            "New Password", _
                                            "Confirm New Password")
End Function

Public Function arrListOfColumns_USERS_TABLE() As Variant
    arrListOfColumns_USERS_TABLE = Array("index", _
                                        "User ID", _
                                        "User Status", _
                                        "User Type", _
                                        "Username", _
                                        "Password")
End Function

Public Function arrListOfColumns_MAIN_Table() As Variant
    arrListOfColumns_MAIN_Table = Array("index", _
                                        "recordID", _
                                        "Record Status", _
                                        "Status Change Date", _
                                        "Customer ID", _
                                        "Material ID", _
                                        "Price", _
                                        "Currency", _
                                        "Unit Of Price", _
                                        "Unit Of Measure", _
                                        "Valid From Date", _
                                        "Valid To Date")
End Function

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
