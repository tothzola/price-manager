Attribute VB_Name = "modDataSources"
'@Folder("GlobalEntities")
Option Explicit

'Data Tables Structure

Public Enum TablesOfThisApplication
    TABLE_MAINRECORDS
    TABLE_USERS
End Enum

Public Enum PasswordManagerFields
    PM_CURRENT_PASSWORD = 1
    PM_NEW_PASSWORD
    PM_CONFIRM_NEW_PASSWORD
End Enum

Public Enum UsersTableFields
    COL_INDEX = 1
    COL_userId
    COL_userStatus
    COL_userType
    COL_userName
    COL_password
    COL_email
End Enum

Public Enum MainTableFields
    COL_MAIN_INDEX = 1
    COL_MAIN_recordID
    COL_MAIN_userID
    COL_MAIN_recordStatus
    COL_MAIN_statusChangeDate
    COL_MAIN_ConditionType
    COL_MAIN_SalesOrganization
    Col_Main_DistributionChannel
    COL_MAIN_customerID
    COL_MAIN_materialID
    COL_MAIN_price
    COL_MAIN_currency
    COL_MAIN_unitOfPrice
    COL_MAIN_unitOfMeasure
    COL_MAIN_validFromDate
    COL_MAIN_validToDate
End Enum

Public Enum ExportFormFields
    FIELD_FROMDATE = 1
    FIELD_TODATE
    FIELD_CUSTOMERID
    FIELD_USERID
    FIELD_RECORDSTATUS
End Enum

'Data Sources Table Name

Public Const USERS_TABLE_NAME As String = "Table_Users"
Public Const MAIN_TABLE_NAME As String = "Table_Main"

'Connection Strings

'ACCESS

Private Function DatabaseFilePath_Access() As String
    DatabaseFilePath_Access = ThisWorkbook.Path & Application.PathSeparator & "DatabaseAccess" _
                              & Application.PathSeparator & "PriceApprovalDatabase.accdb"
End Function

Public Function GetConnectionString(ByVal TypeOfRepository As RepositoryType) As String
    Select Case TypeOfRepository
        Case RepositoryType.TYPE_EXCEL_NAMED_RANGE
            GetConnectionString = vbNullString
        Case RepositoryType.TYPE_ACCESS
            GetConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DatabaseFilePath_Access & ";Persist Security Info=False;"
        Case RepositoryType.TYPE_POSTGRESQL
            'Attributes of PostgreSQL Connection String
            Const ServerAddress As String = "ec2-54-164-40-66.compute-1.amazonaws.com"
            Const PortNumber As Integer = 5432
            Const DatabaseName As String = "d34l6r35cqkfjd"
            Const UserName As String = "lxfsytloshyamh"
            Const Password As String = "6501b05101dba6b4ac0b2f32bbc18b81096f716bfc92343856d01cab6078153f"
            'Assembling attributes of PostgreSQL to form a ConnectionString
            GetConnectionString = "Driver={PostgreSQL ANSI}" & _
                                  ";Server=" & ServerAddress & _
                                  ";Port=" & PortNumber & _
                                  ";Database=" & DatabaseName & _
                                  ";Uid=" & UserName & _
                                  ";Pwd=" & Password & _
                                  ";sslmode=require;"
        Case RepositoryType.TYPE_SHAREPOINT_LIST
            GetConnectionString = vbNullString
    End Select
End Function

'Elements of The Table

Public Function arrListOfFields_PASSWORD_MANAGER() As Variant
    arrListOfFields_PASSWORD_MANAGER = Array("Current Password", _
                                             "New Password", _
                                             "Confirm New Password")
End Function

Public Function arrListOfColumns_USERS_TABLE() As Variant
    arrListOfColumns_USERS_TABLE = Array("Index", _
                                         "User_ID", _
                                         "User_Status", _
                                         "User_Type", _
                                         "Username", _
                                         "Password", _
                                         "Email")
End Function

Public Function arrListOfColumns_MAIN_Table() As Variant
    arrListOfColumns_MAIN_Table = Array("Index", _
                                        "Record_ID", _
                                        "User_ID", _
                                        "Record_Status", _
                                        "Status_Change_Date", _
                                        "Condition_Type", _
                                        "Sales_Organization", _
                                        "Distribution_Channel", _
                                        "Customer_ID", _
                                        "Material_ID", _
                                        "Price", _
                                        "CurrencyField", _
                                        "Unit_Of_Price", _
                                        "Unit_Of_Measure", _
                                        "Valid_From_Date", _
                                        "Valid_To_Date")
End Function

Public Function arrListOfFields_EXPORT_Form() As Variant
    arrListOfFields_EXPORT_Form = Array("Date From", _
                                        "Date To", _
                                        "Customer ID", _
                                        "User ID", _
                                        "Record Status")
End Function

Public Function arrHeaders_Export_Report() As Variant
    arrHeaders_Export_Report = Array("Index", _
                                     "Record ID", _
                                     "User", _
                                     "Record Status", _
                                     "Status Change Date", _
                                     "Condition Type", _
                                     "Sales Organization", _
                                     "Distribution Channel", _
                                     "Customer ID", _
                                     "Material ID", _
                                     "Price", _
                                     "CurrencyField", _
                                     "Unit Of Price", _
                                     "Unit Of Measure", _
                                     "Valid From Date", _
                                     "Valid To Date")
End Function

Public Function arrListOfColumnsMainTable() As Variant
    arrListOfColumnsMainTable = Array("", _
                                    "User_Name", _
                                    "Record_Status", _
                                    "Status_Change_Date", _
                                    "Customer_ID", _
                                    "Material_ID", _
                                    "Price", _
                                    "CurrencyField", _
                                    "Unit_Of_Price")
End Function

Public Function arrListOfColumnsMainTableFull() As Variant
    arrListOfColumnsMainTableFull = Array("", _
                                    "Index", _
                                    "Record_ID", _
                                    "User_Name", _
                                    "Record_Status", _
                                    "Status_Change_Date", _
                                    "Condition_Type", _
                                    "Sales_Organization", _
                                    "Distribution_Channel", _
                                    "Customer_ID", _
                                    "Material_ID", _
                                    "Price", _
                                    "CurrencyField", _
                                    "Unit_Of_Price", _
                                    "Unit_Of_Measure", _
                                    "Valid_From_Date", _
                                    "Valid_To_Date")
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

Public Function arrRecordStatusesList() As Variant
    arrRecordStatusesList = Array(vbNullString, "PENDING", "APPROVED", "REJECTED", "PROCESSED")
End Function

Public Function arrSalesOrganizationsList() As Variant
    arrSalesOrganizationsList = Array("2961")
End Function

Public Function arrDistributionChannelsList() As Variant
    arrDistributionChannelsList = Array("01", "GY", "HD")
End Function
