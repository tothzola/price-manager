Attribute VB_Name = "modGlobal"
'@Folder("GlobalEntities")
Option Explicit

'SETTINGS
Public Const SIGN As String = "Demo Project"

'GENERAL SETTINGS
Public Const DATEFORMAT_BACKEND As String = "yyyy-mm-dd;@"
'Public Const DATEFORMAT_FRONTEND As String = "dd.mm.yyyy;@"
Public Const END_OF_THE_EARTH As String = "9999-12-31"
Public Const START_OF_THE_CENTURY As String = "2000-01-01"
Public Const CURRENCYFORMAT_FRONTEND As String = "Standard"

'NUMERICAL RANGES
Public Const INDEX_RECORDID_FIRST As Long = 1000000
Public Const INDEX_RECORDID_LAST As Long = 9999999

Public Const INDEX_CUSTOMERID_FIRST As Long = 399999
Public Const INDEX_CUSTOMERID_LAST As Long = 599999

Public Const INDEX_USERID_FIRST As Long = 100
Public Const INDEX_USERID_LAST As Long = 999

Public Const INDEX_MATERIALID_FIRST As Long = 49999999
Public Const INDEX_MATERIALID_LAST As Long = 59999999

Public Const MIN_PRICE_VALUE As Long = 0
Public Const MAX_PRICE_VALUE As Long = 999999

Public Const MIN_UNITOFPRICE_VALUE As Long = 0
Public Const MAX_UNITOFPRICE_VALUE As Long = 9999

'SEPERATORS FOR STRINGS MANIPULATION
Public Const SEPERATOR_LINE As String = "<LINE>"
Public Const SEPERATOR_ITEM As String = "<ITEM>"

'STATUSES
Public Const USERSTATUS_ACTIVE As String = "ACTIVE"
Public Const USERSTATUS_INACTIVE As String = "INACTIVE"

Public Const USERTYPE_CLIENT As String = "CLIENT"
Public Const USERTYPE_APPROVER As String = "APPROVER"

Public Const RECORDSTATUS_PENDING As String = "PENDING"
Public Const RECORDSTATUS_APPROVED As String = "APPROVED"
Public Const RECORDSTATUS_REJECTED As String = "REJECTED"

'COLORS
Public Const COLOR_OF_OKAY As Long = &H8000&        'GREEN TINT
Public Const COLOR_OF_NOT_OKAY As Long = &H2D04D2   'RED TINT

'Other Settings

Public Const BULLET_LISTITEM As String = ">>  "

Public Enum FormOperation
    OPERATION_NEW
    OPERATION_UPDATE
    OPERATION_DELETE
End Enum

Public Enum RepositoryType
    TYPE_EXCEL_NAMED_RANGE
    TYPE_SHAREPOINT_LIST
    TYPE_POSTGRESQL
    TYPE_ACCESS
End Enum

Public Enum UserApprovalStatus
    TYPE_PENDING
    TYPE_APPROVED
    TYPE_REJECTED
End Enum

Public Enum messageType
    TYPE_CRITICAL
    TYPE_INFORMATION
End Enum

Public Enum WarningType
    TYPE_NA
    TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied
    TYPE_NUMBERSONLY
    TYPE_STRINGSNOTMATCHED
    TYPE_WRONGPASSWORDPATTERN
    TYPE_FIXEDLENGTHSTRING
    TYPE_CUSTOM
End Enum

Public Enum ApplicationForms
    FORM_LOGIN = 1
    FORM_PASSWORDMANAGER
    FORM_USERMANAGER
    FORM_PRICEFORM
    FORM_DATAFORM
    FORM_EXPORTUTILITY
End Enum

Public Enum CRUDOperations
    CRUD_OPERATION_ADDNEW
    CRUD_OPERATION_UPDATE
    CRUD_OPERATION_DELETE
    CRUD_OPERATION_APPROVE
    CRUD_OPERATION_REJECT
End Enum

Public Enum DataContainer
    FOR_CLIENTHISTORY
    FOR_PENDINGAPPROVALS
    FOR_ALLHISTORY
End Enum

Public Enum DataTypes
    TYPE_DATE
    TYPE_CURRENCY
End Enum

Public Enum ValidationCheckTypes
    TYPE_STRINGMATCH
    TYPE_DATEBETWEENRANGE
End Enum

Private Enum Region
    US = 1                                 'United States
    UK = 44                                'United Kindom
    DE = 49                                'Germany
End Enum

Public Function DATEFORMAT_FRONTEND() As String
    Select Case GetRegion
    
        Case "US"
            DATEFORMAT_FRONTEND = "DD-MMM-YYYY"
            
        Case "EU"
            DATEFORMAT_FRONTEND = GetRegionalShortDate
            
        Case vbNullString
            DATEFORMAT_FRONTEND = "YYYY-MM-DD"
            
    End Select
End Function

Public Function GetRegion() As String

    Dim Code As String
    Select Case Application.International(xlCountrySetting)

    Case Region.US:
        Code = "US"

    Case Region.DE, Region.UK:
        Code = "EU"

    Case Else:
        Code = vbNullString

    End Select

    GetRegion = Code

End Function

Public Function EXPORTREPORT_CURRENCYFORMAT() As String
    EXPORTREPORT_CURRENCYFORMAT = "Comma"
End Function

Public Sub WaitForOneSecond()
    VBA.DoEvents
    Call Excel.Application.Wait(Now + TimeValue("00:00:01"))
End Sub

