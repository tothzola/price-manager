Attribute VB_Name = "modGlobal"
Option Explicit

'Main Settings
Public Const DATEFORMAT_BACKEND As String = "DD-MMM-YYYY"
Public Const DATEFORMAT_FRONTEND As String = "DD.MM.YYYY"

Public Const INDEX_USERID_FIRST As Long = 399999
Public Const INDEX_USERID_LAST As Long = 599999

Public Const INDEX_MATERIALID_FIRST As Long = 49999999
Public Const INDEX_MATERIALID_LAST As Long = 59999999

Public Const MIN_PRICE_VALUE As Long = 0
Public Const MAX_PRICE_VALUE As Long = 999999

Public Const MIN_UNITOFPRICE_VALUE As Long = 0
Public Const MAX_UNITOFPRICE_VALUE As Long = 9999

Public Const SIGN As String = "Demo Project"

Public Const SEPERATOR_LINE As String = "<LINE>"
Public Const SEPERATOR_ITEM As String = "<ITEM>"

Public Const USERSTATUS_ACTIVE As String = "ACTIVE"
Public Const USERSTATUS_INACTIVE As String = "INACTIVE"

Public Const USERTYPE_CLIENT As String = "CLIENT"
Public Const USERTYPE_APPROVER As String = "APPROVER"

Public Const BULLET_LISTITEM As String = ">>  "

'COLORS
Public Const COLOR_OF_OKAY As Long = &H8000& 'GREEN TINT
Public Const COLOR_OF_NOT_OKAY As Long = &H2D04D2 'RED TINT

'Other Settings

Public Enum FormOperation
    OPERATION_NEW
    OPERATION_UPDATE
End Enum

Public Enum RepositoryType
    TYPE_EXCEL_NAMED_RANGE
    TYPE_SHAREPOINT_LIST
    TYPE_MYSQL
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
    TYPE_NUMBERSONLY
    TYPE_STRINGSNOTMATCHED
    TYPE_WRONGPASSWORDPATTERN
    TYPE_FIXEDLENGTHSTRING
End Enum

Public Enum ApplicationForms
    FORM_LOGIN = 1
    FORM_PASSWORDMANAGER
    FORM_USERMANAGER
    FORM_PRICEFORM
End Enum

Public Enum CRUDOperations
    CRUD_OPERATION_ADDNEW
    CRUD_OPERATION_UPDATE
    CRUD_OPERATION_DELETE
    CRUD_OPERATION_APPROVE
    CRUD_OPERATION_REJECT
End Enum


