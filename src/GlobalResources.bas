Attribute VB_Name = "GlobalResources"
'@Folder("Constants")
Option Explicit
Option Private Module

#If VBA7 Then
    Private Declare PtrSafe Function EnumDateFormatsA Lib "Kernel32" (ByVal lpDateFmtEnumProc As LongPtr, ByVal LCID As Long, ByVal dwFlags As Long) As Boolean
    Private Declare PtrSafe Function lstrlenA Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
#Else
    Private Declare Function EnumDateFormatsA Lib "Kernel32" (ByVal lpDateFmtEnumProc As Long, ByVal LCID As Long, ByVal dwFlags As Long) As Boolean
    Private Declare Function lstrlenA Lib "kernel32.dll" (ByVal lpString As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

'SETTINGS
Public Const SIGN As String = "Demo Project"

'GENERAL SETTINGS
Public Const DATEFORMAT_BACKEND As String = "yyyy-mm-dd;@"
Public Const END_OF_THE_EARTH As String = "9999-12-31"
Public Const START_OF_THE_CENTURY As String = "2000-01-01"

'NUMERICAL RANGES
Public Const INDEX_RECORDID_FIRST As Long = 1000000

Public Const INDEX_CUSTOMERID_FIRST As Long = 399999
Public Const INDEX_CUSTOMERID_LAST As Long = 599999

Public Const INDEX_USERID_FIRST As Long = 100

Public Const INDEX_MATERIALID_FIRST As Long = 49999999
Public Const INDEX_MATERIALID_LAST As Long = 59999999

Public Const MIN_PRICE_VALUE As Long = 0
Public Const MAX_PRICE_VALUE As Long = 999999

Public Const MIN_UNITOFPRICE_VALUE As Long = 0
Public Const MAX_UNITOFPRICE_VALUE As Long = 9999

'SEPERATORS FOR STRINGS MANIPULATION
Public Const SEPERATOR_ITEM As String = "<ITEM>"

'STATUSES
Public Const USERSTATUS_ACTIVE As String = "ACTIVE"
Public Const USERSTATUS_INACTIVE As String = "INACTIVE"

Public Const USERTYPE_CLIENT As String = "CLIENT"
Public Const USERTYPE_APPROVER As String = "APPROVER"

Public Const RECORDSTATUS_PENDING As String = "PENDING"
Public Const RECORDSTATUS_APPROVED As String = "APPROVED"
Public Const RECORDSTATUS_REJECTED As String = "REJECTED"
Public Const RECORDSTATUS_PROCESSED As String = "PROCESSED"

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

Public Enum UserApprovalStatus
    TYPE_PENDING
    TYPE_APPROVED
    TYPE_REJECTED
    TYPE_PROCESSED
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

' this enum is based on the "Date Flags for GetDateFormat." from WinNls.h
Public Enum DateFormat
    ShortDate = &H1                              ' use short date picture
    LongDate = &H2                               ' use long date picture
    YearMonth = &H8                              ' use year month picture
End Enum

Public Enum SettingContext
    UserDefault = &H400
    SystemDefault = &H800
End Enum

Private m_dateFormat As String

#If VBA7 Then
Private Function StringFromPointer(ByVal pointerToString As LongPtr) As String
#Else
Private Function StringFromPointer(ByVal pointerToString As Long) As String
#End If

    Dim tmpBuffer()    As Byte
    Dim byteCount      As Long
    Dim retVal         As String
 
    byteCount = lstrlenA(pointerToString)
    
    If byteCount > 0 Then
  
        ' Resize the buffer as required
        ReDim tmpBuffer(0 To byteCount - 1) As Byte
        
        ' Copy the bytes from pointerToString to out tmpBuffer
        CopyMemory VarPtr(tmpBuffer(0)), pointerToString, byteCount
    End If
 
    ' Convert Buffer to string
    retVal = StrConv(tmpBuffer, vbUnicode)
    
    StringFromPointer = retVal

End Function

#If VBA7 Then
Private Function EnumDateFormatsProc(ByVal lpDateFormatString As LongPtr) As Boolean
#Else
Private Function EnumDateFormatsProc(ByVal lpDateFormatString As Long) As Boolean
#End If
    m_dateFormat = StringFromPointer(lpDateFormatString)
    EnumDateFormatsProc = True
End Function

Public Function GetDateFormat(Optional ByVal Format As DateFormat = DateFormat.ShortDate, Optional ByVal context As SettingContext = SettingContext.UserDefault) As String

    Dim apiRetVal As Boolean
    m_dateFormat = vbNullString
    
    apiRetVal = EnumDateFormatsA(AddressOf EnumDateFormatsProc, context, Format)
    
    If apiRetVal Then
        GetDateFormat = m_dateFormat
    End If

End Function

Public Function EXPORTREPORT_CURRENCYFORMAT() As String
    EXPORTREPORT_CURRENCYFORMAT = "Comma"
End Function

Public Sub WaitForOneSecond()
    VBA.DoEvents
    Excel.Application.Wait (VBA.Now + VBA.TimeValue("00:00:01"))
End Sub

