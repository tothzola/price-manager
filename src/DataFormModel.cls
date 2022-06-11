VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataFormModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "DataForm"
Option Explicit

Private Type TDataFormModelComponents
    'Validation Components
    Validator As ValidationServices
    'Tables
    dataTable As Variant
    usersTable As Variant
    'List Single Columns
    DataColumnsList As Variant
    ValuesList As Variant
    'Input fields
    selectedColumn As Variant
    selectedValue As Variant
    'Fields
    index As Long
    RecordID As String
    userID As String
    recordStatus As String
    statusChangeDate As String
    conditionType As String
    salesOrganization As String
    distributionChannel As String
    customerID As String
    materialID As String
    price As String
    currencyType As String
    unitOfPrice As String
    unitOfMeasure As String
    validFromDate  As String
    validToDate As String
    'Attributes
    ActiveDataContainer As Long
    IsApprover As Boolean
End Type

Private this As TDataFormModelComponents

'-------------------------------------------------------------------------

Public Property Get Validator() As ValidationServices
    Set Validator = this.Validator
End Property

Public Property Set Validator(ByVal vNewValue As ValidationServices)
    Set this.Validator = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get dataTable() As Variant
    dataTable = this.dataTable
End Property

Public Property Let dataTable(ByVal vNewValue As Variant)
    this.dataTable = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get usersTable() As Variant
    usersTable = this.usersTable
End Property

Public Property Let usersTable(ByVal vNewValue As Variant)
    this.usersTable = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get DataColumnsList() As Variant
    DataColumnsList = this.DataColumnsList
End Property

Public Property Let DataColumnsList(ByVal vNewValue As Variant)
    this.DataColumnsList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ValuesList() As Variant
    ValuesList = this.ValuesList
End Property

Public Property Let ValuesList(ByVal vNewValue As Variant)
    this.ValuesList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get selectedValue() As Variant
    Select Case Me.selectedColumn
    Case MainTableFields.COL_MAIN_statusChangeDate
        If this.selectedValue <> vbNullString Then selectedValue = GetMethod.GetBackendDateFromDateString(this.selectedValue)
    Case Else
        selectedValue = this.selectedValue
    End Select
End Property

Public Property Let selectedValue(ByVal vNewValue As Variant)
    this.selectedValue = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get selectedColumn() As Variant
    selectedColumn = GetMethod.GetTargetColumnIndex(this.selectedColumn, modDataSources.arrListOfColumnsMainTableFull)
End Property

Public Property Let selectedColumn(ByVal vNewValue As Variant)
    this.selectedColumn = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get index() As Long
    index = this.index
End Property

Public Property Let index(ByVal vNewValue As Long)
    this.index = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get RecordID() As String
    RecordID = this.RecordID
End Property

Public Property Let RecordID(ByVal vNewValue As String)
    this.RecordID = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userID() As String
    userID = this.userID
End Property

Public Property Let userID(ByVal vNewValue As String)
    this.userID = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get recordStatus() As String
    recordStatus = this.recordStatus
End Property

Public Property Let recordStatus(ByVal vNewValue As String)
    this.recordStatus = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get statusChangeDate() As String
    statusChangeDate = this.statusChangeDate
End Property

Public Property Let statusChangeDate(ByVal vNewValue As String)
    this.statusChangeDate = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get distributionChannel() As String
    distributionChannel = this.distributionChannel
End Property

Public Property Let distributionChannel(ByVal vNewValue As String)
    this.distributionChannel = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get salesOrganization() As String
    salesOrganization = this.salesOrganization
End Property

Public Property Let salesOrganization(ByVal vNewValue As String)
    this.salesOrganization = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get conditionType() As String
    conditionType = this.conditionType
End Property

Public Property Let conditionType(ByVal vNewValue As String)
    this.conditionType = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get customerID() As String
    customerID = this.customerID
End Property

Public Property Let customerID(ByVal vNewValue As String)
    this.customerID = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get materialID() As String
    materialID = this.materialID
End Property

Public Property Let materialID(ByVal vNewValue As String)
    this.materialID = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get price() As String
    price = this.price
End Property

Public Property Let price(ByVal vNewValue As String)
    this.price = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get currencyType() As String
    currencyType = this.currencyType
End Property

Public Property Let currencyType(ByVal vNewValue As String)
    this.currencyType = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get unitOfPrice() As String
    unitOfPrice = this.unitOfPrice
End Property

Public Property Let unitOfPrice(ByVal vNewValue As String)
    this.unitOfPrice = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get unitOfMeasure() As String
    unitOfMeasure = this.unitOfMeasure
End Property

Public Property Let unitOfMeasure(ByVal vNewValue As String)
    this.unitOfMeasure = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get validFromDate() As String
    validFromDate = this.validFromDate
End Property

Public Property Let validFromDate(ByVal vNewValue As String)
    this.validFromDate = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get validToDate() As String
    validToDate = this.validToDate
End Property

Public Property Let validToDate(ByVal vNewValue As String)
    this.validToDate = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ActiveDataContainer() As Long
    ActiveDataContainer = this.ActiveDataContainer
End Property

Public Property Let ActiveDataContainer(ByVal vNewValue As Long)
    this.ActiveDataContainer = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ListTitle() As String
    Select Case Me.ActiveDataContainer
        Case DataContainer.FOR_CLIENTHISTORY
            ListTitle = "Client's List"
        Case DataContainer.FOR_PENDINGAPPROVALS
            ListTitle = "All Pending Records List"
        Case DataContainer.FOR_ALLHISTORY
            ListTitle = "All Records List"
    End Select
End Property

'-------------------------------------------------------------------------

Public Property Get IsApprover() As Boolean
    IsApprover = this.IsApprover
End Property

Public Property Let IsApprover(ByVal vNewValue As Boolean)
    this.IsApprover = vNewValue
End Property

'-------------------------------------------------------------------------
'Validations
'-------------------------------------------------------------------------

Public Function IsValidRecordToEdit() As Variant
    'Variable declaration
    Dim recordStatus    As String
    Dim TargetIndex     As Long
    Dim firstRow        As Long
    Dim lastRow         As Long
    Dim i               As Long
    'Calculate driving paramters
    TargetIndex = GetMethod.GetTargetRowIndex(Me.dataTable, Me.index, MainTableFields.COL_MAIN_INDEX)
    firstRow = LBound(Me.dataTable, 1)
    lastRow = UBound(Me.dataTable, 1)
    'Finding Status!
    For i = firstRow + 1 To lastRow
        If i = TargetIndex Then
            recordStatus = Me.dataTable(i, MainTableFields.COL_MAIN_recordStatus)
            Exit For
        End If
    Next i
    'Decision
    Select Case recordStatus
    Case RECORDSTATUS_PENDING
        IsValidRecordToEdit = True
    Case RECORDSTATUS_APPROVED
        IsValidRecordToEdit = "Your record has been already Approved! You have no permission to edit it further!"
    Case RECORDSTATUS_REJECTED
        IsValidRecordToEdit = True               'We are allowing client to edit Rejected records!
    End Select
    'Overriding
    If IsApprover Then IsValidRecordToEdit = True
End Function

Public Function GetDataForRecordsList() As Variant
    Dim i       As Long
    Dim j       As Long
    Dim k       As Long
    Dim UB1     As Long
    Dim UB2     As Long
    Dim tmp     As Variant
    Dim uname   As String
    UB1 = UBound(Me.dataTable, 1)
    UB2 = UBound(Me.dataTable, 2)
    ReDim tmp(1 To UB1, 1 To UB2)
    k = 1
    For i = 1 To UB1
        uname = GetMethod.GetLookupValue(Me.usersTable, Me.dataTable(i, MainTableFields.COL_MAIN_userID), UsersTableFields.COL_userId, UsersTableFields.COL_userName)
        If uname <> vbNullString Then
            For j = 1 To UB2
                If i = 1 Then
                    If j = MainTableFields.COL_MAIN_userID Then
                        tmp(k, j) = "User_Name"
                    Else
                        tmp(k, j) = Me.dataTable(i, j)
                    End If
                Else
                    If j = MainTableFields.COL_MAIN_userID Then
                        tmp(k, j) = uname
                    Else
                        tmp(k, j) = Me.dataTable(i, j)
                    End If
                End If
            Next j
            k = k + 1
        End If
    Next i
    GetDataForRecordsList = tmp
End Function

Public Function GetFilteredAndSortedList() As Variant
    
    'Variables Declaration
    Dim firstRowIndex       As Long
    Dim lastRowIndex        As Long
    Dim firstColumnIndex    As Long
    Dim lastColumnIndex     As Long
    Dim TotalMatchingRows   As Long
    Dim MatchingIndexes     As String
    Dim Records()           As String
    Dim Record              As Variant
    Dim tmp                 As Variant
    Dim i                   As Long
    Dim j                   As Long
    Dim IsValidRow          As Boolean
    Dim uname               As String
    
    'you are here because user has set some criteria!
    firstRowIndex = LBound(Me.dataTable, 1)
    lastRowIndex = UBound(Me.dataTable, 1)
    firstColumnIndex = LBound(Me.dataTable, 2)
    lastColumnIndex = UBound(Me.dataTable, 2)
    
    'if Selected Column is User_Name then
    If Me.selectedColumn = 3 Then
        Me.selectedValue = GetMethod.GetLookupValue(Me.usersTable, Me.selectedValue, UsersTableFields.COL_userName, UsersTableFields.COL_userId)
    End If
    
    'Init Parameters
    TotalMatchingRows = 1
    MatchingIndexes = 1 & SEPERATOR_ITEM
    
    'Getting Matching Indexes
    
    For i = firstRowIndex + 1 To lastRowIndex
        
        Select Case Me.selectedColumn
                
        Case MainTableFields.COL_MAIN_statusChangeDate
                
            IsValidRow = Me.Validator.IsValidToInclude(Me.dataTable(i, Me.selectedColumn), TYPE_DATEBETWEENRANGE, Me.selectedValue, Me.selectedValue)
            If Not IsValidRow Then GoTo IgnoreThisRow
            
        Case Else
            
            IsValidRow = Me.Validator.IsValidToInclude(Me.dataTable(i, Me.selectedColumn), TYPE_STRINGMATCH, Me.selectedValue)
            If Not IsValidRow Then GoTo IgnoreThisRow
                
        End Select
        
        'if you are here means, the row meets all criterias and happy to include!
        If IsValidRow Then
            TotalMatchingRows = TotalMatchingRows + 1
            MatchingIndexes = MatchingIndexes & i & SEPERATOR_ITEM
        End If
        
IgnoreThisRow:
    Next i
    
    'Remove Last Seperator Characters
    MatchingIndexes = VBA.Left$(MatchingIndexes, VBA.Len(MatchingIndexes) - VBA.Len(SEPERATOR_ITEM))
    
    'Generate Array of all matching indexes
    Records = VBA.Split(MatchingIndexes, SEPERATOR_ITEM)
    
    'Redim Tmp Array
    ReDim tmp(firstRowIndex To TotalMatchingRows, firstColumnIndex To lastColumnIndex)
    
    'Create Export Table
    i = 1
    For Each Record In Records
        uname = GetMethod.GetLookupValue(Me.usersTable, Me.dataTable(Record, MainTableFields.COL_MAIN_userID), UsersTableFields.COL_userId, UsersTableFields.COL_userName)
        If uname <> vbNullString Then
            For j = firstColumnIndex To lastColumnIndex
            
                'GUARD CLAUSE : For preventing cells original value instead cell's value will be used for _
                                looking up relative values from the another table!
                If j = MainTableFields.COL_MAIN_userID Then
                    tmp(i, j) = uname
                    GoTo NextIteration
                End If
            
                'NO GUARD CLAUSE : For Rest of the cells
                tmp(i, j) = Me.dataTable(Record, j)
NextIteration:
            Next j
            i = i + 1
        End If
    Next Record

    'Submit value to Export Data Array
    GetFilteredAndSortedList = tmp

End Function

'-------------------------------------------------------------------------
'Class Events
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Validator = Nothing
End Sub

