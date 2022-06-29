VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExportFormModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("AppObjects.Model.ExportForm")
Option Explicit

Private Type TExportFormModelComponents

    CheckFields(ExportFormFields.FIELD_FROMDATE To ExportFormFields.FIELD_RECORDSTATUS) As Variant
    'Table
    dataTable As Variant
    userTable As Variant
    exportTable As Variant
    'Single Column List
    customerIDsList As Variant
    userIDsList As Variant
    statusesList As Variant
    'Fields
    FromDate As String
    ToDate As String
    customerID As String
    userID As String
    recordStatus As String
    'Attributes
    RecordsToProcess As String
    IsRecordsNeedsToBeProcessed As Boolean
End Type

Private this As TExportFormModelComponents

'-------------------------------------------------------------------------

Public Property Get dataTable() As Variant
    dataTable = this.dataTable
End Property

Public Property Let dataTable(ByVal vNewValue As Variant)
    this.dataTable = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userTable() As Variant
    userTable = this.userTable
End Property

Public Property Let userTable(ByVal vNewValue As Variant)
    this.userTable = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get exportTable() As Variant
    exportTable = this.exportTable
End Property

Public Property Let exportTable(ByVal vNewValue As Variant)
    this.exportTable = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get statusesList() As Variant
    statusesList = this.statusesList
End Property

Public Property Let statusesList(ByVal vNewValue As Variant)
    this.statusesList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userIDsList() As Variant
    userIDsList = this.userIDsList
End Property

Public Property Let userIDsList(ByVal vNewValue As Variant)
    this.userIDsList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get customerIDsList() As Variant
    customerIDsList = this.customerIDsList
End Property

Public Property Let customerIDsList(ByVal vNewValue As Variant)
    this.customerIDsList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get recordStatus() As String
    recordStatus = this.recordStatus
End Property

Public Property Let recordStatus(ByVal vNewValue As String)
    this.recordStatus = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userID() As String
    userID = this.userID
End Property

Public Property Let userID(ByVal vNewValue As String)
    this.userID = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get customerID() As String
    customerID = this.customerID
End Property

Public Property Let customerID(ByVal vNewValue As String)
    this.customerID = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get FromDate() As String
    FromDate = this.FromDate
End Property

Public Property Let FromDate(ByVal vNewValue As String)
    this.FromDate = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ToDate() As String
    ToDate = this.ToDate
End Property

Public Property Let ToDate(ByVal vNewValue As String)
    this.ToDate = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get RecordsToProcess() As String
    RecordsToProcess = this.RecordsToProcess
End Property

Public Property Let RecordsToProcess(ByVal vNewValue As String)
    this.RecordsToProcess = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get IsRecordsNeedsToBeProcessed() As Boolean
    IsRecordsNeedsToBeProcessed = this.IsRecordsNeedsToBeProcessed
End Property

Public Property Let IsRecordsNeedsToBeProcessed(ByVal vNewValue As Boolean)
    this.IsRecordsNeedsToBeProcessed = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt Method
'-------------------------------------------------------------------------

Public Sub InItModel(ByVal mainData As Variant, _
                    ByVal userData As Variant, _
                    ByVal listOfCustomerIDs As Variant, _
                    ByVal listOfUserIDs As Variant, _
                    ByVal listOfRecordStatuses As Variant)
                    
    'populate fields
    Me.dataTable = mainData
    Me.userTable = userData
    Me.customerIDsList = listOfCustomerIDs
    Me.userIDsList = listOfUserIDs
    Me.statusesList = listOfRecordStatuses
    'reset check field array
    Call AppMethods.ResetSingleDimensionalArray(this.CheckFields, True)
    
End Sub

'-------------------------------------------------------------------------
'public Methods
'-------------------------------------------------------------------------

Public Sub SetPropertiesToDefaultState()
    'Flush All the properties before we start operation
    Call SetPropertiesToNullState
    With Me
        .FromDate = VBA.Format$(VBA.Now, DATEFORMAT_BACKEND)
        .ToDate = VBA.Format$(VBA.Now, DATEFORMAT_BACKEND)
        .recordStatus = DataResources.arrRecordStatusesList(2)
    End With
End Sub

Private Sub SetPropertiesToNullState()
    With Me
        .FromDate = vbNullString
        .ToDate = vbNullString
        .customerID = vbNullString
        .userID = vbNullString
        .recordStatus = vbNullString
        'Setting up the default value of Attributes of this Model
        .RecordsToProcess = vbNullString
        .IsRecordsNeedsToBeProcessed = False
    End With
End Sub

'-------------------------------------------------------------------------
'Validation Methods
'-------------------------------------------------------------------------

Public Function IsEverythingOkayToExport() As Variant
    'check 1
    If (Me.FromDate <> vbNullString And Me.ToDate = vbNullString) Or _
        (Me.FromDate = vbNullString And Me.ToDate <> vbNullString) Then
        IsEverythingOkayToExport = "`Date From` & `Date To` fields either they should be filled with proper dates or either they should left blank! But one from both field should not be left blank!"
        Exit Function
    End If
    'check 2
    IsEverythingOkayToExport = _
        ValidationServices.IsEverythingOkay(this.CheckFields, _
            DataResources.arrListOfFields_EXPORT_Form, _
            ExportFormFields.FIELD_FROMDATE, _
            ExportFormFields.FIELD_RECORDSTATUS)
End Function

Public Function IsValidField(ByVal FieldIndex As ExportFormFields) As Boolean
    
    Dim startDate As Date
    Dim endDate As Date
    
    'Reset CheckField's State to False
    this.CheckFields(FieldIndex) = False
    
    'Check the validation of field based on field index
    Select Case FieldIndex
        
        Case ExportFormFields.FIELD_FROMDATE
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.FromDate, VBA.TypeName(Me.FromDate))
            If IsValidField = False Then
                IsValidField = True
                this.CheckFields(FieldIndex) = True
                Exit Function
            End If
            'Check 2
            IsValidField = ValidationServices.IsValidGermanFormatDate(Me.FromDate)
            If Not IsValidField = True Then Exit Function
            'Check 3
            startDate = VBA.Format$(START_OF_THE_CENTURY, DATEFORMAT_BACKEND)
            endDate = VBA.Format$(VBA.Now, DATEFORMAT_BACKEND)
            IsValidField = ValidationServices.IsValidDateInRange(Me.FromDate, startDate, endDate)
            If Not IsValidField = True Then Exit Function
            'if everything okay then save date in backend format
            Me.FromDate = VBA.Format$(Me.FromDate, GetDateFormat)
            
            
        Case ExportFormFields.FIELD_TODATE
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.ToDate, VBA.TypeName(Me.ToDate))
            If IsValidField = False Then
                IsValidField = True
                this.CheckFields(FieldIndex) = True
                Exit Function
            End If
            'Check 2
            IsValidField = ValidationServices.IsValidGermanFormatDate(Me.ToDate)
            If Not IsValidField = True Then Exit Function
            'Check 3
            startDate = VBA.Format$(START_OF_THE_CENTURY, DATEFORMAT_BACKEND)
            endDate = VBA.Format$(VBA.Now, DATEFORMAT_BACKEND)
            IsValidField = ValidationServices.IsValidDateInRange(Me.ToDate, startDate, endDate)
            If Not IsValidField = True Then Exit Function
            'if everything okay then save date in backend format
            Me.ToDate = VBA.Format$(Me.ToDate, GetDateFormat)
            
        
        Case ExportFormFields.FIELD_CUSTOMERID
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.customerID, VBA.TypeName(Me.customerID))
            If IsValidField = False Then
                IsValidField = True
                this.CheckFields(FieldIndex) = True
                Exit Function
            End If
            
        Case ExportFormFields.FIELD_USERID
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.userID, VBA.TypeName(Me.userID))
            If IsValidField = False Then
                IsValidField = True
                this.CheckFields(FieldIndex) = True
                Exit Function
            End If
            
        Case ExportFormFields.FIELD_RECORDSTATUS
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.recordStatus, VBA.TypeName(Me.recordStatus))
            If IsValidField = False Then
                IsValidField = True
                this.CheckFields(FieldIndex) = True
                Exit Function
            End If
            
    End Select
    
    'If All Goes Well!
    this.CheckFields(FieldIndex) = True
    
End Function

'-------------------------------------------------------------------------
'Functions & methods
'-------------------------------------------------------------------------

'Following function will filter data based on selected criteria!

Public Sub GenerateExportData()
    
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
    Dim uName               As String
    Dim Answer              As VbMsgBoxResult
    
    'you are here because user has set some criteria!
    firstRowIndex = LBound(Me.dataTable, 1)
    lastRowIndex = UBound(Me.dataTable, 1)
    firstColumnIndex = LBound(Me.dataTable, 2)
    lastColumnIndex = UBound(Me.dataTable, 2)
    
    'Init Parameters
    TotalMatchingRows = 1
    MatchingIndexes = 1 & SEPERATOR_ITEM
    
    'Getting Matching Indexes
    
    For i = firstRowIndex + 1 To lastRowIndex
        
    'Filter Layer 1 : Date Filter
        'Checking First condition : Edit Date
        
        IsValidRow = ValidationServices.IsValidToInclude(Me.dataTable(i, MainTableFields.COL_MAIN_statusChangeDate), TYPE_DATEBETWEENRANGE, Me.FromDate, Me.ToDate)
        If Not IsValidRow Then GoTo IgnoreThisRow
        
    'Filter Layer 2 : String/Number Matching
        'Check Second condition : Customer ID
        IsValidRow = ValidationServices.IsValidToInclude(Me.dataTable(i, MainTableFields.COL_MAIN_customerID), TYPE_STRINGMATCH, Me.customerID)
        If Not IsValidRow Then GoTo IgnoreThisRow
        
    'Filter Layer 3 : String/Number Matching
        'Check Third condition : User ID
        IsValidRow = ValidationServices.IsValidToInclude(AppMethods.GetLookupValue(Me.userTable, Me.dataTable(i, MainTableFields.COL_MAIN_userID), UsersTableFields.COL_userId, UsersTableFields.COL_userName), TYPE_STRINGMATCH, Me.userID)
        If Not IsValidRow Then GoTo IgnoreThisRow
        
    'Filter Layer 4 : String/Number Matching
        'Check Fourth condition : Record Status
        IsValidRow = ValidationServices.IsValidToInclude(Me.dataTable(i, MainTableFields.COL_MAIN_recordStatus), TYPE_STRINGMATCH, Me.recordStatus)
        If Not IsValidRow Then GoTo IgnoreThisRow
        
    'Keep adding here filters!
        
        'if you are here means, the row meets all criterias and happy to include!
        If IsValidRow Then
            TotalMatchingRows = TotalMatchingRows + 1
            MatchingIndexes = MatchingIndexes & i & SEPERATOR_ITEM
        End If
        
IgnoreThisRow:
    Next i
    
    'Remove Last Seperator Characters
    MatchingIndexes = VBA.Left$(MatchingIndexes, VBA.Len(MatchingIndexes) - VBA.Len(SEPERATOR_ITEM))
    
    'If there is no record then there's no meaning of processing. hence we are bypassing this.
    If VBA.Len(MatchingIndexes) > 1 Then
        'Set RecordToProcess Property with Value of MatchingIndex! Because we have to mark exporting records as processed as well
        Me.IsRecordsNeedsToBeProcessed = False
        Me.RecordsToProcess = VBA.Right(MatchingIndexes, VBA.Len(MatchingIndexes) - VBA.Len("1" & SEPERATOR_ITEM))
        
        'Check if the Export is valid to process or not? We have to ask a question as well.
        If Me.recordStatus = RECORDSTATUS_APPROVED Then
            Answer = MsgBox("Do you want to Process the Export Data?", vbInformation + vbYesNo, SIGN)
            Me.IsRecordsNeedsToBeProcessed = (Answer = vbYes)
        End If
    End If
    
    'Generate Array of all matching indexes
    Records = VBA.Split(MatchingIndexes, SEPERATOR_ITEM)
    
    'Redim Tmp Array
    ReDim tmp(firstRowIndex To TotalMatchingRows, firstColumnIndex To lastColumnIndex)
    
    'Create Export Table
    i = 1
    For Each Record In Records
        
        uName = AppMethods.GetLookupValue(Me.userTable, Me.dataTable(Record, MainTableFields.COL_MAIN_userID), UsersTableFields.COL_userId, UsersTableFields.COL_userName)
        
        If uName <> vbNullString Then
            
            For j = firstColumnIndex To lastColumnIndex
            
                'GUARD CLAUSE : For preventing excel's automatic conversion of cell values from specific value to _
                                numerical whole number due
                If j = MainTableFields.COL_Main_DistributionChannel And VBA.IsNumeric(Me.dataTable(Record, j)) Then
                    tmp(i, j) = "'" & Me.dataTable(Record, j)
                    GoTo NextIteration
                End If
                
                'GUARD CLAUSE : For preventing cells original value instead cell's value will be used for _
                                looking up relative values from the another table!
                If j = MainTableFields.COL_MAIN_userID Then
                    tmp(i, j) = uName
                    GoTo NextIteration
                End If
                
                'If the export reports are being processed then change the value
                If j = MainTableFields.COL_MAIN_recordStatus Then
                    If Me.IsRecordsNeedsToBeProcessed Then
                        tmp(i, j) = GlobalResources.RECORDSTATUS_PROCESSED
                        GoTo NextIteration
                    End If
                End If
                
                'NO GUARD CLAUSE : For Rest of the cells
                tmp(i, j) = Me.dataTable(Record, j)
NextIteration:

            Next j
            
            i = i + 1
            
        End If
        
    Next Record

    'Submit value to Export Data Array
    Me.exportTable = tmp
    
End Sub


