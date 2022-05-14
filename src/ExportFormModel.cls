VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExportFormModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ExportForm"
Option Explicit

Private Type TExportFormModelComponents
    'Validation Object
    Validator As ValidationServices
    CheckFields(ExportFormFields.FIELD_FROMDATE To ExportFormFields.FIELD_RECORDSTATUS) As Variant
    'Table
    dataTable As Variant
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
End Type

Private this As TExportFormModelComponents

'-------------------------------------------------------------------------

Private Property Get Validator() As ValidationServices
    Set Validator = this.Validator
End Property

Private Property Set Validator(ByVal vNewValue As ValidationServices)
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
'InIt Method
'-------------------------------------------------------------------------

Public Sub InItModel(ByVal mainData As Variant, _
                    ByVal listOfCustomerIDs As Variant, _
                    ByVal listOfUserIDs As Variant, _
                    ByVal listOfRecordStatuses As Variant)
    'In it Validator
    If Validator Is Nothing Then Set Validator = New ValidationServices
    'populate fields
    Me.dataTable = mainData
    Me.customerIDsList = listOfCustomerIDs
    Me.userIDsList = listOfUserIDs
    Me.statusesList = listOfRecordStatuses
    'reset check field array
    Call Validator.ResetSingleDimensionalArray(this.CheckFields, True)
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
    End With
End Sub

Private Sub SetPropertiesToNullState()
    With Me
        .FromDate = vbNullString
        .ToDate = vbNullString
        .customerID = vbNullString
        .userID = vbNullString
        .recordStatus = vbNullString
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
        Validator.IsEverythingOkay(this.CheckFields, _
            modDataSources.arrListOfFields_EXPORT_Form, _
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
            IsValidField = Validator.IsNotNullValue(Me.FromDate, VBA.TypeName(Me.FromDate))
            If IsValidField = False Then
                IsValidField = True
                this.CheckFields(FieldIndex) = True
                Exit Function
            End If
            'Check 2
            IsValidField = Validator.IsValidGermanFormatDate(Me.FromDate)
            If Not IsValidField = True Then Exit Function
            'Check 3
            startDate = VBA.Format$(START_OF_THE_CENTURY, DATEFORMAT_BACKEND)
            endDate = VBA.Format$(VBA.Now, DATEFORMAT_BACKEND)
            IsValidField = Validator.IsValidDateInRange(Me.FromDate, startDate, endDate)
            If Not IsValidField = True Then Exit Function
            'if everything okay then save date in backend format
            Me.FromDate = Validator.GetDateFromGermanDateString(Me.FromDate)
            
        Case ExportFormFields.FIELD_TODATE
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.ToDate, VBA.TypeName(Me.ToDate))
            If IsValidField = False Then
                IsValidField = True
                this.CheckFields(FieldIndex) = True
                Exit Function
            End If
            'Check 2
            IsValidField = Validator.IsValidGermanFormatDate(Me.ToDate)
            If Not IsValidField = True Then Exit Function
            'Check 3
            startDate = VBA.Format$(START_OF_THE_CENTURY, DATEFORMAT_BACKEND)
            endDate = VBA.Format$(VBA.Now, DATEFORMAT_BACKEND)
            IsValidField = Validator.IsValidDateInRange(Me.ToDate, startDate, endDate)
            If Not IsValidField = True Then Exit Function
            'if everything okay then save date in backend format
            Me.ToDate = Validator.GetDateFromGermanDateString(Me.ToDate)
        
        Case ExportFormFields.FIELD_CUSTOMERID
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.customerID, VBA.TypeName(Me.customerID))
            If IsValidField = False Then
                IsValidField = True
                this.CheckFields(FieldIndex) = True
                Exit Function
            End If
            
        Case ExportFormFields.FIELD_USERID
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.userID, VBA.TypeName(Me.userID))
            If IsValidField = False Then
                IsValidField = True
                this.CheckFields(FieldIndex) = True
                Exit Function
            End If
            
        Case ExportFormFields.FIELD_RECORDSTATUS
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.recordStatus, VBA.TypeName(Me.recordStatus))
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
    Dim I                   As Long
    Dim J                   As Long
    Dim IsValidRow          As Boolean
    
    'If every criteria field is empty then simply every row of the main data has to be exported!
    If Me.FromDate = vbNullString And _
        Me.ToDate = vbNullString And _
        Me.customerID = vbNullString And _
        Me.userID = vbNullString And _
        Me.recordStatus = vbNullString Then
        Me.exportTable = Me.dataTable
        Exit Sub
    End If
    
    'you are here because user has set some criteria!
    firstRowIndex = LBound(Me.dataTable, 1)
    lastRowIndex = UBound(Me.dataTable, 1)
    firstColumnIndex = LBound(Me.dataTable, 2)
    lastColumnIndex = UBound(Me.dataTable, 2)
    
    'Init Parameters
    TotalMatchingRows = 1
    MatchingIndexes = 1 & SEPERATOR_ITEM
    
    'Getting Matching Indexes
    
    For I = firstRowIndex + 1 To lastRowIndex
        
    'Filter Layer 1 : Date Filter
        'Checking First condition : Edit Date
        
        IsValidRow = Validator.IsValidToInclude(Me.dataTable(I, MainTableFields.COL_MAIN_statusChangeDate), TYPE_DATEBETWEENRANGE, Me.FromDate, Me.ToDate)
        If Not IsValidRow Then GoTo IgnoreThisRow
        
    'Filter Layer 2 : String/Number Matching
        'Check Second condition : Customer ID
        IsValidRow = Validator.IsValidToInclude(Me.dataTable(I, MainTableFields.COL_MAIN_customerID), TYPE_STRINGMATCH, Me.customerID)
        If Not IsValidRow Then GoTo IgnoreThisRow
        
    'Filter Layer 3 : String/Number Matching
        'Check Third condition : User ID
        IsValidRow = Validator.IsValidToInclude(Me.dataTable(I, MainTableFields.COL_MAIN_userID), TYPE_STRINGMATCH, Me.userID)
        If Not IsValidRow Then GoTo IgnoreThisRow
        
    'Filter Layer 4 : String/Number Matching
        'Check Fourth condition : Record Status
        IsValidRow = Validator.IsValidToInclude(Me.dataTable(I, MainTableFields.COL_MAIN_recordStatus), TYPE_STRINGMATCH, Me.recordStatus)
        If Not IsValidRow Then GoTo IgnoreThisRow
        
    'Keep adding here filters!
        
        'if you are here means, the row meets all criterias and happy to include!
        If IsValidRow Then
            TotalMatchingRows = TotalMatchingRows + 1
            MatchingIndexes = MatchingIndexes & I & SEPERATOR_ITEM
        End If
        
IgnoreThisRow:
    Next I
    
    'Remove Last Seperator Characters
    MatchingIndexes = VBA.Left$(MatchingIndexes, VBA.Len(MatchingIndexes) - VBA.Len(SEPERATOR_ITEM))
    
    'Generate Array of all matching indexes
    Records = VBA.Split(MatchingIndexes, SEPERATOR_ITEM)
    
    'Redim Tmp Array
    ReDim tmp(firstRowIndex To TotalMatchingRows, firstColumnIndex To lastColumnIndex)
    
    'Create Export Table
    I = 1
    For Each Record In Records
        For J = firstColumnIndex To lastColumnIndex
            If VBA.IsNumeric(Me.dataTable(Record, J)) Then
                tmp(I, J) = "'" & Me.dataTable(Record, J)
            Else
                tmp(I, J) = Me.dataTable(Record, J)
            End If
        Next J
        I = I + 1
    Next Record

    'Submit value to Export Data Array
    Me.exportTable = tmp
    
End Sub

'-------------------------------------------------------------------------
'Class Events
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Validator = Nothing
End Sub


