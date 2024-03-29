VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PriceFormModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PriceApproval.Model.PriceForm")
Option Explicit

'-------------------------------------------------------------------------

Private Type TPriceFormModelComponents
    CheckFields(MainTableFields.COL_MAIN_ConditionType To MainTableFields.COL_MAIN_validToDate) As Variant
    'Tables
    mainTable As Variant
    'Single Column List
    curenciesList As Variant
    unitOfMeasuresList As Variant
    salesOrganizationList As Variant
    distributionChannelList As Variant
    'Fields
    Index As Long
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
    'Attribute
    operation As FormOperation
    
End Type

Private this As TPriceFormModelComponents

'-------------------------------------------------------------------------

Public Property Get mainTable() As Variant
    mainTable = this.mainTable
End Property

Public Property Let mainTable(ByVal vNewValue As Variant)
    this.mainTable = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get curenciesList() As Variant
    curenciesList = this.curenciesList
End Property

Public Property Let curenciesList(ByVal vNewValue As Variant)
    this.curenciesList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get unitOfMeasuresList() As Variant
    unitOfMeasuresList = this.unitOfMeasuresList
End Property

Public Property Let unitOfMeasuresList(ByVal vNewValue As Variant)
    this.unitOfMeasuresList = vNewValue
End Property


'-------------------------------------------------------------------------

Public Property Get distributionChannelList() As Variant
    distributionChannelList = this.distributionChannelList
End Property

Public Property Let distributionChannelList(ByVal vNewValue As Variant)
    this.distributionChannelList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get salesOrganizationList() As Variant
    salesOrganizationList = this.salesOrganizationList
End Property

Public Property Let salesOrganizationList(ByVal vNewValue As Variant)
    this.salesOrganizationList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get Index() As Long
    Index = this.Index
End Property

Public Property Let Index(ByVal vNewValue As Long)
    this.Index = vNewValue
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

Public Property Get operation() As FormOperation
    operation = this.operation
End Property

Public Property Let operation(ByVal vNewValue As FormOperation)
    this.operation = vNewValue
End Property

'-------------------------------------------------------------------------
'Init Method
'-------------------------------------------------------------------------

Public Sub InItModel(ByVal mainData As Variant, _
                     ByVal ListofCurrencies As Variant, _
                     ByVal ListOfUnits As Variant, _
                     ByVal ListOfSalesOrganizations As Variant, _
                     ByVal ListOfDistributionChannels As Variant, _
                     Optional TargetIndex As Long = 0)
    Me.mainTable = mainData
    Me.curenciesList = ListofCurrencies
    Me.unitOfMeasuresList = ListOfUnits
    Me.salesOrganizationList = ListOfSalesOrganizations
    Me.distributionChannelList = ListOfDistributionChannels
    Me.Index = TargetIndex
    'reset check field array
    'Changed this to approach because earlier code wasn't resetting the each items of the array to false!
    Call AppMethods.ResetSingleDimensionalArray(this.CheckFields)
    'Keep Not-Mendatory Field Always True!
    this.CheckFields(MainTableFields.COL_MAIN_unitOfPrice) = True 'Unit of Price!
    
End Sub

'-------------------------------------------------------------------------
'public Methods
'-------------------------------------------------------------------------

Public Sub SetPropertiesToNewRecordState(ByVal ActiveUserID As String)
    'Flush All the properties before we start operation
    Call SetPropertiesToNullState
    With Me
        .operation = OPERATION_NEW
        .Index = 0
        .recordStatus = RECORDSTATUS_PENDING
        .userID = ActiveUserID
        .statusChangeDate = VBA.Format$(VBA.Now, DATEFORMAT_BACKEND)
        .conditionType = "Y0PR"
        .salesOrganization = "2961"
        .unitOfPrice = 1
        .validFromDate = VBA.Format$(VBA.Now, DATEFORMAT_BACKEND)
        .validToDate = END_OF_THE_EARTH
    End With
End Sub

Public Sub SetPropertiesToUpdateRecordState()
    'Flush All the properties before we start operation
    Call SetPropertiesToNullState
    'update operation begins
    Dim TargetRow As Long
    TargetRow = AppMethods.GetTargetRowIndex(Me.mainTable, Me.Index, MainTableFields.COL_MAIN_INDEX)
    With Me
        .operation = OPERATION_UPDATE
        .RecordID = Me.mainTable(TargetRow, MainTableFields.COL_MAIN_recordID)
        .userID = Me.mainTable(TargetRow, MainTableFields.COL_MAIN_userID)
        .recordStatus = Me.mainTable(TargetRow, MainTableFields.COL_MAIN_recordStatus)
        .statusChangeDate = Me.mainTable(TargetRow, MainTableFields.COL_MAIN_statusChangeDate)
        .conditionType = Me.mainTable(TargetRow, MainTableFields.COL_MAIN_ConditionType)
        .salesOrganization = Me.mainTable(TargetRow, MainTableFields.COL_MAIN_SalesOrganization)
        .distributionChannel = Me.mainTable(TargetRow, MainTableFields.COL_Main_DistributionChannel)
        .customerID = Me.mainTable(TargetRow, MainTableFields.COL_MAIN_customerID)
        .materialID = Me.mainTable(TargetRow, MainTableFields.COL_MAIN_materialID)
        .price = Me.mainTable(TargetRow, MainTableFields.COL_MAIN_price)
        .currencyType = Me.mainTable(TargetRow, MainTableFields.COL_MAIN_currency)
        .unitOfPrice = Me.mainTable(TargetRow, MainTableFields.COL_MAIN_unitOfPrice)
        .unitOfMeasure = Me.mainTable(TargetRow, MainTableFields.COL_MAIN_unitOfMeasure)
        .validFromDate = Me.mainTable(TargetRow, MainTableFields.COL_MAIN_validFromDate)
        .validToDate = Me.mainTable(TargetRow, MainTableFields.COL_MAIN_validToDate)
    End With
End Sub

Private Sub SetPropertiesToNullState()
    With Me
        .RecordID = vbNullString
        .userID = vbNullString
        .recordStatus = vbNullString
        .statusChangeDate = vbNullString
        .conditionType = vbNullString
        .salesOrganization = vbNullString
        .distributionChannel = vbNullString
        .customerID = vbNullString
        .materialID = vbNullString
        .price = vbNullString
        .currencyType = vbNullString
        .unitOfPrice = vbNullString
        .unitOfMeasure = vbNullString
        .validFromDate = vbNullString
        .validToDate = vbNullString
    End With
End Sub

'-------------------------------------------------------------------------
'Validation Method
'-------------------------------------------------------------------------

Public Function IsEverythingOkayToSaveRecord() As Variant
    IsEverythingOkayToSaveRecord = ValidationServices.IsEverythingOkay(this.CheckFields, _
                                                              PriceApprovalDataResources.arrListOfColumns_MAIN_Table, _
                                                              MainTableFields.COL_MAIN_ConditionType, _
                                                              MainTableFields.COL_MAIN_validToDate)
End Function

Public Function IsValidField(ByVal FieldIndex As MainTableFields) As Boolean
    Dim startDate As Date
    Dim endDate As Date
    
    'Reset CheckField's State to False
    this.CheckFields(FieldIndex) = False
    
    'Check the validation of field based on field index
    Select Case FieldIndex
        
        Case MainTableFields.COL_MAIN_ConditionType
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.conditionType, VBA.TypeName(Me.conditionType))
            If Not IsValidField = True Then Exit Function
            'check 2
            IsValidField = ValidationServices.IsValidStringLength(Me.conditionType, 4)
            If Not IsValidField = True Then Exit Function
            
        Case MainTableFields.COL_MAIN_SalesOrganization
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.salesOrganization, VBA.TypeName(Me.salesOrganization))
            If Not IsValidField = True Then Exit Function
            
        Case MainTableFields.COL_Main_DistributionChannel
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.distributionChannel, VBA.TypeName(Me.distributionChannel))
            If Not IsValidField = True Then Exit Function
        
        Case MainTableFields.COL_MAIN_customerID
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.customerID, VBA.TypeName(Me.customerID))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = VBA.IsNumeric(Me.customerID)
            If Not IsValidField = True Then Exit Function
            'check 3
            IsValidField = ValidationServices.IsValidStringLength(Me.customerID, 6)
            If Not IsValidField = True Then Exit Function
            'Check 4
            IsValidField = ValidationServices.IsBetweenNumericalRange(Me.customerID, INDEX_CUSTOMERID_FIRST, INDEX_CUSTOMERID_LAST)
            If Not IsValidField = True Then Exit Function
         
        Case MainTableFields.COL_MAIN_materialID
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.materialID, VBA.TypeName(Me.materialID))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = VBA.IsNumeric(Me.materialID)
            If Not IsValidField = True Then Exit Function
            'check 3
            IsValidField = ValidationServices.IsValidStringLength(Me.materialID, 8)
            If Not IsValidField = True Then Exit Function
            'Check 4
            IsValidField = ValidationServices.IsBetweenNumericalRange(Me.materialID, INDEX_MATERIALID_FIRST, INDEX_MATERIALID_LAST)
            If Not IsValidField = True Then Exit Function
            
            
        Case MainTableFields.COL_MAIN_price
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.price, VBA.TypeName(Me.price))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = VBA.IsNumeric(Me.price)
            If Not IsValidField = True Then Exit Function
            'Check 3
            IsValidField = ValidationServices.IsBetweenNumericalRange(Me.price, MIN_PRICE_VALUE, MAX_PRICE_VALUE)
            If Not IsValidField = True Then Exit Function
            'Check 4
            IsValidField = ValidationServices.IsValidStringLengthUnderLimit(Me.price, 6)
            
            
        Case MainTableFields.COL_MAIN_currency
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.currencyType, VBA.TypeName(Me.currencyType))
            If Not IsValidField = True Then Exit Function
            
            
        Case MainTableFields.COL_MAIN_unitOfPrice
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.unitOfPrice, VBA.TypeName(Me.unitOfPrice))
            If IsValidField = False Then
                IsValidField = True
                this.CheckFields(FieldIndex) = True
                Exit Function
            End If
            'Check 2
            IsValidField = VBA.IsNumeric(Me.unitOfPrice)
            If Not IsValidField = True Then Exit Function
            'Check 3
            IsValidField = ValidationServices.IsBetweenNumericalRange(Me.unitOfPrice, MIN_UNITOFPRICE_VALUE, MAX_UNITOFPRICE_VALUE)
            If Not IsValidField = True Then Exit Function
            
            
        Case MainTableFields.COL_MAIN_unitOfMeasure
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.unitOfMeasure, VBA.TypeName(Me.unitOfMeasure))
            If Not IsValidField = True Then Exit Function
            
            
        Case MainTableFields.COL_MAIN_validFromDate
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.validFromDate, VBA.TypeName(Me.validFromDate))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = ValidationServices.IsValidGermanFormatDate(Me.validFromDate)
            If Not IsValidField = True Then Exit Function
            'Check 3
            'This check confirm that the from date should be current date!
            'but when user will open the record to update, at that time,
            'date would be future and hence this validation will failed! so,
            'to handle that only for the UPDATE operation we are not going to
            'perform this check!
            If Me.operation = OPERATION_NEW Then
                startDate = VBA.Format$(VBA.Now, DATEFORMAT_BACKEND)
                endDate = startDate
                IsValidField = ValidationServices.IsValidDateInRange(Me.validFromDate, startDate, endDate)
                If Not IsValidField = True Then Exit Function
            End If
            'if everything okay then save date in backend format
            Me.validFromDate = VBA.Format$(Me.validFromDate, GetDateFormat)
            
        Case MainTableFields.COL_MAIN_validToDate
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.validToDate, VBA.TypeName(Me.validToDate))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = ValidationServices.IsValidGermanFormatDate(Me.validToDate)
            If Not IsValidField = True Then Exit Function
            'Check 3
            If Me.operation = OPERATION_NEW Then
                startDate = VBA.Format$(VBA.Now, GetDateFormat)
                endDate = VBA.CDate(END_OF_THE_EARTH)
                IsValidField = ValidationServices.IsValidDateInRange(Me.validToDate, startDate, endDate)
                If Not IsValidField = True Then Exit Function
            End If
            'Check 4
            If Not Me.validToDate = VBA.CDate(END_OF_THE_EARTH) Then
                IsValidField = CDate(Me.validToDate) > CDate(Me.validFromDate)
                If Not IsValidField = True Then Exit Function
            End If
            'if everything okay then save date in backedn format
            Me.validToDate = VBA.Format$(Me.validToDate, GetDateFormat)
            
    End Select
    
    'If All Goes Well!
    this.CheckFields(FieldIndex) = True
    
End Function


