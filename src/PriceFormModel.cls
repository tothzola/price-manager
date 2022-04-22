VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PriceFormModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PriceForm"
Option Explicit

'-------------------------------------------------------------------------

Private Type TPriceFormModelComponents
    'Validation Object
    Validator As ValidationServices
    CheckFields(MainTableFields.COL_MAIN_recordStatus To MainTableFields.COL_MAIN_validToDate) As Variant
    'Tables
    mainTable As Variant
    'Fields
    index As Long
    recordID As String
    recordStatus As String
    statusChangeDate As String
    customerID As String
    materialID As String
    price As String
    currencyType As String
    unitOfPrice As String
    unitOfMeasure As String
    validFromDate  As String
    validToDate As String
End Type

Private this As TPriceFormModelComponents

'-------------------------------------------------------------------------

Private Property Get Validator() As ValidationServices
    Set Validator = this.Validator
End Property

Private Property Set Validator(ByVal vNewValue As ValidationServices)
    Set this.Validator = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get mainTable() As Variant
    mainTable = this.mainTable
End Property

Public Property Let mainTable(ByVal vNewValue As Variant)
    this.mainTable = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get index() As Long
    index = this.index
End Property

Public Property Let index(ByVal vNewValue As Long)
    this.index = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get recordID() As String
    recordID = this.recordID
End Property

Public Property Let recordID(ByVal vNewValue As String)
    this.recordID = vNewValue
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
'Init Method
'-------------------------------------------------------------------------

Public Sub InItModel(ByVal mainData As Variant)
    If Validator Is Nothing Then Set Validator = New ValidationServices
    Me.mainTable = mainData
    'reset check field array
    Dim CheckField As Variant
    For Each CheckField In this.CheckFields
        CheckField = False
    Next CheckField
End Sub

'-------------------------------------------------------------------------
'Validation Method
'-------------------------------------------------------------------------

Public Function IsEverythingOkayToSaveRecord() As Variant
    IsEverythingOkayToSaveRecord = Validator.IsEverythingOkay(this.CheckFields, _
                                                            modDataSources.arrListOfColumns_MAIN_Table, _
                                                            MainTableFields.COL_MAIN_recordStatus, _
                                                            MainTableFields.COL_MAIN_validToDate)
End Function

Public Function IsValidField(ByVal FieldIndex As MainTableFields) As Boolean
    Dim startDate As Date
    Dim endDate As Date
    'Reset CheckField's State to False
    this.CheckFields(FieldIndex) = False
    'Check the validation of field based on field index
    Select Case FieldIndex
        Case MainTableFields.COL_MAIN_recordStatus
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.recordStatus, VBA.TypeName(Me.recordStatus))
            If Not IsValidField = True Then Exit Function
        Case MainTableFields.COL_MAIN_statusChangeDate
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.statusChangeDate, VBA.TypeName(Me.statusChangeDate))
            If Not IsValidField = True Then Exit Function
        Case MainTableFields.COL_MAIN_customerID
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.customerID, VBA.TypeName(Me.customerID))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = VBA.IsNumeric(Me.customerID)
            If Not IsValidField = True Then Exit Function
            'check 3
            IsValidField = Validator.IsValidStringLength(Me.customerID, 6)
            If Not IsValidField = True Then Exit Function
            'Check 4
            IsValidField = Validator.IsBetweenNumericalRange(Me.customerID, INDEX_USERID_FIRST, INDEX_USERID_LAST)
            If Not IsValidField = True Then Exit Function
        Case MainTableFields.COL_MAIN_materialID
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.materialID, VBA.TypeName(Me.materialID))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = VBA.IsNumeric(Me.materialID)
            If Not IsValidField = True Then Exit Function
            'check 3
            IsValidField = Validator.IsValidStringLength(Me.materialID, 8)
            If Not IsValidField = True Then Exit Function
            'Check 4
            IsValidField = Validator.IsBetweenNumericalRange(Me.materialID, INDEX_MATERIALID_FIRST, INDEX_MATERIALID_LAST)
            If Not IsValidField = True Then Exit Function
        Case MainTableFields.COL_MAIN_price
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.price, VBA.TypeName(Me.price))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = VBA.IsNumeric(Me.price)
            If Not IsValidField = True Then Exit Function
            'Check 3
            IsValidField = Validator.IsBetweenNumericalRange(Me.price, MIN_PRICE_VALUE, MAX_PRICE_VALUE)
            If Not IsValidField = True Then Exit Function
        Case MainTableFields.COL_MAIN_currency
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.currencyType, VBA.TypeName(Me.currencyType))
            If Not IsValidField = True Then Exit Function
        Case MainTableFields.COL_MAIN_unitOfPrice
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.unitOfPrice, VBA.TypeName(Me.unitOfPrice))
            If IsValidField = False Then
                IsValidField = True
                Exit Function
            End If
            'Check 2
            IsValidField = VBA.IsNumeric(Me.unitOfPrice)
            If Not IsValidField = True Then Exit Function
            'Check 3
            IsValidField = Validator.IsValidStringLength(Me.unitOfPrice, 4)
            If Not IsValidField = True Then Exit Function
            'Check 4
            IsValidField = Validator.IsBetweenNumericalRange(Me.unitOfPrice, MIN_UNITOFPRICE_VALUE, MAX_UNITOFPRICE_VALUE)
            If Not IsValidField = True Then Exit Function
        Case MainTableFields.COL_MAIN_unitOfMeasure
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.unitOfMeasure, VBA.TypeName(Me.unitOfMeasure))
            If Not IsValidField = True Then Exit Function
        Case MainTableFields.COL_MAIN_validFromDate
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.validFromDate, VBA.TypeName(Me.validFromDate))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = Validator.IsValidGermanFormatDate(Me.validFromDate)
            If Not IsValidField = True Then Exit Function
            'Check 3
            startDate = VBA.Format(VBA.Now, DATEFORMAT_BACKEND)
            endDate = startDate
            IsValidField = Validator.IsValidDateInRange(Me.validFromDate, startDate, endDate)
            If Not IsValidField = True Then Exit Function
        Case MainTableFields.COL_MAIN_validToDate
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.validToDate, VBA.TypeName(Me.validToDate))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = Validator.IsValidGermanFormatDate(Me.validToDate)
            If Not IsValidField = True Then Exit Function
            'Check 3
            startDate = VBA.Format(VBA.Now, DATEFORMAT_BACKEND)
            endDate = VBA.CDate("31-Dec-9999")
            IsValidField = Validator.IsValidDateInRange(Me.validToDate, startDate, endDate)
            If Not IsValidField = True Then Exit Function
    End Select
    'If All Goes Well!
    this.CheckFields(FieldIndex) = True
End Function

'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Validator = Nothing
End Sub

