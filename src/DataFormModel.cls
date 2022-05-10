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
    'Validation Object
    Validator As ValidationServices
    'Tables
    dataTable As Variant
    'Fields
    index As Long
    RecordID As String
    userID As String
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
    'Attributes
    ActiveDataContainer As Long
    IsApprover As Boolean
End Type

Private this As TDataFormModelComponents

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
'Init Method
'-------------------------------------------------------------------------

Public Sub InItModel(ByVal mainData As Variant)
    If Validator Is Nothing Then Set Validator = New ValidationServices
    Me.dataTable = mainData
End Sub

'-------------------------------------------------------------------------
'Validations
'-------------------------------------------------------------------------

Public Function IsValidRecordToEdit() As Variant
    'Variable declaration
    Dim recordStatus    As String
    Dim TargetIndex     As Long
    Dim firstRow        As Long
    Dim lastRow         As Long
    Dim I               As Long
    'Calculate driving paramters
    TargetIndex = Me.index - 1
    firstRow = LBound(Me.dataTable, 1)
    lastRow = UBound(Me.dataTable, 1)
    'Finding Status!
    For I = firstRow + 1 To lastRow
        If Me.dataTable(I, MainTableFields.COL_MAIN_INDEX) = TargetIndex Then
            recordStatus = Me.dataTable(I, MainTableFields.COL_MAIN_recordStatus)
            Exit For
        End If
    Next I
    'Decision
    Select Case recordStatus
        Case RECORDSTATUS_PENDING
            IsValidRecordToEdit = True
        Case RECORDSTATUS_APPROVED
            IsValidRecordToEdit = "Your record has been already Approved! You have no permission to edit it further!"
        Case RECORDSTATUS_REJECTED
            IsValidRecordToEdit = "Your record has been already Rejected! Kindly submit new record for Approval!"
    End Select
    'Overriding
    If IsApprover Then IsValidRecordToEdit = True
End Function

'-------------------------------------------------------------------------
'Class Event
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Validator = Nothing
End Sub
