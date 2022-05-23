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
    'Tables
    dataTable As Variant
    usersTable As Variant
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
    Dim I               As Long
    'Calculate driving paramters
    TargetIndex = GetTargetRowIndex(Me.dataTable, Me.index, MainTableFields.COL_MAIN_INDEX)
    firstRow = LBound(Me.dataTable, 1)
    lastRow = UBound(Me.dataTable, 1)
    'Finding Status!
    For I = firstRow + 1 To lastRow
        If I = TargetIndex Then
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
            IsValidRecordToEdit = True 'We are allowing client to edit Rejected records!
    End Select
    'Overriding
    If IsApprover Then IsValidRecordToEdit = True
End Function

Public Function GetDataForRecordsList() As Variant
    Dim I       As Long
    Dim J       As Long
    Dim K       As Long
    Dim UB1     As Long
    Dim UB2     As Long
    Dim tmp     As Variant
    Dim uname   As String
    UB1 = UBound(Me.dataTable, 1)
    UB2 = UBound(Me.dataTable, 2)
    ReDim tmp(1 To UB1, 1 To UB2)
    K = 1
    For I = 1 To UB1
        uname = GetLookupValue(Me.usersTable, Me.dataTable(I, MainTableFields.COL_MAIN_userID), UsersTableFields.COL_userId, UsersTableFields.COL_userName)
        If uname <> vbNullString Then
            For J = 1 To UB2
                If I = 1 Then
                    If J = MainTableFields.COL_MAIN_userID Then
                        tmp(K, J) = "User_Name"
                    Else
                        tmp(K, J) = Me.dataTable(I, J)
                    End If
                Else
                    If J = MainTableFields.COL_MAIN_userID Then
                        tmp(K, J) = uname
                    Else
                        tmp(K, J) = Me.dataTable(I, J)
                    End If
                End If
            Next J
            K = K + 1
        End If
    Next I
    GetDataForRecordsList = tmp
End Function

'following function act as Vlookup

Private Function GetLookupValue(ByVal TargetTable As Variant, _
                                ByVal LookupValue As Variant, _
                                ByVal LookupIndex As Long, _
                                ByVal TargetIndex As Long) As String
    Dim I   As Long
    Dim LB  As Long
    Dim UB  As Long
    LB = LBound(TargetTable, 1)
    UB = UBound(TargetTable, 1)
    For I = LB To UB
        If TargetTable(I, LookupIndex) = LookupValue Then
            GetLookupValue = TargetTable(I, TargetIndex)
            Exit Function
        End If
    Next I
End Function

Private Function GetTargetRowIndex(ByVal TargetTable As Variant, ByVal TargetIndex As Long, ByVal TargetColumn As Long) As Long
    Dim I As Long
    For I = LBound(TargetTable, 1) To UBound(TargetTable, 1)
        If TargetTable(I, TargetColumn) = TargetIndex Then
            GetTargetRowIndex = I
            Exit Function
        End If
    Next I
End Function
