VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentGeneratorExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVP.Services")
Option Explicit

'This Document Generator Class is designed for the reports where we have to _
and paste whole range object carried within array! This class will help user _
to manage automatically by setting paramters from the calling procedure!

'----------------------------------------------------------------------------
'Settings
'----------------------------------------------------------------------------

Private Const SHEETNAME As String = "REPORT"

Private Type TDocumentGeneratorComponents
    IsReportGeneratedSuccessfully As Boolean
    TargetSheet As Worksheet
    LoadData As Variant
    Headers As Variant
    DateColumns As Variant
    DateFormat As String
    CurrencyColumns As Variant
    CURRENCYFORMAT As String
End Type

Private this As TDocumentGeneratorComponents

'----------------------------------------------------------------------------

Public Property Get IsReportGeneratedSuccessfully() As Boolean
    IsReportGeneratedSuccessfully = this.IsReportGeneratedSuccessfully
End Property

Public Property Let IsReportGeneratedSuccessfully(ByVal vNewValue As Boolean)
    this.IsReportGeneratedSuccessfully = vNewValue
End Property

'----------------------------------------------------------------------------

Private Property Get TargetSheet() As Worksheet
    Set TargetSheet = this.TargetSheet
End Property

Private Property Set TargetSheet(ByVal vNewValue As Worksheet)
    Set this.TargetSheet = vNewValue
End Property

'----------------------------------------------------------------------------

Public Property Get LoadData() As Variant
    LoadData = this.LoadData
End Property

Public Property Let LoadData(ByVal vNewValue As Variant)
    this.LoadData = vNewValue
End Property

'----------------------------------------------------------------------------

Public Property Get Headers() As Variant
    Headers = this.Headers
End Property

Public Property Let Headers(ByVal vNewValue As Variant)
    this.Headers = vNewValue
End Property

'----------------------------------------------------------------------------

Public Property Get DateColumns() As Variant
    DateColumns = this.DateColumns
End Property

Public Property Let DateColumns(ByVal vNewValue As Variant)
    this.DateColumns = vNewValue
End Property

'----------------------------------------------------------------------------

Public Property Get DateFormat() As String
    DateFormat = this.DateFormat
End Property

Public Property Let DateFormat(ByVal vNewValue As String)
    this.DateFormat = vNewValue
End Property

'----------------------------------------------------------------------------

Public Property Get CurrencyColumns() As Variant
    CurrencyColumns = this.CurrencyColumns
End Property

Public Property Let CurrencyColumns(ByVal vNewValue As Variant)
    this.CurrencyColumns = vNewValue
End Property

'----------------------------------------------------------------------------

Public Property Get CURRENCYFORMAT() As String
    CURRENCYFORMAT = this.CURRENCYFORMAT
End Property

Public Property Let CURRENCYFORMAT(ByVal vNewValue As String)
    this.CURRENCYFORMAT = vNewValue
End Property

'----------------------------------------------------------------------------

Public Sub GenerateReport()
    'Object declaration
    Dim ExcelApp As Excel.Application
    Dim NewBook As Excel.Workbook
    'Setting up Report objects
    Set ExcelApp = New Excel.Application
    Set NewBook = ExcelApp.Workbooks.Add
    Set TargetSheet = NewBook.Worksheets(1)
    'Error
    On Error GoTo EhhErr
    TargetSheet.Name = SHEETNAME
    'Pasting the Data
    TargetSheet.Range("A1").Resize(UBound(Me.LoadData, 1), UBound(Me.LoadData, 2)).Value = Me.LoadData
    'Paste Headers if Present
    If VBA.IsArray(Me.Headers) Then GiveHeadings
    'Decorate Report
    Call DecorateReport
    'Update Success Flag
    IsReportGeneratedSuccessfully = True
    'Setup Excel App!
    With ExcelApp
        .Visible = True
        .Application.WindowState = xlMaximized
        TargetSheet.Activate
    End With
    
CleanExit:
    Set TargetSheet = Nothing
    Set NewBook = Nothing
    Set ExcelApp = Nothing
    Exit Sub
    
EhhErr:
    MsgBox Err.Number & vbTab & Err.Description, vbCritical, Title:=SIGN
    LogManager.Log ErrorLevel, "Error: " & Err.Number & vbTab & Err.Description
    'Update Failour Flag
    IsReportGeneratedSuccessfully = False
    Err.Clear
    Resume CleanExit
    Resume
    
End Sub

Private Sub GiveHeadings()
    Dim hCell       As Range
    Dim i As Long
    i = 0
    For Each hCell In TargetSheet.UsedRange.Rows(1).Cells
        hCell.Value = Me.Headers(i)
        i = i + 1
    Next hCell
    Set hCell = Nothing
End Sub

Private Sub DecorateReport()
    With TargetSheet.UsedRange.Columns
        .EntireColumn.AutoFit
        .HorizontalAlignment = xlCenter
    End With
    Call ApplyDateFormatToColumns
    Call ApplyCurrencyFormatToColumns
End Sub

Private Sub ApplyDateFormatToColumns()
    Dim col         As Variant
    With TargetSheet
        For Each col In Me.DateColumns
            .Columns(col).NumberFormat = Me.DateFormat
            '.Columns(col).TextToColumns Destination:=.Columns(col), _
                                                   DataType:=xlDelimited, _
                                                   FieldInfo:=Array(1, xlDMYFormat)
        Next col
    End With
End Sub

Private Sub ApplyCurrencyFormatToColumns()
    Dim col         As Variant
    For Each col In Me.CurrencyColumns
        TargetSheet.Columns(col).Style = Me.CURRENCYFORMAT
    Next col
End Sub
