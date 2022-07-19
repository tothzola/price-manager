VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressIndicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule ProcedureCanBeWrittenAsFunction
'@Folder("Progress Indicator")
'@PredeclaredId
'@Exposed

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Private Const DEFAULT_CAPTION As String = "Progress"
Private Const DEFAULT_LABEL As String = "Please wait..."
Private Const ERR_NOT_INITIALIZED As String = "ProgressIndicator is not initialized."
Private Const ERR_PROC_NOT_FOUND As String = "Specified macro or object member was not found."
Private Const ERR_INVALID_OPERATION As String = "Worker procedure cannot be cancelled by assigning to this property."
Private Const VBERR_MEMBER_NOT_FOUND As Long = 438

Public Enum ProgressIndicatorError
    Error_NotInitialized = vbObjectError + 1001
    Error_ProcedureNotFound
    Error_InvalidOperation
End Enum

Private Type TProgressIndicator
    Procedure As String
    Parameter As Variant
    Instance As Object
    SleepDelay As Long
    CanCancel As Boolean
    CanCelling As Boolean
    CurrentProgressValue As Double
End Type

Private this As TProgressIndicator
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents View As SplashView
Attribute View.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set View = New SplashView
    View.Caption = DEFAULT_CAPTION
    View.ProgressLabel = DEFAULT_LABEL
End Sub

Private Sub Class_Terminate()
    Set View = Nothing
    Set this.Instance = Nothing
End Sub

Private Function QualifyMacroName(ByVal Book As Workbook, ByVal Procedure As String) As String
    QualifyMacroName = "'" & Book.FullName & "'!" & Procedure
End Function

Public Function Create(ByVal Procedure As String, _
Optional ByVal Parameter As Variant = Empty, _
Optional ByRef Instance As Object = Nothing, _
Optional ByVal InitialLabelValue As String, _
Optional ByVal InitialCaptionValue As String, _
Optional ByVal CompletedSleepMilliseconds As Long = 1000, _
Optional ByRef CanCancel As Boolean = False) As ProgressIndicator
    
    Dim result As ProgressIndicator
    Set result = New ProgressIndicator

    result.Parameter = Parameter
    result.CanCellable = CanCancel
    result.SleepMilliseconds = CompletedSleepMilliseconds
 
    If Not Instance Is Nothing Then
        Set result.OwnerInstance = Instance
    ElseIf InStr(Procedure, "'!") = 0 Then
        Procedure = QualifyMacroName(ThisWorkbook, Procedure)
    End If
 
    result.ProcedureName = Procedure
 
    If InitialLabelValue <> vbNullString Then result.ProgressView.ProgressLabel = InitialLabelValue
    If InitialCaptionValue <> vbNullString Then result.ProgressView.ProgressBar.Caption = InitialCaptionValue
        
    Set Create = result
 
End Function

Friend Property Get ProgressView() As SplashView
    Set ProgressView = View
End Property

Friend Property Get ProcedureName() As String
    ProcedureName = this.Procedure
End Property

Friend Property Let ProcedureName(ByVal Value As String)
    this.Procedure = Value
End Property

Friend Property Get Parameter() As String
    Parameter = this.Parameter
End Property

Friend Property Let Parameter(ByVal Value As String)
    this.Parameter = Value
End Property

Friend Property Get OwnerInstance() As Object
    Set OwnerInstance = this.Instance
End Property

Friend Property Set OwnerInstance(ByVal Value As Object)
    Set this.Instance = Value
End Property

Friend Property Get SleepMilliseconds() As Long
    SleepMilliseconds = this.SleepDelay
End Property

Friend Property Let SleepMilliseconds(ByVal Value As Long)
    this.SleepDelay = Value
End Property

Public Property Get CurrentProgress() As Double
    CurrentProgress = this.CurrentProgressValue
End Property

Public Property Get CanCellable() As Boolean
    CanCellable = this.CanCancel
End Property

Friend Property Let CanCellable(ByVal Value As Boolean)
    this.CanCancel = Value
End Property

Public Property Get IsCancelRequested() As Boolean
    IsCancelRequested = this.CanCelling
End Property

Public Sub AbortCancellation()
    Debug.Assert this.CanCelling
    this.CanCelling = False
End Sub

Public Sub Execute()
    View.Show vbModal
End Sub

Public Sub CloseScreen()
    View.Hide
End Sub

Public Sub Update(ByVal percentValue As Double, Optional ByVal labelValue As String, Optional ByVal captionValue As String)

    On Error GoTo CleanFail
    ThrowIfNotInitialized

    ValidatePercentValue percentValue
    this.CurrentProgressValue = percentValue
 
    View.Update this.CurrentProgressValue, labelValue

CleanExit:
    If percentValue = 1 Then Sleep 1000          ' pause on completion
    Exit Sub

CleanFail:
    MsgBox Err.Number & vbTab & Err.Description, vbCritical, PriceApprovalSignature
    LogManager.Log ErrorLevel, "Error: " & Err.Number & vbTab & Err.Description
    Resume CleanExit
    Resume
    
End Sub

Public Sub UpdatePercent(ByVal percentValue As Double, Optional ByVal captionValue As String)
    ValidatePercentValue percentValue
    Update percentValue, Format$(percentValue, "0.0% Completed")
End Sub

Private Sub ValidatePercentValue(ByRef percentValue As Double)
    If percentValue > 1 Then percentValue = percentValue / 100
End Sub

Private Sub ThrowIfNotInitialized()
    If this.Procedure = vbNullString Then
        Err.Raise ProgressIndicatorError.Error_NotInitialized, TypeName(Me), ERR_NOT_INITIALIZED
        LogManager.Log ErrorLevel, "Error: " & ProgressIndicatorError.Error_NotInitialized & ". " & ERR_NOT_INITIALIZED
    End If
End Sub

Private Sub View_Activated()
 
    On Error GoTo CleanFail
    ThrowIfNotInitialized
 
    If Not this.Instance Is Nothing Then
        ExecuteInstanceMethod
    Else
        ExecuteMacro this.Parameter
    End If
 
CleanExit:
    View.Hide
    Exit Sub
 
CleanFail:
    MsgBox Err.Number & vbTab & Err.Description, vbCritical, PriceApprovalSignature
    LogManager.Log ErrorLevel, "Error: " & Err.Number & vbTab & Err.Description
    Resume CleanExit
    Resume
    
End Sub

Private Sub Cancelled()
    If Not this.CanCancel Then Exit Sub
    this.CanCelling = True
End Sub

Private Sub ExecuteMacro(Optional ByVal Param As Variant)

    On Error GoTo CleanFail
    Dim App As Excel.Application
    Set App = Excel.Application
    
    If Param = Empty Then
        App.Run this.Procedure, Me
    Else
        App.Run this.Procedure, Me, Param
    End If
CleanExit:
    Exit Sub

CleanFail:
    If Err.Number = VBERR_MEMBER_NOT_FOUND Then
        Err.Raise ProgressIndicatorError.Error_ProcedureNotFound, TypeName(Me), ERR_PROC_NOT_FOUND
        LogManager.Log ErrorLevel, "Error: " & ProgressIndicatorError.Error_ProcedureNotFound & ". " & ERR_PROC_NOT_FOUND
    Else
        Err.Raise VBA.Err.Number, VBA.Err.Source, VBA.Err.Description, VBA.Err.HelpFile, VBA.Err.HelpContext
        LogManager.Log ErrorLevel, "Error: " & VBA.Err.Number & vbTab & VBA.Err.Description
    End If
    Resume CleanExit
    Resume
    
End Sub

Private Sub ExecuteInstanceMethod()
    On Error GoTo CleanFail
 
    Dim Param As ProgressIndicator
    Set Param = Me 'Me cannot be passed to CallByName directly
    
    '@Ignore ValueRequired
    VBA.CallByName this.Instance, this.Procedure, VbMethod, Param

CleanExit:
    Exit Sub

CleanFail:
    If Err.Number = VBERR_MEMBER_NOT_FOUND Then
        Err.Raise ProgressIndicatorError.Error_ProcedureNotFound, TypeName(Me), ERR_PROC_NOT_FOUND
        LogManager.Log ErrorLevel, "Error: " & ProgressIndicatorError.Error_ProcedureNotFound & ". " & ERR_PROC_NOT_FOUND
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
        LogManager.Log ErrorLevel, "Error: " & VBA.Err.Number & vbTab & VBA.Err.Description
    End If
    Resume CleanExit
    Resume
    
End Sub

Private Sub View_Cancelled()
    If Not this.CanCancel Then Exit Sub
    this.CanCelling = True
End Sub


