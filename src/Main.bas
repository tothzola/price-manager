Attribute VB_Name = "Main"
'@Folder("PAMXLAM")
Option Explicit
Option Private Module

'@ModuleDescription("As we want our Main View to be VbModeless, We have to take out our main driving object _
which is nothing but the Presenter! Yes The whole Application is dependent on the scope _
of Presenter object. So, What happens when Form becomes vbmodeless, it simply allow _
compiler to run next steps. So to preven Presenter to go out of scope, we have to take _
out the object defination from the Mehtod and keep it as Public Object.")

Public Presenter As IAppPresenter

'@EntryPoint
Public Sub StartApp()

    On Error GoTo CleanFail:
    With ProgressIndicator.Create("OpenApplication", CanCancel:=True)
        .Execute
    End With

CleanExit:
    Exit Sub
    
CleanFail:
    MsgBox Err.Number & vbTab & Err.Description, vbCritical, Title:=SIGN
    LogManager.Log ErrorLevel, "Error: " & VBA.Err.Number & ". " & VBA.Err.Description
    Resume CleanExit
    Resume
    
End Sub


Private Sub OpenApplication(ByVal Progress As ProgressIndicator)

    Progress.Update 30, "Application State ..."
    Dim context As IAppContext
    Set context = AppContext.Create
    
    Progress.Update 50, "Validating Data ..."
    If Not context.IsRepositoryReachable Then GoTo CleanExit
    
    Progress.Update 60, "Loading Model ..."
    Dim Model As AppModel
    Set Model = AppModel.Create(context)
    
    Progress.Update 70, "Building View ..."
    Dim View As IView
    Set View = PriceApprovalView.Create(Model)
    
    Progress.Update 80, "Opening App ..."
    Set Presenter = AppPresenter.Create(context, Model, View)
    
    Progress.Update 100, "Application Status = OK"
    GlobalResources.WaitForOneSecond
    Progress.CloseScreen
    
    Presenter.ShowView

CleanExit:
    Progress.CloseScreen

End Sub


