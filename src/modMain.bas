Attribute VB_Name = "modMain"
'@Folder("Main")
Option Explicit

'As we want our Main View to be VbModeless, We have to take out our main driving object _
which is nothing but the Presenter! Yes The whole Application is dependent on the scope _
of Presenter object. So, What happens when Form becomes vbmodeless, it simply allow _
compiler to run next steps. So to preven Presenter to go out of scope, we have to take _
out the object defination from the Mehtod and keep it as Public Object.

Public Presenter As AppPresenter

Public Sub MainPAM()

    On Error GoTo CleanFail:
    With ProgressIndicator.Create("InitializeApplication", CanCancel:=True)
        .Execute
    End With

CleanExit:
    Exit Sub
    
CleanFail:
    MsgBox VBA.Err.Description, Title:=VBA.Err.Number
    LogManager.Log ErrorLevel, "Error: " & VBA.Err.Number & ". " & VBA.Err.Description
    Resume CleanExit
    Resume
    
End Sub

Private Sub InitializeApplication(ByVal Progress As ProgressIndicator)

    'Object Declaration
    Dim RepositoryInUse     As RepositoryType
    
    'Initialize App
    Set Presenter = New AppPresenter
    
    Progress.Update 10, "Loading Repository..."

    'Switch Database type from here
    RepositoryInUse = TYPE_POSTGRESQL
    
    'Configure Presenter and Attach Important Datasets with Application Components
    With Presenter
    
        'Splash Screen Stage : Checking if Tables are accessible or not?
        Progress.Update 30, "Validating Data Sources..."
        
        'Attach main table of the database with Application and configure Related Services Object
        Call .InItMainService(RepositoryInUse, _
                              modDataSources.MAIN_TABLE_NAME, _
                              modDataSources.arrListOfColumns_MAIN_Table, _
                              modDataSources.GetConnectionString(RepositoryInUse))
                                
        'Check if Database is connected or not? if not then do not open app!
        If .databaseConnectionStatus = False Then GoTo CleanExit
                                
        'Attach users table of the database with application and configure Related Services Object
        Call .InItUserService(RepositoryInUse, _
                              modDataSources.USERS_TABLE_NAME, _
                              modDataSources.arrListOfColumns_USERS_TABLE, _
                              modDataSources.GetConnectionString(RepositoryInUse))
             
        'Check if Database is connected or not? if not then do not open app!
        If .databaseConnectionStatus = False Then GoTo CleanExit
        
        'Splash Screen Stage : Loading Data To App Model
        Progress.Update 50, "Loading Data..."
        
        'Configure Application Model with Important DataSet
        Call .InItApplicationModel(modDataSources.arrListofCurrencies, _
                                    modDataSources.arrListOfUnitOfMeasure, _
                                    modDataSources.arrListofTypesOfUser, _
                                    modDataSources.arrListofStatusOfUser, _
                                    modDataSources.arrRecordStatusesList, _
                                    modDataSources.arrSalesOrganizationsList, _
                                    modDataSources.arrDistributionChannelsList)
        
        'Splash Screen Stage : Final
        Progress.Update 80, "Opening App..."
        Call WaitForOneSecond
        Progress.Update 100, "Status: Ok"
        
        'Splash Screen Exit
        Progress.CloseScreen
        
        'Attach and Configure VIEW with Application
        Call .InItApp
        
    End With

CleanExit:
    'Splash Screen Exit
    Progress.CloseScreen
    
End Sub
