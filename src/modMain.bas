Attribute VB_Name = "modMain"
'@Folder("Main")
Option Explicit

Public Sub MainPAM()

    On Error GoTo CleanFail:
    With ProgressIndicator.Create("InitilaizeApplication")
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


Private Sub InitilaizeApplication(ByVal SPLASH As ProgressIndicator)
    
    'Object Declaration
    Dim Presenter           As AppPresenter
    Dim RepositoryInUse     As RepositoryType
    
    'Initialize App
    Set Presenter = New AppPresenter
    
    SPLASH.Update 10, "Loading Repository..."
    
    'Switch Database type from here
    RepositoryInUse = TYPE_POSTGRESQL
    
    'Configure Presenter and Attach Important Datasets with Application Components
    With Presenter
    
        'Splash Screen Stage : Checking if Tables are accessible or not?
        SPLASH.Update 20, "Validating Data Sources..."
        
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
        SPLASH.Update 50, "Loading Data..."
        
        'Configure Application Model with Important DataSet
        Call .InItApplicationModel(modDataSources.arrListofCurrencies, _
                                    modDataSources.arrListOfUnitOfMeasure, _
                                    modDataSources.arrListofTypesOfUser, _
                                    modDataSources.arrListofStatusOfUser, _
                                    modDataSources.arrRecordStatusesList, _
                                    modDataSources.arrSalesOrganizationsList, _
                                    modDataSources.arrDistributionChannelsList)
        
        'Splash Screen Stage : Final
        SPLASH.Update 85, "Opening App..."
        Call WaitForOneSecond
        SPLASH.Update 100, "Status: Ok"
        
        'Splash Screen Exit
        SPLASH.CloseScreen
        
        'Attach and Configure VIEW with Application
        Call .InItApp
        
    End With
    
    'Exiting from Application!
    Set Presenter = Nothing
    Exit Sub
    
CleanExit:
    
    'Splash Screen Exit
    SPLASH.CloseScreen
    
    'Exiting from Application!
    Set Presenter = Nothing
    
End Sub


