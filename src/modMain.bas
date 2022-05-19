Attribute VB_Name = "modMain"
Option Explicit

Public Sub MainPAM()
    
    Dim SplashScreen As SPLASH
    Dim RepositoryInUse As RepositoryType
    Dim Presenter As AppPresenter
    
    Set SplashScreen = New SPLASH
    Set Presenter = New AppPresenter
    
    SplashScreen.Show vbModeless
    SplashScreen.lblMessage.Caption = "Loading Repository..."
    Call WaitForOneSecond
    
    'Switch Database type from here
    RepositoryInUse = TYPE_POSTGRESQL
    
    'Configure Presenter and Attach Important Datasets with Application Components
    With Presenter
    
        SplashScreen.lblMessage.Caption = "Validating Data Sources..."
        Call WaitForOneSecond
        
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
        
        SplashScreen.lblMessage.Caption = "Loading Data..."
        Call WaitForOneSecond
        
        'Configure Application Model with Important DataSet
        Call .InItApplicationModel(modDataSources.arrListofCurrencies, _
                                    modDataSources.arrListOfUnitOfMeasure, _
                                    modDataSources.arrListofTypesOfUser, _
                                    modDataSources.arrListofStatusOfUser, _
                                    modDataSources.arrRecordStatusesList, _
                                    modDataSources.arrSalesOrganizationsList, _
                                    modDataSources.arrDistributionChannelsList)
        
        SplashScreen.lblMessage.Caption = "Opening App..."
        Call WaitForOneSecond
        Call WaitForOneSecond
        
        SplashScreen.Hide
        Set SplashScreen = Nothing
        
        'Attach and Configure VIEW with Application
        Call .InItApp
        
    End With
    
CleanExit:
    
    'Exiting from Application!
    Set Presenter = Nothing
    
End Sub
