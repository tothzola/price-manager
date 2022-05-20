Attribute VB_Name = "modMain"
Option Explicit

Private Presenter As AppPresenter

Public Sub MainPAM()
    
    'Object Declaration
    Dim SplashScreen        As SPLASH
    Dim RepositoryInUse     As RepositoryType
    
    'Initialize App
    Set Presenter = New AppPresenter
    Set SplashScreen = New SPLASH
    
    'Splash Screen Stage : Initializing Splash Screen
    SplashScreen.Show vbModeless
    SplashScreen.lblMessage.Caption = "Loading Repository..."
    Call WaitForOneSecond
    
    'Switch Database type from here
    RepositoryInUse = TYPE_POSTGRESQL
    
    'Configure Presenter and Attach Important Datasets with Application Components
    With Presenter
    
        'Splash Screen Stage : Checking if Tables are accessible or not?
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
        
        'Splash Screen Stage : Loading Data To App Model
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
        
        'Splash Screen Stage : Final
        SplashScreen.lblMessage.Caption = "Opening App..."
        Call WaitForOneSecond
        Call WaitForOneSecond
        
        'Splash Screen Exit
        SplashScreen.Hide
        Set SplashScreen = Nothing
        
        'Attach and Configure VIEW with Application
        Call .InItApp
        
    End With
    
    Exit Sub
    
CleanExit:
    
    'Splash Screen Exit
    SplashScreen.Hide
    Set SplashScreen = Nothing
    
End Sub
