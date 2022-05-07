Attribute VB_Name = "modMain"
Option Explicit

Public Sub MainPAM()

    Dim RepositoryInUse As RepositoryType
    Dim presenter As AppPresenter
    
    Set presenter = New AppPresenter
    'RepositoryInUse = TYPE_EXCEL_NAMED_RANGE
    RepositoryInUse = TYPE_ACCESS
    
    'Configure Presenter and Attach Important Datasets with Application Components
    With presenter
        
        'Attach main table of the database with Application and configure Related Services Object
        Call .InItMainService(RepositoryInUse, _
                                modDataSources.MAIN_TABLE_NAME, _
                                modDataSources.arrListOfColumns_MAIN_Table)
                                
        'Check if Database is connected or not? if not then do not open app!
        If .databaseConnectionStatus = True Then
                                
            'Attach users table of the database with application and configure Related Services Object
            Call .InItUserService(RepositoryInUse, _
                                    modDataSources.USERS_TABLE_NAME, _
                                    modDataSources.arrListOfColumns_USERS_TABLE)
                                        
            'Check if Database is connected or not? if not then do not open app!
            If .databaseConnectionStatus = True Then
                                    
                'Configure Application Model with Important DataSet
                Call .InItApplicationModel(modDataSources.arrListofCurrencies, _
                                            modDataSources.arrListOfUnitOfMeasure, _
                                            modDataSources.arrListofTypesOfUser, _
                                            modDataSources.arrListofStatusOfUser, _
                                            modDataSources.arrRecordStatusesList)
                                        
                'Attach and Configure VIEW with Application
                Call .InItApp
            
            Else
                MsgBox "Having an issue connecting Users Table!", vbCritical, SIGN
            End If
        
        Else
            MsgBox "Having an issue connecting Main Table!", vbCritical, SIGN
        End If
        
    End With
    
    'Exiting from Application!
    Set presenter = Nothing
    
End Sub
