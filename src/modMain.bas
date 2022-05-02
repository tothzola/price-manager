Attribute VB_Name = "modMain"
Option Explicit

Public Sub MainPAM()

    Dim presenter As AppPresenter
    Set presenter = New AppPresenter
    
    'Configure Presenter and Attach Important Datasets with Application Components
    With presenter
    
        'Attach main table of the database with Application and configure Related Services Object
        Call .InItMainService(TYPE_EXCEL_NAMED_RANGE, _
                                modDataSources.MAIN_TABLE_NAME)
                                
        'Attach users table of the database with application and configure Related Services Object
        Call .InItUserService(TYPE_EXCEL_NAMED_RANGE, _
                                modDataSources.USERS_TABLE_NAME)
                                
        'Configure Application Model with Important DataSet
        Call .InItApplicationModel(modDataSources.arrListofCurrencies, _
                                    modDataSources.arrListOfUnitOfMeasure, _
                                    modDataSources.arrListofTypesOfUser, _
                                    modDataSources.arrListofStatusOfUser, _
                                    modDataSources.arrRecordStatusesList)
                                
        'Attach and Configure VIEW with Application
        Call .InItApp
        
    End With
    
    'Exiting from Application!
    Set presenter = Nothing
    
End Sub


Sub test()

    Debug.Print ThisWorkbook.Worksheets.Count
    
End Sub
