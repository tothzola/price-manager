Attribute VB_Name = "TestStringBuilder"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.StringFormating")

Private Assert As Object

Private Type TState
    ConcreteSUT As StringBuilder
    
End Type

Private Test As TState

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")

End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing

End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set Test.ConcreteSUT = New StringBuilder
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Test.ConcreteSUT = Nothing
End Sub

'@TestMethod("StringFormat")
Private Sub TestCurrencyFormatUS()
    
    Dim result As String
    result = Test.ConcreteSUT.AppendFormat("{0:C2}", 123.456).ToString
    
    Assert.AreEqual "$123.46", result

End Sub

'@TestMethod("StringFormat")
Private Sub TestDigitFormat()

    Dim result As String
    result = Test.ConcreteSUT.AppendFormat("{0:D}", 12345).ToString
    
    Assert.AreEqual "12345", result

End Sub

'@TestMethod("StringFormat")
Private Sub Test8DigitFormat()

    Dim result As String
    result = Test.ConcreteSUT.AppendFormat("{0:D8}", 12345).ToString
    
    Assert.AreEqual "00012345", result

End Sub

'@TestMethod("StringFormat")
Private Sub TestPercent2DigitFormat()

    Dim result As String
    result = Test.ConcreteSUT.AppendFormat("{0:P}", 0.2468013).ToString
    
    Assert.AreEqual "24.68 %", result

End Sub

'@TestMethod("StringFormat")
Private Sub TestPercent8DigitFormat()

    Dim result As String
    result = Test.ConcreteSUT.AppendFormat("{0:P8}", 0.2468013).ToString
    
    Assert.AreEqual "24.68013000 %", result

End Sub

'@TestMethod("StringFormat")
Private Sub TestDayNameFormat()
    
    Dim result As String
    result = Test.ConcreteSUT.AppendFormat("{0:dddd}", VBA.Date).ToString
    
    Assert.AreEqual VBA.Format$(VBA.Now, "dddd"), result
    
End Sub

'@TestMethod("StringFormat")
Private Sub TestDayNumberFormat()

    Dim result As String
    result = Test.ConcreteSUT.AppendFormat("{0:dd}", VBA.Date).ToString
    
    Assert.AreEqual VBA.Mid$(VBA.Now, 3, 2), result
    
End Sub

'@TestMethod("StringFormat")
Private Sub TestMonthNameFormat()
    
    Dim result As String
    result = Test.ConcreteSUT.AppendFormat("{0:MMMM}", VBA.Date).ToString
    
    Assert.AreEqual VBA.Format$(VBA.Now, "MMMM"), result

End Sub

'@TestMethod("StringFormat")
Private Sub TestMonthNumberFormat()

    Dim result As String
    result = Test.ConcreteSUT.AppendFormat("{0:MM}", VBA.Date).ToString
    
    Assert.AreEqual VBA.Format(Now, "MM"), result
    
End Sub

'@TestMethod("StringFormat")
Private Sub TestParamArrayFormat()

    Dim result As String
    result = Test.ConcreteSUT.AppendFormat( _
    "Jack {0} Jill {1} up {2} hill {3} fetch {4} pail {5} water {6} fell {7} And {8} his {9} And {10} came {11} after" _
    , "and", "Went", "the", "To", "a", "of", "Jack", "down", "broke", "crown,", "Jill", "tumbling").ToString
    
    Assert.AreEqual "Jack and Jill Went up the hill To fetch a pail of water Jack fell down And broke his crown, And Jill came tumbling after", result

End Sub

'@TestMethod("StringFormat")
Private Sub TestArrayFormat()

    Dim result As String
    result = Test.ConcreteSUT.AppendFormat( _
    "Jack {0} Jill {1} up {2} hill {3} fetch {4} pail {5} water {6} fell {7} And {8} his {9} And {10} came {11} after", _
    Array("and", "Went", "the", "To", "a", "of", "Jack", "down", "broke", "crown,", "Jill", "tumbling")).ToString
    
    Assert.AreEqual "Jack and Jill Went up the hill To fetch a pail of water Jack fell down And broke his crown, And Jill came tumbling after", result

End Sub


