Attribute VB_Name = "TestStringBuilder"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

'https://docs.microsoft.com/en-us/dotnet/standard/base-types/standard-numeric-format-strings
Private Assert As Object
Private Fakes As Object

Private Type TState
    ConcreteSUT As StringBuilder
    
End Type

Private Test As TState

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
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
Private Sub TestNumericFormatStrings()
    On Error GoTo TestFail

    'Act:
    Test.ConcreteSUT.AppendLine "Using Standard Numeric Format Strings - Currency (""C"") Format Specifier"
    Test.ConcreteSUT.AppendFormat "{0:C2}", 123.456
    Test.ConcreteSUT.AppendLine
    Test.ConcreteSUT.AppendLine "Currency With Alignment Arguments"
    Test.ConcreteSUT.AppendLine "   Beginning Balance           Ending Balance"
    Test.ConcreteSUT.AppendFormat "   {0,-28:C2}{1,14:C2}", 16305.32, 18794.16
    Test.ConcreteSUT.AppendLine
    Test.ConcreteSUT.AppendLine "The Decimal (""D"") Format Specifier"
    Test.ConcreteSUT.AppendFormat "{0:D}", 12345
    Test.ConcreteSUT.AppendLine
    Test.ConcreteSUT.AppendLine "8 Digit Format Specifier"
    Test.ConcreteSUT.AppendFormat "{0:D8}", 12345
    Test.ConcreteSUT.AppendLine
    Test.ConcreteSUT.AppendLine "The Percent (""P"") Format Specifier"
    Test.ConcreteSUT.AppendFormat "{0:P}", 0.2468013
    Test.ConcreteSUT.AppendLine
    Test.ConcreteSUT.AppendLine "8 Digit Format Specifier"
    Test.ConcreteSUT.AppendFormat "{0:P8}", 0.2468013
    Test.ConcreteSUT.AppendLine
    Test.ConcreteSUT.AppendLine "Custom Tests" & vbNewLine & String(50, "*")
    Test.ConcreteSUT.AppendLine "AppendFormat: Dates"
    Test.ConcreteSUT.AppendFormat "Day ## {0:dd}, Day Name {0:dddd}, Month ## {0:MM}, Month Name {0:MMMM}, YYYY {0:yyyy}", Date
    Test.ConcreteSUT.InsertFormat "Date {0}", 0, 0, "Formats: "
    Test.ConcreteSUT.AppendLine
    Test.ConcreteSUT.AppendLine "AppendFormat: ParamArray"
    Test.ConcreteSUT.AppendFormat "Jack {0} Jill {1} up {2} hill {3} fetch {4} pail {5} water {6} fell {7} And {8} his {9} And {10} came {11} after", "and", "Went", "the", "To", "a", "of", "Jack", "down", "broke", "crown,", "Jill", "tumbling"
    Test.ConcreteSUT.AppendLine
    Test.ConcreteSUT.AppendLine "AppendFormat: Array"
    Test.ConcreteSUT.AppendFormat "Jack {0} Jill {1} up {2} hill {3} fetch {4} pail {5} water {6} fell {7} And {8} his {9} And {10} came {11} after", Array("and", "Went", "the", "To", "a", "of", "Jack", "down", "broke", "crown,", "Jill", "tumbling")
    
    Debug.Print Test.ConcreteSUT.ToString

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    Resume
    
End Sub

