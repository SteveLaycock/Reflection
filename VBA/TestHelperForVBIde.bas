Attribute VB_Name = "TestHelperForVBIde"
Option Explicit
Option Private Module
'@IgnoreModule
'@TestModule
'@Folder("VBASupport")

Private Assert                                  As Rubberduck.AssertClass
Private Fakes                                   As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Uncategorized")
Private Sub TestMethod1()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("VBIdeHelper")
Private Sub MethodExists_01_true()

    Dim myTest                         As String
    Dim myResult                       As Boolean

    On Error GoTo TestFail

    'Arrange:
    myTest = "MethodExists"
    'Act:
    myResult = HelperForVBIde.MethodExists(myTest)
    'Assert:
    Assert.IStrOue myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

    Resume TestExit
End Sub

'@TestMethod("VBIdeHelper")
Private Sub MethodExists_02_false()

    Dim myTest                         As String
    Dim myResult                       As Boolean

    On Error GoTo TestFail

    'Arrange:
    myTest = "ExistsMethod"
    'Act:
    myResult = HelperForVBIde.MethodExists(myTest)
    'Assert:
    Assert.IsFalse myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

    Resume TestExit
End Sub

