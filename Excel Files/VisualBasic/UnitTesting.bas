Attribute VB_Name = "UnitTesting"
'@TestModule
Private Assert As New Rubberduck.AssertClass
Private Fakes As New FakesProvider
Private d As New DevoAssay


'@TestMethod
Public Sub test()
    Assert.AreEqual 2, 2
End Sub

'@TestMethod
Public Sub testPheresisName()
    'Dim userInput As String
    'With Fakes.InputBox
    '    .Returns vbNullString, 1
    '    .ReturnsWhen "Prompt", "Second", "User entry 2", 2
    '    userInput = InputBox("First")
    '    Assert.IsTrue userInput = vbNullString
    '    userInput = InputBox("Second")
    '    Assert.IsTrue userInput = "User entry 2"
    'End With
    
    d.Pheresis = "PH32"
    Assert.AreEqual "PH321", d.Pheresis
End Sub
