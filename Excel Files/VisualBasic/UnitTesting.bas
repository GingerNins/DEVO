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
    d.Pheresis = "PH321"
    Assert.AreEqual "PH321", d.Pheresis
End Sub
