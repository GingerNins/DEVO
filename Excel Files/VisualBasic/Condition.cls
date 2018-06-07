VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Condition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: Condition
' ************************************************************************

' Properties
' ************************************************************************
Private sCondition As String
Private vDilutions As Variant
Private sFeeders As String

' Get/Set Properties
' ************************************************************************

' Get the name of the condition
Property Get Condition() As String
    Condition = sCondition
End Property

' Set the name of the condition
Property Let Condition(strCondition As String)
    sCondition = strCondition
End Property

' Get the Feeders for condition
Property Get Feeders() As String
    Feeders = sFeeders
End Property

' Set the Feeders for the condition
Property Let Feeders(strFeeders As String)
    sFeeders = strFeeders
End Property

' Methods
' ************************************************************************

