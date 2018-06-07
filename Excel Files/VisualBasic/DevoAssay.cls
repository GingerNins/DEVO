VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DevoAssay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: DevoAssay
' ************************************************************************

' Properties
' ************************************************************************

' Pheresis or Blood Draw
Private sPheresis As String
' Overall MAD for all data in assay
Private dAssayMAD As Double
' All LRAs/conditions in assay
Private cConditions As New Collection
' All Runs in assay
Private cAssayRuns As Collection
' Culture volume in microliters (optional)
Private dCultureVolume As Double
' All feeders used in the assay
Private dFeeders As Collection

' Get/Set Properties
' ************************************************************************

' Gets the name of the pheresis or blood draw
Property Get Pheresis() As String
    Pheresis = sPheresis
End Property

' Sets the name of the pheresis or blood draw
' Must be in the format BD### or PH###
Property Let Pheresis(strName As String)
    If strName Like "BD###" Or strName Like "PH###" Then
        sPheresis = strName
    Else
        Pheresis = InputBox(strName & " should be in the format BD### or PH###")
    End If
End Property

' Get the culture volume in microliters
Property Get CultureVolume() As Double
    CultureVolume = dCultureVolume
End Property

' Sets the culture volume in microliters (optional)
Property Let CultureVolume(dblCultureVolume As Double)
    dCultureVolume = dblCultureVolume
End Property

' Gets the list of conditions in the assay
Property Get Conditions() As Collection
    Set Conditions = cConditions
End Property

' Sets the list of conditions in the assay
Property Set Conditions(collConditions As Collection)

End Property


' Methods
' ************************************************************************
' Initializes the assay with the name and the conditions
Public Sub Init(strName As String, arrConditions As Variant, arrFeeders As Variant)
    Pheresis = strName
    
    ' Create condition objects for each condtion in array
    Dim item As Variant
    For Each item In arrConditions
        collConditions.Add item
    Next item
End Sub

Public Sub printPheresis()
    Debug.Print strPheresis
    
    Dim item As Variant
    For Each item In collConditions
        Debug.Print item
    Next item
End Sub

