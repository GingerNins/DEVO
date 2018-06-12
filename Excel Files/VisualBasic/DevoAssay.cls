VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DevoAssay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ************************************************************************
' Class Module: DevoAssay
' ************************************************************************

' ************************************************************************
' Properties
' ************************************************************************

' Pheresis or Blood Draw
Private sPheresis As String
' Patient ID
Private sPID As String
' Date of assay setup
Private sDate As String
' All LRAs/conditions in assay
Private cConditions As New Collection
' All targets used in the assay
Private dTargets As Collection
' All dilutions/condition used
Private cDilutions As Collection
' Number of replicates per dilution
Private iReplicates As Integer
' All harvest days in assay
Private cHarvestDays As Collection
' Other factors (such as culture vol) that will affect # of conditions
Private dOtherFactors As Collection
' All Simoa batch runs in assay
Private cBatches As Collection
' Overall MAD for all data in assay
Private dAssayMAD As Double

' Error constant for any invalid data
Private Const ERROR_INVALID_DATA As Long = vbObjectError + 514

' ************************************************************************
' Get/Set Properties
' ************************************************************************

' Gets the name of the pheresis or blood draw
Property Get Pheresis() As String
    Pheresis = sPheresis
End Property

' Sets the name of the pheresis or blood draw
' @param strName Name of the Pheresis or Blood Draw
' @Throws error if not in the format BD### or PH###
Property Let Pheresis(strName As String)
    If strName Like "BD###" Or strName Like "PH###" Then
        sPheresis = strName
    Else
        Err.Raise ERROR_INVALID_DATA, "Let Pheresis", "Pheresis/Blood Draw should be in the format BD### or PH###"
    End If
End Property

' Gets the list of conditions in the assay
Property Get Conditions() As Collection
    Set Conditions = cConditions
End Property

' Sets the list of conditions in the assay
Property Set Conditions(collConditions As Collection)

End Property

' ************************************************************************
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

