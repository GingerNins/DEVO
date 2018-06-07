VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimoaRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: SimoaRun

' Properties
' ************************************************************************
' Name of assay run
Private sBatchName As String
' MAD for the run
Private dMAD As Double
' LLOQ for the run
Private dLLOQ As Double
' Holds replicate concentrations for control 1
Private dQC1(1 To 2) As Double
' Holds replicate concentrations for control 2
Private dQC2(1 To 2) As Double
' List of Carrier Barcodes in assay
' Not sure if I need this (yes I think so... this will be the plates)
Private vCarrierBarcodes As Variant

' Get/Set Properties
' ************************************************************************
Property Get BatchName() As String
    BatchName = sBatchName
End Property

Property Let BatchName(strBatchName As String)
    sBatchName = strBatchName
End Property

'Property Set QualityControls(arrQC1 As Double, arrQC2 As Double)
 '   dQC1 = arrQC1
'    dQC2 = arrQC2
 '
'    Debug.Print dQC1
'    Debug.Print dQC2
'
'End Property

' Methods
' ************************************************************************
' Determines the ranges for the raw data
Sub setRanges(strSheetName As String)
    
End Sub

