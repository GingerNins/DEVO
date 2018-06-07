VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RawDataIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ************************************************************************
' Class Module: RawData IO
'   Reads raw data from specified worksheet and creates one or more
'   SimoaRuns
' ************************************************************************


' ************************************************************************
' Properties
' ************************************************************************
Private shtRawData As Worksheet
Private rSampleBarcode As Range
Private rReplicateAeb As Range
Private rReplicateConc As Range
Private rBatchName As Range
Private rCarrierBarcode As Range
Private rLocation As Range
Private rErrors As Range

Private vSimoaRun As New SimoaRun

' Constants for header strings found in the raw data sheet
Private Const SAMPLE_BARCODE As String = "Sample Barcode"
Private Const REPLICATE_AEB As String = "Replicate AEB"
Private Const REPLICATE_CONC As String = "Replicate Conc."
Private Const BATCH_NAME As String = "Batch Name"
Private Const CARRIER_BARCODE As String = "Carrier Barcode"
Private Const LOCATION As String = "Location"
Private Const ERRORS As String = "Errors"

' ************************************************************************
' Get/Set/Let Properties
' ************************************************************************
Public Property Get SimoaRunData(shtName As Worksheet) As SimoaRun
    readRawDataSheet shtName
    Set SimoaRunData = vSimoaRun
End Property


' ************************************************************************
' Methods
' ************************************************************************

' Process the raw data on the worksheet
' Main method for this class module
Private Sub readRawDataSheet(shtName As Worksheet)
    Set shtRawData = shtName
    
    ' Remove unecessary rows & store QC information
    
    
    ' Set ranges for raw data
    setRawDataRanges
    
    ' Store QC Information
    setQCConcentrations
    vSimoaRun.BatchName = "Hi I'm a batch name"
    
End Sub

Private Sub setQCConcentrations()
    Dim sample As Range
    
    For Each sample In rSampleBarcode
        If sample.Value = "qc1" Then
            MsgBox sample.Offset(0, 1)
        End If
    Next sample
    
End Sub

Private Sub setRawDataRanges()
    ' Last row in the worksheet
    Dim lRow As Long: lRow = LastRow(shtRawData)
    ' Last column in the worksheet
    Dim lCol As Long: lCol = LastColumn(shtRawData)
    ' Converts last column number into column letter
    Dim lColLetter As String: lColLetter = ColLetter(lCol)
    ' Header row
    Dim rHeaders As Range: Set rHeaders = shtRawData.Range("A1:" & lColLetter & "1")
    ' Temp range variable for looping
    Dim header As Range
    
    Dim headerCol As String
    For Each header In rHeaders
        headerCol = ColLetter(header.Column)
        Select Case header.Value
            Case SAMPLE_BARCODE
                Set rSampleBarcode = shtRawData.Range(headerCol & "2:" & headerCol & lRow)
            Case REPLICATE_AEB
                Set rReplicateAeb = shtRawData.Range(headerCol & "2:" & headerCol & lRow)
            Case REPLICATE_CONC
                Set rReplicateConc = shtRawData.Range(headerCol & "2:" & headerCol & lRow)
            Case BATCH_NAME
                Set rBatchName = shtRawData.Range(headerCol & "2:" & headerCol & lRow)
            Case CARRIER_BARCODE
                Set rCarrierBarcode = shtRawData.Range(headerCol & "2:" & headerCol & lRow)
            Case LOCATION
                Set rLocation = shtRawData.Range(headerCol & "2:" & headerCol & lRow)
            Case ERRORS
                Set rErrors = shtRawData.Range(headerCol & "2:" & headerCol & lRow)
            Case Else
                ' There are columns that will not be used
        End Select
    Next header
End Sub

' Converts the column number to a letter
Function ColLetter(col As Long) As String
    Dim vArr
    vArr = Split(Cells(1, col).Address(True, False), "$")
    ColLetter = vArr(0)
End Function

' Finds the last used column in the raw data worksheet
Function LastColumn(shtName As Worksheet) As Long
    LastColumn = shtName.Cells(1, shtName.Columns.count).End(xlToLeft).Column
End Function

' Finds the last used row in the raw data worksheet
Function LastRow(shtName As Worksheet) As Long
    LastRow = shtName.Cells(shtName.Rows.count, "A").End(xlUp).row
End Function
