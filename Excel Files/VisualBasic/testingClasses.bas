Attribute VB_Name = "testingClasses"
Sub testing()
    ' Test data master worksheet
    Dim ws As Worksheet: Set ws = ActiveWorkbook.Sheets("Test Data")
    
    ' Copy the master worksheet for testing purposes
    ws.Copy After:=Worksheets(ws.Name)
    Set ws = ActiveWorkbook.Sheets(Worksheets.count - 1)
    
    ' Process raw data
    Dim readRawData As New RawDataIO
    Dim runBatches As New Collection
    
    Set runBatches = readRawData.Batches(ws)
    MsgBox "I'm at the end of the testing sub"
    
    
    ' Delete the copied master worksheet and activate the actual master sheet
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    ActiveWorkbook.Sheets("Test Data").Activate
    
    ' ********************************************************************
    ' Current unused/testing code
    ' ********************************************************************
    Dim a As New DevoAssay
    Dim frm As New frmDevoAssayGUI
    'frm.Show
    'If frm.Cancelled = True Then
    '    MsgBox "Cancelled"
    'Else
    '    MsgBox "Sheet selected: " & frm.RawDataWorksheet
    'End If
    
    ' Clean up
    'Unload frm
    'Set frm = Nothing
   
    'a.Init "BD001", Array("PHA", "IL-2", "SAHA"), "CD8-Depleted Targets"
    'a.printPheresis
End Sub


