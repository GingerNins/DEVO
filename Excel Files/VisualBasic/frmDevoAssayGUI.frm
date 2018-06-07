VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDevoAssayGUI 
   Caption         =   "DEVO Assay Data Analysis"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7335
   OleObjectBlob   =   "frmDevoAssayGUI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDevoAssayGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' User Form to get input about a specific run
' ************************************************************************

' Properties
' ************************************************************************
Private m_Cancelled As Boolean


' Get/Set Properties
' ************************************************************************
Public Property Get Cancelled() As Variant
    Cancelled = m_Cancelled
End Property

Public Property Get RawDataWorksheet() As String
    RawDataWorksheet = lbWorksheets.Value
End Property

' Methods
' ************************************************************************

' Initialize the user form:
'   - Fill in the sheets list box for the user to select from
Private Sub UserForm_Initialize()
    Dim sh As Worksheet
    For Each sh In ActiveWorkbook.Sheets
        lbWorksheets.AddItem sh.Name
    Next sh
End Sub


' Select ranges for each raw data information
Private Sub cbSelectSampleBarcode_Click()
    Dim rng As Range
    Set rng = Application.InputBox("Select range", "Get Range", Type:=8)
End Sub


' Form Close/Ok handling methods
Private Sub cbOk_Click()
    Hide
End Sub

Private Sub cbCancel_Click()
    Hide
    m_Cancelled = True
End Sub

' Handle user clicking on the X button
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True
    Hide
    m_Cancelled = True
End Sub


