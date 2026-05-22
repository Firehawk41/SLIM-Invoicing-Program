VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceSampleRowManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceSampleRowManager
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-16
' Description: Manager object for sample rows
' DependsOn: IInvoiceSampleRow, clsLoggingSystem, clsInvoiceFactory
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Private m_Logger As clsLoggingSystem
Private m_SampleRows As Collection

Public Sub Initialize(Submissions As Collection, Logger As clsLoggingSystem, Factory As clsInvoiceFactory)
    Call Logger.LogMessage("Initialize", LogLevelEnum.LogDebug, "Initializing InvoiceSampleRowManager")
    
    Set m_Logger = Logger
    Set m_SampleRows = Factory.SetSampleRowCollection(Submissions)
    
    Call Logger.LogMessage("Initialize", LogLevelEnum.LogInfo, "Successfully initiated InvoiceSampleRowManager")
    
End Sub
Public Sub FillInSampleTable(Table As Word.Table)
    Dim SampleRow As IInvoiceSampleRow
    Dim LineNumber As Long
    
    Call m_Logger.LogMessage("FillInSampleRowTable", LogLevelEnum.LogDebug, "Initiating filling in the Sample Table")
    
    LineNumber = 1
    
    For Each SampleRow In m_SampleRows
        
        If Table.Rows.Count < LineNumber Then
            Table.Rows.Add
        End If
        
        Call SampleRow.WriteSampleRow(Table, LineNumber)
        LineNumber = LineNumber + 1
    Next SampleRow
    
    Call m_Logger.LogMessage("FillInSampleRowTable", LogLevelEnum.LogInfo, "Successfully filled in the Sample Table")
End Sub

' Class constructor
Private Sub Class_Initialize()
    ' Initialize the sample rows collection when the class is instantiated.
    Set m_SampleRows = New Collection
End Sub

' Class destructor
Private Sub Class_Terminate()
    Dim SampleRow As IInvoiceSampleRow
    
    ' Ensure the collection is not empty before attempting cleanup.
    If Not m_SampleRows Is Nothing Then
        For Each SampleRow In m_SampleRows
            ' Set each sample row to Nothing to release memory.
            Set SampleRow = Nothing
        Next SampleRow
        ' set the collection to nothing to release memory
        Set m_SampleRows = Nothing
    End If
End Sub
