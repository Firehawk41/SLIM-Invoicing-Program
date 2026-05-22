VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceSampleRowReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceSampleRowReg
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-16
' Description: Specialized object for regular Sample Rows
' DependsOn: IInvoiceSampleRow, clsLoggingSystem, clsInvoiceSampleRowBase, clsTestingRequestSubmission, clsTestingRequestSample
' ChangeLog:
'   - 1.0.1 - 2025-06-16 - Added metadata
Implements IInvoiceSampleRow
Private m_Logger As clsLoggingSystem
Private m_SampleRowBase As clsInvoiceSampleRowBase
Public Sub IInvoiceSampleRow_Initialize(Submission As clsTestingRequestSubmission, Sample As clsTestingRequestSample, TextColour As Long, Logger As clsLoggingSystem)
    
    Call Logger.LogMessage("clsInvoiceSampleRowReg.Initialize", LogLevelEnum.LogDebug, "Initializing InvoiceSampleRowReg")
    
    Set m_Logger = Logger

    On Error GoTo ErrorHandler
    
    Call m_SampleRowBase.Initialize(Submission, Sample, TextColour, Logger)

    Call m_Logger.LogMessage("Initialize", LogLevelEnum.LogInfo, "Successfully initialized InvoiceSampleRowReg")
    Exit Sub
    
ErrorHandler:
    Call m_Logger.LogError("InitializeSampleTableRow", Err.Number, Err.Description)
End Sub

Public Function IInvoiceSampleRow_WriteSampleRow(Table As Word.Table, LineNumber As Long)
    Call m_SampleRowBase.WriteSampleRow(Table, LineNumber)
End Function

Private Sub Class_Initialize()
    Set m_SampleRowBase = New clsInvoiceSampleRowBase
End Sub

Private Sub Class_Terminate()
    Set m_SampleRowBase = Nothing
End Sub
