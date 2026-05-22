VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceSampleRowCustomerA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceSampleRowCustomerA
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-16
' Description: Specialized object for Customer A Sample Row
' DependsOn: IInvoiceSampleRow, clsLoggingSystem, clsInvoiceSampleRowBase, clsTestingRequestSubmission, clsTestingRequestSample
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Implements IInvoiceSampleRow
Private m_Logger As clsLoggingSystem
Private m_SampleRowBase As clsInvoiceSampleRowBase
Private m_TestingRequestNumber As Variant
Private m_TextColour As Long
Private Sub IInvoiceSampleRow_Initialize(Submission As clsTestingRequestSubmission, Sample As clsTestingRequestSample, TextColour As Long, Logger As clsLoggingSystem)
    Const PRECILAB_BLUE As Long = 16711680
    
    Call Logger.LogMessage("Initialize", LogLevelEnum.LogDebug, "Initializing InvoiceSampleRowCustomerA")
    
    Set m_Logger = Logger
    
    On Error GoTo ErrorHandler
    
    m_TestingRequestNumber = Submission.POInformation
    ' Set the text colour as alternating with testing request number
    m_TextColour = IIf(m_TestingRequestNumber Mod 2 = 0, vbBlack, PRECILAB_BLUE)
    Call m_SampleRowBase.Initialize(Submission, Sample, m_TextColour, Logger)

    Call m_Logger.LogMessage("Initialize", LogLevelEnum.LogInfo, "Successfully initialized InvoiceSampleRowCustomerA")
    Exit Sub
    
ErrorHandler:
    Call m_Logger.LogError("InitializeSampleTableRow", Err.Number, Err.Description)
End Sub

Private Function IInvoiceSampleRow_WriteSampleRow(Table As Word.Table, LineNumber As Long) As Variant
    Table.Cell(LineNumber, 1).Range.Text = m_TestingRequestNumber
    Call m_SampleRowBase.WriteSampleRow(Table, LineNumber, 2)
End Function
Private Sub Class_Initialize()
    Set m_SampleRowBase = New clsInvoiceSampleRowBase
End Sub

Private Sub Class_Terminate()
    Set m_SampleRowBase = Nothing
End Sub
