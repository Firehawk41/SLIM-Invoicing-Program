VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceSampleRowBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceSampleRowBase
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-16
' Description: Base object for sample row
' DependsOn: clsLoggingSystem, clsTestingRequestSubmission, clsTestingRequestSample
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Private m_Logger As clsLoggingSystem
Private m_SampleString As String
Private m_TextColour As Long
Public Sub Initialize(Submission As clsTestingRequestSubmission, Sample As clsTestingRequestSample, TextColour As Long, Logger As clsLoggingSystem)
    
    Set m_Logger = Logger
    
    If Not m_Logger.DebugMode Then
        On Error GoTo ErrorHandler
    End If
    
    Call m_Logger.LogMessage("Initialize", LogLevelEnum.LogDebug, "Initializing InvoiceSampleRow")
    
    m_SampleString = BuildSampleString(Submission, Sample)
    m_TextColour = TextColour

    Call m_Logger.LogMessage("Initialize", LogLevelEnum.LogInfo, "Successfully initialized InvoiceSampleRow: " & m_SampleString)
    Exit Sub
    
ErrorHandler:
    Call m_Logger.LogError("InitializeSampleTableRow", Err.Number, Err.Description)
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Function BuildSampleString(Submission As clsTestingRequestSubmission, Sample As clsTestingRequestSample) As String
    ' Sample String format: mmddyy-FormChemicalName-CustomerName-SampleName
    BuildSampleString = format(Submission.DateReceived, "mmddyy") & "-" & _
                        Sample.FormChemicalName & "-" & _
                        Submission.Customer.CustomerName & "-" & _
                        Sample.SampleName
End Function
Public Function WriteSampleRow(Table As Word.Table, LineNumber As Long, Optional ColumnNumber As Long = 1)
    Table.Cell(LineNumber, ColumnNumber).Range.Text = m_SampleString
    Table.Rows(LineNumber).Range.Font.Color = m_TextColour
End Function

