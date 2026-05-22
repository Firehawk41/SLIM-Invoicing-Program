VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IInvoiceSampleRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: IInvoiceSampleRow
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-11
' Description: Interface for Invoice Sample Rows
' DependsOn: clsTestingRequestSubmission, clsTestingRequestSample, clsLoggingSystem, clsInvoiceSampleRowCustomerA, clsInvoiceSampleRowReg
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Public Sub Initialize(Submission As clsTestingRequestSubmission, Sample As clsTestingRequestSample, TextColour As Long, Logger As clsLoggingSystem)

End Sub

Public Function WriteSampleRow(Table As Word.Table, LineNumber As Long)

End Function

