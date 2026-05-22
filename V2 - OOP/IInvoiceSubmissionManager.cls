VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IInvoiceSubmissionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: IInvoiceSubmissionManager
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-05
' Description: Interface for Submission Managers
' DependsOn: clsLoggingSystem, clsInvoiceSubmissionManagerCustomerB, clsInvoiceSubmissionManagerCustomerA, clsInvoiceSubmissionManagerReg
' ChangeLog:
'   - 1.0.1 - 2025-06-05 - Added metadata
Public Property Get Submissions() As Collection
    
End Property
Public Sub Initialize(StartDate As Date, SelectedFilePath As String, Logger As clsLoggingSystem)

End Sub

