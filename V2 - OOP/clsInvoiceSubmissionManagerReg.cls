VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceSubmissionManagerReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceSubmissionManagerReg
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-05
' Description: Submission Manager for regular summary invoices
' DependsOn: clsLoggingSystem, IInvoiceSubmissionManager, clsInvoiceSubmissionManagerBase, clsTestingRequestSubmission
' ChangeLog:
'   - 1.0.1 - 2025-06-05 - Added metadata
Implements IInvoiceSubmissionManager
Private m_Logger As clsLoggingSystem
Private m_BaseClass As clsInvoiceSubmissionManagerBase
Private m_Submissions As Collection
Public Property Get IInvoiceSubmissionManager_Submissions() As Collection
    Set IInvoiceSubmissionManager_Submissions = m_Submissions
End Property
Public Sub IInvoiceSubmissionManager_Initialize(StartDate As Date, SelectedFilePath As String, Logger As clsLoggingSystem)

    Set m_Logger = Logger
    Set m_BaseClass = New clsInvoiceSubmissionManagerBase
    Call m_BaseClass.Initialize(StartDate, SelectedFilePath, Logger)

    
    ' Set the collection
    Set m_Submissions = m_BaseClass.SetIndividualCollection(SelectedFilePath)
    
End Sub
