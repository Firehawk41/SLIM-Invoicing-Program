VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceSubmissionManagerCustomerB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceSubmissionManagerCustomerB
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-05
' Description: Submission Manager for Customer B summary invoices
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
    Dim EndDate As Date
    Const DOW_FILE_PATH As String = "\\PRECILAB-SERVER\LabPlusServer\Testing_Requests\Customer B Chemical\"
    
    Set m_Logger = Logger
    
    Call m_BaseClass.Initialize(StartDate, SelectedFilePath, Logger)
    
    
    ' Set the end date as the last day of the month
    EndDate = DateSerial(Year(StartDate), Month(StartDate) + 1, 0)
    
    ' Set the collection
    Set m_Submissions = m_BaseClass.SetSubmissionCollection(StartDate, EndDate, DOW_FILE_PATH)
    
End Sub
Private Sub Class_Initialize()
    Set m_BaseClass = New clsInvoiceSubmissionManagerBase
End Sub
Private Sub Class_Terminate()
    Set m_BaseClass = New clsInvoiceSubmissionManagerBase
End Sub
