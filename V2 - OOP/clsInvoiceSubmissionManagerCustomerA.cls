VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceSubmissionManagerCustomerA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceSubmissionManagerCustomerA
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-05
' Description: Submission Manager for Customer A summary invoices
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
    
    Call m_BaseClass.Initialize(StartDate, SelectedFilePath, Logger)
    
    ' Set the collection
    Set m_Submissions = SetCustomerASubmissions(StartDate, StartDate + 5)
    
End Sub
Private Function SetCustomerASubmissions(StartDate As Date, EndDate As Date) As Collection
    Const FEUP_TX_FILE_PATH As String = "\\PRECILAB-SERVER\LabPlusServer\Testing_Requests\Customer A TX\"
    Const FEUP_CA_FILE_PATH As String = "\\PRECILAB-SERVER\LabPlusServer\Testing_Requests\Customer A CA\"
    Dim TXSubmissions As Collection
    Dim CASubmissions As Collection
    
    ' Set the Texas and California collections and then combine them in order to create both Customer A summary invoices in one fell swoop
    Set TXSubmissions = m_BaseClass.SetSubmissionCollection(StartDate, EndDate, FEUP_TX_FILE_PATH)
    Set CASubmissions = m_BaseClass.SetSubmissionCollection(StartDate, EndDate, FEUP_CA_FILE_PATH)
    
    ' Return the combined collection to the calling sub
    Set SetCustomerASubmissions = CombineSubmissions(TXSubmissions, CASubmissions)

End Function
Private Function CombineSubmissions(Collection1 As Collection, Collection2 As Collection) As Collection
    Dim Submission As clsTestingRequestSubmission
    
    Call m_Logger.LogMessage("CombineSubmissions", LogLevelEnum.LogDebug, "Combining TX and CA submissions.")
    
    For Each Submission In Collection2
        Collection1.Add Submission
    Next Submission
    
    Set CombineSubmissions = Collection1
End Function
Private Sub Class_Initialize()
    Set m_BaseClass = New clsInvoiceSubmissionManagerBase
End Sub
Private Sub Class_Terminate()
    Set m_BaseClass = New clsInvoiceSubmissionManagerBase
End Sub
