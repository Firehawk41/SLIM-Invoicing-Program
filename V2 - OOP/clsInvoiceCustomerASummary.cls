VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceCustomerASummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IInvoice
' MacroName: clsInvoiceCustomerA
' Version: 1.0.2
' Author: JT
' Created: ?
' LastModified: 2025-07-22
' Description: Specialized Invoice item for Customer A Summary Invoices
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsInvoiceBase, clsTestingRequestSubmission, clsInvoiceFactory, clsInvoicePriceRetriever, clsSLIMDictionaryManager, clsInvoiceWriter
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
'   - 1.0.2 - 2025-07-22 - Added additional logging statements, changed Submission.SentDate to Submission.ReceivedDate to reflect the renamed property in clsInvoiceBase
Private m_DebugMode As Boolean
Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_BaseInvoice As clsInvoiceBase
Private m_FileName As String
Public Property Get IInvoice_CustomerName() As String
    If Not m_BaseInvoice Is Nothing Then
        IInvoice_CustomerName = m_BaseInvoice.IInvoice_CustomerName
    End If
End Property
Public Property Get IInvoice_WordInvoice() As Word.Document
    Set IInvoice_WordInvoice = m_BaseInvoice.WordInvoice
End Property
Public Property Get IInvoice_FileName() As String
    IInvoice_FileName = m_FileName
End Property
Private Sub IInvoice_Initialize(Submissions As Collection, SequenceNumber As Long, Logger As clsLoggingSystem, AccessDB As clsAccessDatabase, Factory As clsInvoiceFactory, PricingRetriever As clsInvoicePriceRetriever, SLIMDict As clsSLIMDictionaryManager)
    Dim Submission As clsTestingRequestSubmission
    Dim SubmissionMonday As Date
    
    Call Logger.LogMessage("IInvoice_Initialize", LogLevelEnum.LogDebug, "Initializing Invoice Customer A Summary")
    
    ' Set the variables
    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    
    ' Set the submission to the first one in the collection (for this kind of invoice, the collection only contains one submission)
    Set Submission = Submissions(1)
    
    ' Initialize the base invoice and modify its properties as needed
    Call m_BaseInvoice.Initialize(Submissions, SequenceNumber, GetTableIndexes(), Logger, AccessDB, Factory, PricingRetriever, SLIMDict)
    Set m_BaseInvoice.DueDateTableDict = CompileDueDateDict(Submissions)
    
    m_BaseInvoice.InvoiceNumber = GetInvoiceNumber(Submission.Customer.CustomerName)
    
    ' Determine the Monday of the week the first submission was received
    SubmissionMonday = Submission.DateReceived - Weekday(Submission.DateReceived, vbMonday) + 1
    m_BaseInvoice.ReceivedDate = SubmissionMonday
    m_FileName = m_BaseInvoice.SanitizeFileName(SetFileName(SubmissionMonday))
    
    Call Logger.LogMessage("IInvoice_Initialize", LogLevelEnum.LogDebug, "Completed initializing Invoice Customer A Summary")
    
End Sub
Private Function GetTableIndexes() As Scripting.Dictionary
    Dim TableIndexes As New Scripting.Dictionary
    TableIndexes.Add "DueDate", 2
    TableIndexes.Add "LineItem", 3
    TableIndexes.Add "TotalPrice", 4
    TableIndexes.Add "SampleTable", 5
    Set GetTableIndexes = TableIndexes
End Function
Private Function CompileDueDateDict(Submissions As Collection) As Scripting.Dictionary
    Dim DueDateDict As Scripting.Dictionary
    Dim TestingRequestRange As String
    Dim FirstSub As clsTestingRequestSubmission
    Dim LastSub As clsTestingRequestSubmission
    
    Set FirstSub = Submissions(1)
    TestingRequestRange = FirstSub.POInformation
    If Submissions.Count > 1 Then
        Set LastSub = Submissions(Submissions.Count)
        TestingRequestRange = "From " & TestingRequestRange & " to " & LastSub.POInformation & " inclusive"
    End If
    
    ' Set the dictionary from the base invoice, starting at column 2 to leave room for the testing request range
    Set DueDateDict = m_BaseInvoice.CompileDueDateDict(2)
    
    DueDateDict.Add "2,1", TestingRequestRange
    
    ' Return to the calling sub
    Set CompileDueDateDict = DueDateDict
    
End Function
Private Function SetFileName(DateReceived As Date) As String
    SetFileName = "PRECILAB invoice " & m_BaseInvoice.InvoiceNumber & " to Customer A-" & Right(m_BaseInvoice.CustomerName, 2) & " week " & GetWeekNumber(DateReceived) & "  " & format(DateReceived, "mm-dd-yy")
End Function
Private Function GetWeekNumber(InputDate As Date) As Integer
    Dim Jan1 As Date
    Dim WeekOneMonday As Date
    Dim WeekNum As Integer

    ' Determine January 1 for the given year.
    Jan1 = DateSerial(Year(InputDate), 1, 1)
    
    ' If January 1 is Monday-Friday, then the week that contains Jan 1 is week one.
    ' That week starts on the Monday before (or on) Jan 1.
    ' Otherwise, if Jan 1 is Saturday or Sunday, the first Monday-Friday week is the one
    ' beginning on the following Monday.
    If Weekday(Jan1, vbMonday) <= 5 Then
        ' For a weekday, go backward to the Monday of that week.
        WeekOneMonday = DateAdd("d", -(Weekday(Jan1, vbMonday) - 1), Jan1)
    Else
        ' For a weekend day, go forward to the next Monday.
        WeekOneMonday = DateAdd("d", 8 - Weekday(Jan1, vbMonday), Jan1)
    End If

    ' Calculate the week number by counting the number of weeks between the starting Monday
    ' and the InputDate. The +1 ensures the starting week is counted as week 1.
    WeekNum = DateDiff("ww", WeekOneMonday, InputDate, vbMonday) + 1

    GetWeekNumber = WeekNum
End Function

Private Function GetInvoiceNumber(CustomerName As String) As Variant
    Dim InvoiceNumber As Variant

    If CustomerName = "CUSTOMER A TX" Then
        InvoiceNumber = format(DateAdd("d", -Weekday(Now, vbMonday) - 6, Now), "yyyymmdd") & 1
    ElseIf CustomerName = "CUSTOMER A CA" Then
        InvoiceNumber = format(DateAdd("d", -Weekday(Now, vbMonday) - 6, Now), "yyyymmdd") & 2
    Else
        Call m_Logger.LogError("GetInvoiceNumber", Err.Number, Err.Description)
    End If
        
    GetInvoiceNumber = InvoiceNumber
        
End Function
Private Function SetTestingRequestRange(Submissions As Collection) As String
    Dim Submission As clsTestingRequestSubmission
    Dim FirstRequest As Long
    Dim LastRequest As Long
    
    ' Check if there's only one submission in the range
    If Submissions.Count = 1 Then
        ' Set the only submission
        Set Submission = Submissions(1)
        ' Return the only testing request number
        SetTestingRequestRange = Submission.POInformation
    Else
        ' Loop through each submission to retrieve the PO information (ie testing request number)
        For Each Submission In Submissions
            ' Check if the number is less than the current first request number
            If FirstRequest = 0 Or Submission.POInformation < FirstRequest Then
                FirstRequest = Submission.POInformation
            End If
            ' Check if the number is greater than the current last request number
            If Submission.POInformation > LastRequest Then
                LastRequest = Submission.POInformation
            End If
        Next Submission
        
        ' Return the range to the calling sub
        SetTestingRequestRange = FirstRequest & " - " & LastRequest
    End If
    
End Function


Public Sub IInvoice_CreateWordInvoice(WordApp As Word.Application, InvoiceWriter As clsInvoiceWriter)
    Dim WordInvoice As Word.Document
    Const TEMPLATE_PATH As String = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Templates and Logbooks\Customer A Summary Invoice Template.dotx"
    
    Call m_Logger.LogMessage("CreateWordInvoice", LogLevelEnum.LogDebug, "Creating Word Invoice")
    
    Set WordInvoice = m_BaseInvoice.CreateWordInvoice(WordApp, TEMPLATE_PATH, InvoiceWriter)
    
End Sub
Public Sub IInvoice_CreateExcelInvoice(InvoiceWriter As clsInvoiceWriter)
    ' Placeholder - currently only applicable to Customer B summary invoices
End Sub
Public Sub IInvoice_AddInvoiceToDatabase()
    ' Pass through method for adding invoice to database
    Call m_BaseInvoice.AddInvoiceToDatabase
End Sub
Private Sub Class_Initialize()
    Set m_BaseInvoice = New clsInvoiceBase
End Sub
Private Sub Class_Terminate()
    Set m_BaseInvoice = Nothing
End Sub
