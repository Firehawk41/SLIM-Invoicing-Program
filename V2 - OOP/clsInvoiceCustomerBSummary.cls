VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceCustomerBSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IInvoice
' MacroName: clsInvoiceCustomerB
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-11
' Description: Specialized Invoice item for Customer B Summary invoices
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsInvoiceBase, clsInvoiceCustomerBSiteInfo, clsTestingRequestSubmission, clsInvoiceFactory, clsInvoicePriceRetriever, clsSLIMDictionaryManager, modEnums, clsInvoiceWriter
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Private m_DebugMode As Boolean
Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_BaseInvoice As clsInvoiceBase
Private m_FileName As String
Private m_POInformation As String
Private m_InvoiceMonth As Date
Private m_SiteInfo As clsInvoiceCustomerBSiteInfo
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
Public Sub IInvoice_Initialize(Submissions As Collection, SequenceNumber As Long, Logger As clsLoggingSystem, AccessDB As clsAccessDatabase, Factory As clsInvoiceFactory, PricingRetriever As clsInvoicePriceRetriever, SLIMDict As clsSLIMDictionaryManager)
    Dim Submission As clsTestingRequestSubmission
    
    Call Logger.LogMessage("IInvoice_Initialize", LogLevelEnum.LogDebug, "Initializing Invoice Customer B Summary")
    
    ' Set the variables
    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    
    ' Set the submission to the first one in the collection
    Set Submission = Submissions(1)
    
    ' Initialize the base invoice and modify its properties as needed
    Call m_BaseInvoice.Initialize(Submissions, SequenceNumber, GetTableIndexes(), Logger, AccessDB, Factory, PricingRetriever, SLIMDict)
    Set m_BaseInvoice.DueDateTableDict = CompileDueDateDict(Submission)
    m_POInformation = SetPOInformation(Submission)
    m_InvoiceMonth = Submission.DateReceived
    
    Call m_SiteInfo.Initialize(Submission.Customer.CustomerName, Submission.Samples(1).Chemical.ChemicalName)
    m_BaseInvoice.InvoiceNumber = SetFileName(Submission.Customer.CustomerName, SequenceNumber)
    m_FileName = m_BaseInvoice.SanitizeFileName(m_BaseInvoice.InvoiceNumber)
    m_BaseInvoice.InvoiceNumber = m_FileName
    
End Sub
Private Function GetTableIndexes() As Scripting.Dictionary
    Dim TableIndexes As New Scripting.Dictionary
    TableIndexes.Add "DueDate", 1
    TableIndexes.Add "LineItem", 2
    TableIndexes.Add "TotalPrice", 3
    TableIndexes.Add "SampleTable", 4
    Set GetTableIndexes = TableIndexes
End Function
Private Function CompileDueDateDict(Submission As clsTestingRequestSubmission) As Scripting.Dictionary
    ' This is a pass-through function that exists to modify base invoice's due date dictionary in future if I need to refactor JT 2/6/25
    Dim DueDateDict As Scripting.Dictionary
    
    ' Set the dictionary from the base invoice
    Set DueDateDict = m_BaseInvoice.CompileDueDateDict()
    
    ' Return to the calling sub
    Set CompileDueDateDict = DueDateDict
    
End Function
Private Function SetFileName(Customer As String, SequenceNumber As Long) As String
    Dim DateCode As String
    Dim IncrementNumber As Integer
    Dim FileName As String
    Const DOW_CODE As String = "PLB"
    
    DateCode = format(Now, "yymmdd")
    
    FileName = m_SiteInfo.CompanyCode & DOW_CODE & DateCode
    
    IncrementNumber = GetIncrementNumber("Customer B", FileName) + SequenceNumber

    SetFileName = FileName & IncrementNumber
    
End Function
Private Function GetIncrementNumber(Customer As String, PreliminaryInvoiceNumber As String) As Integer
    
    Dim SQLQuery As String
    Dim DataArr As Variant
    Dim IncrementNumber As Variant
    
    ' Set the SQL Query for the database
    SQLQuery = "SELECT COUNT(*) FROM invoices " & _
           "INNER JOIN customers ON invoices.customer_id = customers.ID " & _
           "WHERE customers.Customer_Name = '" & Customer & "' " & _
           "AND invoices.invoice_date = Date();"
           
    ' Use the Database Access object to retrieve the result
    IncrementNumber = m_AccessDB.ExecuteScalarQuery(SQLQuery)
    
    ' Return the number to the calling sub
    GetIncrementNumber = IncrementNumber
    
End Function
Private Function SetPOInformation(Submission As clsTestingRequestSubmission) As String
    Dim POInformation As String
    
    ' Retrieve the PO number from the submission
    POInformation = Submission.POInformation
    
    ' Format the PO information
    If POInformation <> "" And Left(POInformation, 2) <> "PO" Then
        POInformation = "PO: " & POInformation
    End If
    
    ' Return the string to the calling sub
    SetPOInformation = POInformation
End Function
Public Sub IInvoice_CreateWordInvoice(WordApp As Word.Application, InvoiceWriter As clsInvoiceWriter)
    Dim WordInvoice As Word.Document
    Const TEMPLATE_PATH As String = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Templates and Logbooks\Customer B Summary Invoice Template.dotx"
    Const PRECILAB_BLUE As Long = 16711680
    
    Call m_Logger.LogMessage("CreateWordInvoice", LogLevelEnum.LogDebug, "Creating Word Invoice")
    
    ' Create the word invoice with the base invoice object
    Set WordInvoice = m_BaseInvoice.CreateWordInvoice(WordApp, TEMPLATE_PATH, InvoiceWriter)
    
    ' Set the properties that are particular to the Customer B invoice
    Call InvoiceWriter.WriteToBookmark(WordInvoice, "Month", format(m_InvoiceMonth, "mmmm yyyy"), 14, PRECILAB_BLUE, True, "Calibri")
    Call InvoiceWriter.WriteToBookmark(WordInvoice, "Site", m_SiteInfo.Site, , PRECILAB_BLUE, True)
    Call InvoiceWriter.WriteToBookmark(WordInvoice, "SiteContact", m_SiteInfo.SiteContact, , PRECILAB_BLUE, True)
    Call InvoiceWriter.WriteToBookmark(WordInvoice, "CompanyCode", m_SiteInfo.CompanyCode, , PRECILAB_BLUE, True)
    Call InvoiceWriter.WriteToBookmark(WordInvoice, "CostCenter", m_SiteInfo.CostCenter, , PRECILAB_BLUE, True)
    Call InvoiceWriter.WriteToBookmark(WordInvoice, "Plant", m_SiteInfo.Plant, , PRECILAB_BLUE, True)
    
End Sub
Public Sub IInvoice_AddInvoiceToDatabase()
    ' Pass through method for adding invoice to database
    Call m_BaseInvoice.AddInvoiceToDatabase
End Sub
Public Sub IInvoice_CreateExcelInvoice(InvoiceWriter As clsInvoiceWriter)
    Dim ExcelInvoice As Workbook
    Const TEMPLATE_PATH As String = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Templates and Logbooks\Customer B Script Template.xltx"
    
    Set ExcelInvoice = m_BaseInvoice.CreateExcelInvoice(TEMPLATE_PATH, InvoiceWriter)
    
End Sub
Private Sub Class_Initialize()
    Set m_BaseInvoice = New clsInvoiceBase
    Set m_SiteInfo = New clsInvoiceCustomerBSiteInfo
End Sub
Private Sub Class_Terminate()
    Set m_BaseInvoice = Nothing
    Set m_SiteInfo = Nothing
End Sub
