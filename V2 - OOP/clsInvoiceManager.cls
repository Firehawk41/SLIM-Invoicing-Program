VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceManager
' Version: 1.0.2
' Author: JT
' Created: ?
' LastModified: 2025-07-08
' Description: Object in charge of managing invoices
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsInvoiceFactory, clsInvoicePriceRetriever, clsSLIMDictionaryManager, clsInvoiceWriter, IInvoice, IInvoiceSubmissionManager
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
'   - 1.0.2 - 2025-07-08 - Add Customer B Plaquemine to the list of 0001 sites and fixed bug that wouldn't use incremented file name on save
Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_Factory As clsInvoiceFactory
Private m_Invoices As Collection
Private m_PriceRetriever As clsInvoicePriceRetriever
Private m_SLIMDict As clsSLIMDictionaryManager
Private m_InvoiceWriter As clsInvoiceWriter
Private m_WrdApp As Word.Application
Public Property Get Invoices() As Collection
    Set Invoices = m_Invoices
End Property

Public Sub Initialize(SubmissionManager As IInvoiceSubmissionManager, Logger As clsLoggingSystem, AccessDB As clsAccessDatabase, Factory As clsInvoiceFactory)
    
    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    Set m_Factory = Factory
    Call m_SLIMDict.Initialize(Logger)
    Call m_PriceRetriever.Initialize(Logger)
    Set m_Invoices = GenerateInvoices(SubmissionManager)
    Call m_InvoiceWriter.Initialize(Logger)

End Sub

Private Function GenerateInvoices(SubmissionManager As IInvoiceSubmissionManager) As Collection
    ' Create a new collection to hold the summary invoice objects

    Dim CustomerGroups As Scripting.Dictionary

    If Not m_Logger.DebugMode Then
        On Error GoTo ErrorHandler
    End If
    
    ' Use a dictionary to group submissions by customer name
    Set CustomerGroups = SortSubmissionsToInvoiceDictionary(SubmissionManager)
    
    ' Process each customer group and generate invoice objects
    Set GenerateInvoices = ProcessCustomerGroups(CustomerGroups)
    Exit Function
ErrorHandler:
    Call m_Logger.LogError("GenerateInvoiceCollection", Err.Number, Err.Description)
    Set GenerateInvoices = Nothing
End Function
Private Function SortSubmissionsToInvoiceDictionary(SubmissionManager As IInvoiceSubmissionManager) As Scripting.Dictionary
    Dim CustomerGroups As Scripting.Dictionary
    Dim Submission As clsTestingRequestSubmission
    Dim Key As Variant

    Set CustomerGroups = New Scripting.Dictionary
    
    If Not m_Logger.DebugMode Then
        On Error GoTo ErrorHandler
    End If

    ' Iterate through each submission in the collection
    For Each Submission In SubmissionManager.Submissions
        ' Set key by CustomerType
        Key = m_Factory.GetKeyByCustomerType(Submission)
        
        If Not CustomerGroups.Exists(Key) Then
            Set CustomerGroups(Key) = New Collection
        End If
        CustomerGroups(Key).Add Submission
    Next Submission
    
    Set SortSubmissionsToInvoiceDictionary = CustomerGroups
    Exit Function
ErrorHandler:
    Call m_Logger.LogError("SortSubmissionsToInvoiceDictionary", Err.Number, Err.Description)
    Set SortSubmissionsToInvoiceDictionary = Nothing
End Function
Private Function ProcessCustomerGroups(CustomerGroups As Scripting.Dictionary) As Collection
    Dim Key As Variant
    Dim KeyCounter As Long
    Dim CustomerSubmissions As Collection
    Dim Invoice As IInvoice
    Dim InvoiceCollection As New Collection
    Dim SequenceNumber As Long
    Dim Counter0001 As Long
    Dim Counter2000 As Long
    
    
    For Each Key In CustomerGroups.Keys
        ' Increment the counters
        KeyCounter = KeyCounter + 1
        SequenceNumber = IncrementCounter(Counter0001, Counter2000, Key)
        
        ' Set the submission collection
        Set CustomerSubmissions = CustomerGroups(Key)
        
        ' Log the processing message
        Call m_Logger.LogMessage("ProcessCustomerGroups", LogLevelEnum.LogInfo, "Processing invoices for " & Key & " (Group " & KeyCounter & ")")
        
        ' Create new invoice object and add it to the collection
        Set Invoice = m_Factory.CreateInvoice()
        Call Invoice.Initialize(CustomerSubmissions, SequenceNumber, m_Logger, m_AccessDB, m_Factory, m_PriceRetriever, m_SLIMDict)
        InvoiceCollection.Add Invoice
    Next Key
    
    ' Return the collection to the calling sub
    Set ProcessCustomerGroups = InvoiceCollection
End Function
Private Function IncrementCounter(Counter0001 As Long, Counter2000 As Long, Key As Variant)
    Dim CustomerName As String

    CustomerName = Split(Key, "_")(0)
    
    ' Determine prefix
    Select Case CustomerName
        Case "Customer B Carteret", "Customer B Midland", "Customer B Monument", "Customer B Third Coast Terminals", "Customer B Freeport", "Customer B Plaquemine"
            Counter0001 = Counter0001 + 1
            IncrementCounter = Counter0001
            Call m_Logger.LogMessage("IncrementCounter", LogLevelEnum.LogDebug, "Incremented Counter0001 to " & Counter0001)
        Case "Customer B Texas City", "Customer B Seadrift", "Customer B Hahnville"
            Counter2000 = Counter2000 + 1
            IncrementCounter = Counter2000
            Call m_Logger.LogMessage("IncrementCounter", LogLevelEnum.LogDebug, "Incremented Counter2000 to " & Counter2000)
        Case Else
            IncrementCounter = 1
            Call m_Logger.LogMessage("IncrementCounter", LogLevelEnum.LogDebug, "This is not a recognized Customer B customer.  The sequence number has been set to 1.")
    End Select

End Function

Public Sub GenerateWordInvoiceCollection()
    Dim Invoice As IInvoice

    Dim InvoiceDoc As Word.Document
    Dim InvoiceWriter As New clsInvoiceWriter

    Call m_Logger.LogMessage("GenerateWordInvoiceCollection", LogLevelEnum.LogDebug, "Initiating Invoice Document Collection")

    Call InvoiceWriter.Initialize(m_Logger)
    Set m_WrdApp = CreateObject("Word.Application")

    If m_Logger.DebugMode Then
       m_WrdApp.Visible = True
    Else
        On Error GoTo ErrorHandler
    End If

    For Each Invoice In m_Invoices

        ' Create an invoice as a Word Document
        Call Invoice.CreateWordInvoice(m_WrdApp, m_InvoiceWriter)
        
        ' Create an Excel version of the invoice (Customer B only for now)
        Call Invoice.CreateExcelInvoice(m_InvoiceWriter)

    Next Invoice

    ' Terminate the invoice writer
    Set InvoiceWriter = Nothing
    
    Exit Sub
ErrorHandler:
    Call m_Logger.LogError("GenerateWordInvoiceCollection", Err.Number, Err.Description)
End Sub
Public Sub SaveInvoices(DebugMode As Boolean)
    Dim Invoice As IInvoice
    
    For Each Invoice In m_Invoices
        Call SaveInvoice(Invoice.WordInvoice, Invoice.FileName, DebugMode)
    Next Invoice
        
End Sub
Private Sub SaveInvoice(WordInvoice As Word.Document, FileName As String, DebugMode As Boolean)
    Const INVOICE_FILE_PATH As String = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Saved Invoices\Auto-Generated Invoices\"
    Dim FullPath As String
    Dim Increment As Long
    Dim FinalFileName As String

    FinalFileName = FileName
    FullPath = INVOICE_FILE_PATH & FinalFileName & ".docx"

    ' Loop until we find a filename that doesn't exist
    Do While Dir(FullPath) <> ""
        Increment = Increment + 1
        FinalFileName = FileName & "_" & Increment
        FullPath = INVOICE_FILE_PATH & FinalFileName & ".docx"
    Loop
    
    If Not DebugMode Then
        ' Save the invoice
        WordInvoice.SaveAs2 FileName:=INVOICE_FILE_PATH & FinalFileName & ".docx"
        Call m_Logger.LogMessage("SaveInvoices", LogLevelEnum.LogInfo, "Invoice saved as " & FileName)
    Else
        Call m_Logger.LogMessage("SaveInvoices", LogLevelEnum.LogInfo, "In debug mode. Invoice would have been saved as " & FileName)
    End If

End Sub
Public Sub CloseInvoices(DebugMode As Boolean)
    Dim Invoice As IInvoice
    
    For Each Invoice In m_Invoices
        If Not DebugMode Then
            Invoice.WordInvoice.Close SaveChanges:=True
            Call m_Logger.LogMessage("CloseInvoices", LogLevelEnum.LogInfo, "Closed invoice: " & Invoice.FileName)
        Else
            Call m_Logger.LogMessage("CloseInvoices", LogLevelEnum.LogInfo, "In debug mode. Invoice would have been closed: " & Invoice.FileName)
        End If
    Next Invoice
    
    If Not DebugMode Then m_WrdApp.Quit
End Sub
Public Sub AddInvoicesToDatabase(DebugMode As Boolean)
    Dim Invoice As IInvoice
    
    For Each Invoice In m_Invoices
        If Not DebugMode Then
            Call Invoice.AddInvoiceToDatabase
            Call m_Logger.LogMessage("CloseInvoices", LogLevelEnum.LogInfo, "Added invoice to database: " & Invoice.FileName)
        Else
            Call m_Logger.LogMessage("CloseInvoices", LogLevelEnum.LogInfo, "In debug mode. Would have added invoice to database: " & Invoice.FileName)
        End If
    Next Invoice
End Sub
Private Sub Class_Initialize()
    Set m_Invoices = New Collection
    Set m_PriceRetriever = New clsInvoicePriceRetriever
    Set m_SLIMDict = New clsSLIMDictionaryManager
    Set m_InvoiceWriter = New clsInvoiceWriter
End Sub

Private Sub Class_Terminate()
    Dim Invoice As IInvoice

    For Each Invoice In m_Invoices
        Set Invoice = Nothing
    Next Invoice
    Set m_Invoices = Nothing
    Set m_PriceRetriever = Nothing
    Set m_SLIMDict = Nothing
    Set m_InvoiceWriter = Nothing
End Sub
