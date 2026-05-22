VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsBaseInvoice
' Version: 1.0.2
' Author: JT
' Created: ?
' LastModified: 2025-06-11
' Description: Base Item for specialized Invoice classes
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsInvoiceFactory, modEnums, clsInvoiceLineItemManager, clsInvoiceSampleRowManager, clsTestingRequestSubmission, clsSLIMDictionaryManager, clsInvoiceWriter
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
'   - 1.0.2 - 2025-07-22 - Changed SentDate from Now to Submission.DateReceived and changed the property's name
Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_Factory As clsInvoiceFactory
Private m_TableIndices As Scripting.Dictionary
Private m_ReceivedDate As Date
Private m_PaymentType As PaymentTypeEnum
Private m_PaymentValue As Long
Private m_DueDate As Variant
Private m_LineItemManager As clsInvoiceLineItemManager
Private m_SampleRowManager As clsInvoiceSampleRowManager
Private m_CustomerName As String
Private m_CustomerID As Long
Private m_InvoiceNumber As String
Private m_DueDateTableDict As Scripting.Dictionary
Private m_TotalTableDict As Scripting.Dictionary
Private m_WordInvoice As Word.Document
Private m_ExcelInvoice As Workbook
Public Property Get CustomerName() As String
    CustomerName = m_CustomerName
End Property
Public Property Get PaymentType() As PaymentTypeEnum
    PaymentType = m_PaymentType
End Property
Public Property Get InvoiceNumber() As String
    InvoiceNumber = CStr(m_InvoiceNumber)
End Property
Public Property Get WordInvoice() As Word.Document
    Set WordInvoice = m_WordInvoice
End Property
Public Property Let InvoiceNumber(Value As String)
    m_InvoiceNumber = Value
End Property
Public Property Let PaymentType(Value As PaymentTypeEnum)
    m_PaymentType = Value
End Property
Public Property Set DueDateTableDict(Dict As Scripting.Dictionary)
    Set m_DueDateTableDict = Dict
End Property
Public Property Let ReceivedDate(Value As Date)
    m_ReceivedDate = Value
End Property
Public Sub Initialize(Submissions As Collection, SequenceNumber As Long, TableIndices As Scripting.Dictionary, Logger As clsLoggingSystem, AccessDB As clsAccessDatabase, Factory As clsInvoiceFactory, PricingRetriever As clsInvoicePriceRetriever, SLIMDict As clsSLIMDictionaryManager)
    Dim Submission As clsTestingRequestSubmission
    
    ' Variable initialization
    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    Set m_Factory = Factory
    Set m_TableIndices = TableIndices
    Set Submission = Submissions(1)

    m_CustomerName = Submission.Customer.CustomerName
    m_CustomerID = Submission.Customer.CustomerID
    m_ReceivedDate = Submission.DateReceived
    Call RetrievePaymentTerms(Submission)
    m_DueDate = m_ReceivedDate + m_PaymentValue

    ' Add line items for the submission
    Call m_LineItemManager.Initialize(Submissions, Logger, AccessDB, Factory, PricingRetriever, SLIMDict)

    ' Add sample items for the submission
    Call m_SampleRowManager.Initialize(Submissions, Logger, Factory)
    
    ' Compile the total table dictionary
    Set m_TotalTableDict = CompileTotalDict()
    
End Sub
Private Sub RetrievePaymentTerms(Submission As clsTestingRequestSubmission)
    Dim SQLQuery As String
    Dim PaymentTerms As Variant
    
    SQLQuery = "SELECT quotes.payment_type, quotes.payment_value " & _
                "FROM quotes INNER JOIN (customers INNER JOIN quote_customer ON customers.ID = quote_customer.customer_id) ON quotes.ID = quote_customer.quote_id " & _
                "WHERE customers.customer_name = '" & m_CustomerName & "'"
                
    PaymentTerms = m_AccessDB.ExecuteQuery(SQLQuery)
    
    If UBound(PaymentTerms) = -1 Then
        ' Set default values if there are not quotes in the database
        m_PaymentType = PaymentTypeEnum.PONumber
        m_PaymentValue = 30
    Else
        If CStr(PaymentTerms(0, 0)) = "Credit Card" Then
            m_PaymentType = PaymentTypeEnum.CreditCard
        Else
            m_PaymentType = PaymentTypeEnum.PONumber
        End If
        m_PaymentValue = CLng(PaymentTerms(1, 0))
    End If
End Sub
Private Function CompileTotalDict() As Scripting.Dictionary
    Dim TotalDict As New Scripting.Dictionary
    
    ' Fills out the Sub total and grand total
    TotalDict.Add "1,2", format(m_LineItemManager.TotalCost, "#,##0.00") & " USD"
    TotalDict.Add "3,2", format(m_LineItemManager.TotalCost, "#,##0.00") & " USD"

    ' Return the dictionary to the calling sub
    Set CompileTotalDict = TotalDict
    
End Function
Public Function CompileDueDateDict(Optional StartColumn As Integer = 1) As Scripting.Dictionary
    Dim DueDateDict As New Scripting.Dictionary
    
    Call m_Logger.LogMessage("SetDueDateInformation", LogLevelEnum.LogDebug, "Filling in due date information")

    ' Set the due date information that all invoices share
    DueDateDict.Add "2," & StartColumn, format(m_ReceivedDate, "mm-dd-yyyy")
    
    If m_PaymentType = PaymentTypeEnum.CreditCard Then
        DueDateDict.Add "2," & StartColumn + 1, "Credit Card"
        DueDateDict.Add "2," & StartColumn + 2, "N.A."
    Else
        DueDateDict.Add "2," & StartColumn + 1, "Net " & m_PaymentValue & " days"
        DueDateDict.Add "2," & StartColumn + 2, IIf(IsDate(m_DueDate), format(m_DueDate, "mm-dd-yyyy"), "N/A")
    End If

    ' Return the dictionary to the calling sub
    Set CompileDueDateDict = DueDateDict

End Function
Public Function CreateWordInvoice(WordApp As Word.Application, TemplatePath As String, InvoiceWriter As clsInvoiceWriter) As Word.Document
    Dim WordInvoice As Word.Document
    Const PRECILAB_BLUE As Long = 16711680
    
    Call m_Logger.LogMessage("CreateWordInvoice", LogLevelEnum.LogDebug, "Creating Word Invoice: " & m_CustomerName)

    ' Create a new word document from the passed template file path and store the reference as a property for later retrieval
    Set m_WordInvoice = WordApp.Documents.Add(Template:=TemplatePath, NewTemplate:=False, DocumentType:=0)

    ' Set common properties
    Call InvoiceWriter.WriteToBookmark(m_WordInvoice, "InvoiceNumber", CStr(m_InvoiceNumber), 10, PRECILAB_BLUE, True)
    Call InvoiceWriter.WriteToBookmark(m_WordInvoice, "InvoiceDate", format(m_ReceivedDate, "mm-dd-yyyy"), 10, PRECILAB_BLUE, True)
        
    ' Fill in the due date information
    Call InvoiceWriter.WriteToTable(m_WordInvoice.Tables(m_TableIndices("DueDate")), m_DueDateTableDict)
    
    ' Fill in the line item table
    Call m_LineItemManager.FillInLineItemTable(m_WordInvoice.Tables(m_TableIndices("LineItem")))
    
    ' Fill in the sample ID table
    Call m_SampleRowManager.FillInSampleTable(m_WordInvoice.Tables(m_TableIndices("SampleTable")))

    ' Fill in sub total and grand total
    Call InvoiceWriter.WriteToTable(m_WordInvoice.Tables(m_TableIndices("TotalPrice")), m_TotalTableDict)
    
    ' Return the invoice to the calling sub
    Set CreateWordInvoice = m_WordInvoice

    Call m_Logger.LogMessage("CreateWordInvoice", LogLevelEnum.LogInfo, "Completed Word Invoice: " & m_CustomerName)

End Function

Public Sub AddInvoiceToDatabase()
    Dim SQLStatement As String
    
    SQLStatement = "INSERT INTO invoices (invoice_date, customer_id, invoice_number, invoice_amount) " & _
               "VALUES (Date(), " & m_CustomerID & ", '" & m_InvoiceNumber & "', " & m_LineItemManager.TotalCost & ");"
               
    Call m_AccessDB.ExecuteNonQuery(SQLStatement)
    
End Sub
Public Function SanitizeFileName(ByVal FileName As String, Optional Replacement As String = "_") As String
    Dim IllegalChars As String
    Dim i As Long
    
    ' List of illegal characters for cross-platform compatibility
    IllegalChars = "\/:*?""<>|"
    
    ' Replace each illegal character with the Replacement character
    For i = 1 To Len(IllegalChars)
        FileName = Replace(FileName, Mid(IllegalChars, i, 1), Replacement)
    Next i
    
    ' Trim trailing spaces and dots (Windows restriction)
    FileName = Trim(FileName)
    Do While Right(FileName, 1) = "." Or Right(FileName, 1) = " "
        FileName = Left(FileName, Len(FileName) - 1)
    Loop
    
    ' Return sanitized filename
    SanitizeFileName = FileName
End Function
Public Function CreateExcelInvoice(TemplatePath As String, InvoiceWriter As clsInvoiceWriter) As Workbook
    
    Call m_Logger.LogMessage("CreateExcelInvoice", LogLevelEnum.LogDebug, "Creating Excel Invoice: " & m_CustomerName)
    
    Set m_ExcelInvoice = Workbooks.Add(Template:=TemplatePath)
    
    
    
    ' Return the workbook to the calling sub
    Set CreateExcelInvoice = m_ExcelInvoice

End Function

Private Sub Class_Initialize()
    Set m_LineItemManager = New clsInvoiceLineItemManager
    Set m_SampleRowManager = New clsInvoiceSampleRowManager
End Sub
Private Sub Class_Terminate()
    Set m_LineItemManager = Nothing
    Set m_SampleRowManager = Nothing
End Sub
