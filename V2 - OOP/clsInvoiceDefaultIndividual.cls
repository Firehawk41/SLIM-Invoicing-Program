VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceDefaultIndividual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IInvoice
' MacroName: clsInvoiceRegular
' Version: 1.0.2
' Author: JT
' Created: ?
' LastModified: 2025-07-22
' Description: Specialized Invoice item for regular, individual invoices
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsInvoiceBase, clsInvoiceContact, clsTestingRequestSubmission, clsInvoiceFactory, clsInvoicePriceRetriever, clsSLIMDictionaryManager, modEnums, clsInvoiceWriter
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
'   - 1.0.2 - 2025-07-22 - Modified how file name is constructed as per Yves's instructions.
'   - 1.1.0 - 2025-07-23 - Modified the ContactDict and Initialization to pass the email of the request contact and invoice main to the writer instead of the contact's full details
Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_BaseInvoice As clsInvoiceBase
Private m_POInformation As String
Private m_CCInfo As String
Private m_CustomerAddress As String
Private m_CustomerContact As clsInvoiceContact
Private m_CustomerCC1 As clsInvoiceContact
Private m_CustomerCC2 As clsInvoiceContact
Private m_FileName As String
Private m_LogoFileLocation As String
Private m_ContactDict As Scripting.Dictionary
Private m_CCInfoDict As Scripting.Dictionary
Private m_InvoiceEmail As String
Private m_RequestEmail As String
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
    
    Call Logger.LogMessage("IInvoice_Initialize", LogLevelEnum.LogDebug, "Initializing Invoice Default Individual")
    
    ' Set the variables
    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    
    ' Set the submission to the first one in the collection (for this kind of invoice, the collection only contains one submission)
    Set Submission = Submissions(1)
    
    ' Initialize the base invoice and modify its properties as needed
    Call m_BaseInvoice.Initialize(Submissions, SequenceNumber, GetTableIndexes(), Logger, AccessDB, Factory, PricingRetriever, SLIMDict)
    Set m_BaseInvoice.DueDateTableDict = CompileDueDateDict(Submission)
    m_BaseInvoice.PaymentType = GetPaymentType(Submission)
    m_BaseInvoice.InvoiceNumber = GetInvoiceNumber(Submission)
    
    m_POInformation = SetPOInformation(Submission)
    m_CCInfo = Submission.CreditCardInformation
    m_FileName = m_BaseInvoice.SanitizeFileName(SetFileName(Submission))
    m_CustomerAddress = Submission.Customer.FormattedCustomerAddress
    Call m_CustomerContact.InitializeFromSubmission(Submission, Logger) ' may be irrelevant JT 7/23/25
    m_RequestEmail = Submission.ResultsEmailMainFirst
    m_InvoiceEmail = Submission.InvoiceEmailMainFirst
    Set m_ContactDict = CompileCustomerContactDict
    Set m_CCInfoDict = CompileCCDict

'        Call m_CustomerCC1.InitializeFromDatabase(CustomerName, "CC1_Name", "CC1_Title", "CC1_Email", "CC1_Phone", AccessDB, Logger)
'        Call m_CustomerCC2.InitializeFromDatabase(CustomerName, "CC2_Name", "CC2_Title", "CC2_Email", "CC2_Phone", AccessDB, Logger)
    m_LogoFileLocation = SetLogoFileLocation(Submission.Customer.CustomerName)
    
    Call m_Logger.LogMessage("IINvoice_Initialize", LogLevelEnum.LogInfo, "Successfully initialized Invoice Default Individual")
    
End Sub

Private Function GetPaymentType(Submission As clsTestingRequestSubmission) As PaymentTypeEnum
    Dim PaymentType As PaymentTypeEnum
    
    ' If the default from the database is credit card, enforce it
    If m_BaseInvoice.PaymentType = CreditCard Then
        PaymentType = PaymentTypeEnum.CreditCard
    ' If PO is provided and POs are allowed (not defaulted to Credit Card), prioritize PO
    ElseIf Submission.POInformation <> "" Then
        PaymentType = PaymentTypeEnum.PONumber
    ' If no PO but a credit card is supplied, use credit card
    ElseIf Submission.CreditCardInformation <> "" And Not Submission.CreditCardInformation Like "*N/A*" Then
        PaymentType = PaymentTypeEnum.CreditCard
    ' Otherwise, default to the base invoice setting
    Else
        PaymentType = m_BaseInvoice.PaymentType
    End If
    
    ' Return the enum to the calling sub
    GetPaymentType = PaymentType
    
End Function
Private Function CompileDueDateDict(Submission As clsTestingRequestSubmission) As Scripting.Dictionary
    ' This is a pass-through function that exists to modify base invoice's due date dictionary in future if I need to refactor JT 2/6/25
    Dim DueDateDict As Scripting.Dictionary
    
    ' Set the dictionary from the base invoice
    Set DueDateDict = m_BaseInvoice.CompileDueDateDict()
    
    ' Return to the calling sub
    Set CompileDueDateDict = DueDateDict
    
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
Private Function CompileCCDict() As Scripting.Dictionary
    Dim CCDict As New Scripting.Dictionary
        
    ' Set the "already paid" string to cell 1,1
    If m_BaseInvoice.PaymentType = PaymentTypeEnum.CreditCard Then
        CCDict.Add "1,1", "DO NOT PAY" & vbNewLine & "Already paid by CREDIT CARD #: " & m_CCInfo & " on " & format(Now, "mm-dd-yyyy")
    End If

    ' Return the dictionary to the calling sub
    Set CompileCCDict = CCDict

End Function

Private Function GetTableIndexes() As Scripting.Dictionary
    Dim TableIndexes As New Scripting.Dictionary
    TableIndexes.Add "DueDate", 2
    TableIndexes.Add "LineItem", 3
    TableIndexes.Add "TotalPrice", 4
    TableIndexes.Add "SampleTable", 5
    Set GetTableIndexes = TableIndexes
End Function

Public Function SetFileName(Submission As clsTestingRequestSubmission) As String
    Dim FileName As String
    Dim POInformation As String
    Dim BaseCase As String
    
    ' Set the base case
    BaseCase = m_BaseInvoice.InvoiceNumber & " to " & m_BaseInvoice.CustomerName & " [" & format(Submission.DateReceived, "mm-dd-yyyy") & "]"
    
    
    'Sets the file name
    If m_BaseInvoice.PaymentType = PaymentTypeEnum.PONumber Then
        ' Set the PO number with correct formatting
        If Not Left(Submission.POInformation, 2) = "PO" Then
            POInformation = "PO " & Submission.POInformation
        Else
            POInformation = Submission.POInformation
        End If
    
        FileName = "PRECILAB invoice " & BaseCase & " " & Submission.POInformation
    
    ElseIf m_BaseInvoice.PaymentType = PaymentTypeEnum.CreditCard Then
        FileName = "PRECILAB invoice receipt " & BaseCase
    End If
    
    ' Return the string to the calling sub
    SetFileName = FileName
    
End Function
Private Function GetInvoiceNumber(Submission As clsTestingRequestSubmission) As Variant

    Dim SequenceNumber As Variant
    Dim Array_Index As Integer
    Dim SQLQuery As String
    Dim InvoiceNumber As Variant

    ' Set the SQL Query for the database
    SQLQuery = "SELECT COUNT(*) FROM invoices " & _
           "INNER JOIN customers ON invoices.customer_id = customers.ID " & _
           "WHERE customers.Customer_Name = '" & Submission.Customer.CustomerName & "' " & _
           "AND invoices.invoice_date = Date();"

    ' Use the Database Access object to retrieve the result
    SequenceNumber = m_AccessDB.ExecuteScalarQuery(SQLQuery)

    ' Combine the date with the sequence number
    InvoiceNumber = format(Submission.DateReceived, "yyyymmdd") & SequenceNumber + 1

    ' Return the number to the calling sub
    GetInvoiceNumber = InvoiceNumber

End Function
Private Function SetLogoFileLocation(CustomerName As String) As String
    Dim LogoPath As String
    Dim FullLogoFolder As String
    Const LOGO_FOLDER As String = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Customer Logos\"
    
    Call m_Logger.LogMessage("GetLogoFileLocation", LogLevelEnum.LogDebug, "Initializing SetLogoFileLocation")
    
    ' Compile the full logo path
    FullLogoFolder = LOGO_FOLDER & CustomerName & " Logo.png"
    Call m_Logger.LogMessage("GetLogoFileLocation", LogLevelEnum.LogDebug, "Compiled logo path as " & LogoPath)
    
    ' Verify the existence of the logo file
    If Dir(FullLogoFolder) <> "" Then
        LogoPath = FullLogoFolder
    Else
        ' Placeholder for userform logic to select a file
        Call m_Logger.LogMessage("GetLogoFileLocation", LogLevelEnum.LogWarning, "Logo file not found")
        LogoPath = ""
    End If
    
    SetLogoFileLocation = LogoPath
End Function

Public Sub IInvoice_CreateWordInvoice(WordApp As Word.Application, InvoiceWriter As clsInvoiceWriter)
    Dim WordInvoice As Word.Document
    Const TEMPLATE_PATH As String = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Templates and Logbooks\Regular Individual Invoice Template.dotx"
    Const PRECILAB_BLUE As Long = 16711680
    
    Call m_Logger.LogMessage("CreateWordInvoice", LogLevelEnum.LogDebug, "Creating Word Invoice")
    
    Set WordInvoice = m_BaseInvoice.CreateWordInvoice(WordApp, TEMPLATE_PATH, InvoiceWriter)
    
    Call InvoiceWriter.WriteToTable(WordInvoice.Tables(1), m_ContactDict)
    Call InvoiceWriter.WriteToTable(WordInvoice.Tables(6), m_CCInfoDict)
    Call InvoiceWriter.WriteToBookmark(WordInvoice, "CustomerAddress", m_CustomerAddress, 10)
    If m_BaseInvoice.PaymentType = PaymentTypeEnum.PONumber Then
        Call InvoiceWriter.WriteToBookmark(WordInvoice, "POInformation", m_POInformation, 14, PRECILAB_BLUE, True)
    End If
    
    If Not m_LogoFileLocation = "" Then
        Call SetCustomerLogoFile(WordInvoice)
    End If
    
End Sub
Private Function CompileCustomerContactDict() As Scripting.Dictionary
    Dim CustomerContactDict As New Scripting.Dictionary

    CustomerContactDict.Add "1,2", m_InvoiceEmail
    CustomerContactDict.Add "3,2", m_RequestEmail

    ' Return the dictionary to the calling sub
    Set CompileCustomerContactDict = CustomerContactDict

End Function
Private Sub FillCustomerContactField(Cell As Word.Cell, Contact As clsInvoiceContact)
    Cell.Range.Text = Contact.FormattedContact
    Cell.Range.Paragraphs(1).Range.Bold = True
End Sub
Private Sub SetCustomerLogoFile(WordInvoice As Word.Document)
    Dim CompanyLogo As Word.InlineShape

    On Error GoTo ErrorHandler
    
    Call m_Logger.LogMessage("SetCustomerLogoFile", LogLevelEnum.LogDebug, "Setting the logo file")
    
    ' Check whether the bookmark exists
    If WordInvoice.Bookmarks.Exists("CompanyLogo") Then
        ' Set the logo file as an object and insert into the word document
        Set CompanyLogo = WordInvoice.Bookmarks("CompanyLogo").Range.InlineShapes.AddPicture(FileName:=m_LogoFileLocation, Linktofile:=False, savewithdocument:=True)
        ' Resize the customer logo
        Call ResizeInlineShapeByWidth(CompanyLogo, 135)
        Call m_Logger.LogMessage("SetCustomerLogoFile", LogLevelEnum.LogInfo, "Inserted and resized the customer logo")
    Else
        ' Log that the bookmark doesn't exist
        Call m_Logger.LogMessage("SetCustomerLogoFile", LogLevelEnum.LogWarning, "The customer logo bookmark doesn't exist!")
    End If
    
    Exit Sub
ErrorHandler:
    Call m_Logger.LogMessage("SetCustomerLogoFile", LogLevelEnum.LogWarning, "Error setting customer logo: " & Err.Description)
End Sub
Private Sub ResizeInlineShapeByWidth(CompanyLogo As Word.InlineShape, MaxWidth As Double)
    Dim AspectRatio As Double
    
    ' Calculate aspect ratio based on current width and height
    AspectRatio = CompanyLogo.Width / CompanyLogo.Height

    ' Set the width to the max allowed and adjust the height based on the aspect ratio
    CompanyLogo.Width = MaxWidth
    CompanyLogo.Height = MaxWidth / AspectRatio
End Sub
Public Sub IInvoice_AddInvoiceToDatabase()
    ' Pass through method for adding invoice to database
    Call m_BaseInvoice.AddInvoiceToDatabase
End Sub
Public Sub IInvoice_CreateExcelInvoice(InvoiceWriter As clsInvoiceWriter)
    ' Placeholder - currently only applicable to Customer B summary invoices
End Sub
Private Sub Class_Initialize()
    Set m_BaseInvoice = New clsInvoiceBase
    Set m_CustomerContact = New clsInvoiceContact
    Set m_CustomerCC1 = New clsInvoiceContact
    Set m_CustomerCC2 = New clsInvoiceContact
End Sub
Private Sub Class_Terminate()
    Set m_BaseInvoice = Nothing
    Set m_CustomerContact = Nothing
    Set m_CustomerCC1 = Nothing
    Set m_CustomerCC2 = Nothing
End Sub
