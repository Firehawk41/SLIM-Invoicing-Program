VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestingRequestSubmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsTestingRequestSubmission
' Version: 1.2.0
' Author: JT
' Created: ?
' LastModified: 2025-10-10
' Description: Stores properties relating to an individual testing request form
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsSLIMDictionaryManager, modEnums, clsTestingRequestCustomer, clsTestingRequestEmail, clsTestingRequestSample
' ChangeLog:
'   - 1.0.1 - 2025-06-05 - Added metadata
'   - 1.0.2 - 2025-07-23 - Exposed the first email on the results and invoice main lists
'   - 1.0.3 - 2025-08-05 - Fixed bug where I was using len instead of .count to count the items in the email collection, updated email locations
'   - 1.1.0 - 2025-08-12 - Added sub to initialize from json file
'   - 1.1.1 - 2025-08-19 - Fixed bugs for LoadFromJson
'   - 1.1.2 - 2025-09-30 - Fixed bug where the wrong cell reference was called for invoice emails
'   - 1.1.3 - 2025-10-03 - Changed the invoice email cell reference to take testing request type into account
'   - 1.2.0 - 2025-10-10 - Now checks for default invoice emails before checking the form
Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_SLIMDictManager As clsSLIMDictionaryManager
Private m_SubmissionID As Long
Private m_DatabaseEntryDate As Date
Private m_DateSubmitted As Date
Private m_DateReceived As Date
Private m_RequestType As TestingRequestTypeEnum
Private m_Customer As clsTestingRequestCustomer
Private m_CustomerAddress As String
Private m_CustomerContact As String
Private m_CustomerPhone As String
Private m_POInformation As String
Private m_CreditCardInformation As String
Private m_FileName As String
Private m_ResultsEmailMain As Collection
Private m_ResultsEmailCC As Collection
Private m_InvoiceEmailMain As Collection
Private m_InvoiceEmailCC As Collection
Private m_Samples As Collection
Public Property Get Customer() As clsTestingRequestCustomer
    Set Customer = m_Customer
End Property
Public Property Get POInformation() As String
    POInformation = m_POInformation
End Property
Public Property Get CreditCardInformation() As String
    CreditCardInformation = m_CreditCardInformation
End Property
Public Property Get Samples() As Collection
    Set Samples = m_Samples
End Property
Public Property Get FileName() As String
    FileName = m_FileName
End Property
Public Property Get RequestType() As TestingRequestTypeEnum
    RequestType = m_RequestType
End Property
Public Property Get DateReceived() As Date
    DateReceived = m_DateReceived
End Property
Public Property Get CustomerContact() As String
    CustomerContact = m_CustomerContact
End Property
Public Property Get ResultsEmailMain() As Collection
    Set ResultsEmailMain = m_ResultsEmailMain
End Property
Public Property Get ResultsEmailCC() As Collection
    Set ResultsEmailCC = m_ResultsEmailCC
End Property
Public Property Get InvoiceEmailMainFirst() As String
    Dim Email As clsTestingRequestEmail
    If m_InvoiceEmailMain.Count > 0 Then
        Set Email = m_InvoiceEmailMain(1)
        InvoiceEmailMainFirst = Email.EmailAddress
    Else
        InvoiceEmailMainFirst = ""
    End If
End Property
Public Property Get ResultsEmailMainFirst() As String
    Dim Email As clsTestingRequestEmail
    If m_ResultsEmailMain.Count > 0 Then
        Set Email = m_ResultsEmailMain(1)
        ResultsEmailMainFirst = Email.EmailAddress
    Else
        ResultsEmailMainFirst = ""
    End If
End Property
Public Sub Initialize(WkS As Worksheet, Logger As clsLoggingSystem, AccessDB As clsAccessDatabase, Optional SLIMDictManager As clsSLIMDictionaryManager = Nothing)
    Dim SampleRange As Range
    
    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    If SLIMDictManager Is Nothing Then
        Set m_SLIMDictManager = New clsSLIMDictionaryManager
        Call m_SLIMDictManager.Initialize(Logger, AccessDB)
    Else
        Set m_SLIMDictManager = SLIMDictManager
    End If
    
    Set SampleRange = GetSampleRange(WkS)
    
    ' Retrieve information for the submission
    m_RequestType = GetRequestType(WkS)
    m_DateSubmitted = GetDateSubmitted(WkS)
    m_DateReceived = GetDateReceived(WkS)
    Set m_Customer = GetCustomer(WkS)
    m_CustomerContact = GetCustomerContact(WkS)
    m_CustomerPhone = GetCustomerPhone(WkS)
    m_POInformation = GetPOInformation(WkS)
    m_CreditCardInformation = GetCreditCardInformation(WkS)
    m_FileName = WkS.Parent.Name
    
    ' Add the results and invoice emails to the submission
    Call AddEmails(WkS)
    
    ' Populate the submission with samples
    Call AddSamples(SampleRange)
    
End Sub
Private Function GetRequestType(WkS As Worksheet) As TestingRequestTypeEnum
    
    Dim RequestType As TestingRequestTypeEnum
    
    Select Case WkS.Range("D3").Value
        Case "Chemical Testing Request Form"
            RequestType = TestingRequestTypeEnum.Chemical
        Case "Water Testing Request Form"
            RequestType = TestingRequestTypeEnum.Water
        Case "Wafer Testing Request Form"
            RequestType = TestingRequestTypeEnum.Wafer
        Case Else
            Err.Raise 1001, "Get Request Type", "Testing Request not selected"
    End Select
    
    ' Return the request type to the calling sub
    GetRequestType = RequestType
    
End Function
Private Function GetDateSubmitted(WkS As Worksheet) As Date
    
    Dim DateRange As Range
    
    Select Case m_RequestType
        Case TestingRequestTypeEnum.Chemical, TestingRequestTypeEnum.Water
            Set DateRange = WkS.Range("I4")
        Case TestingRequestTypeEnum.Wafer
            Set DateRange = WkS.Range("G4")
    End Select
    
    ' Return the date to the calling sub
    If IsDate(DateRange.Value) Then
        GetDateSubmitted = DateRange.Value
    End If
    
End Function
Private Function GetDateReceived(WkS As Worksheet) As Date
    
    Dim DateRange As Range
    
    Select Case m_RequestType
        Case TestingRequestTypeEnum.Chemical
            Set DateRange = WkS.Range("P5")
        Case TestingRequestTypeEnum.Water
            Set DateRange = WkS.Range("S5")
        Case TestingRequestTypeEnum.Wafer
            Set DateRange = WkS.Range("L5")
    End Select
    
    ' Return the date to the calling sub
    If IsDate(DateRange.Value) Then
        GetDateReceived = DateRange.Value
    End If
            
End Function
Private Function GetCustomer(WkS As Worksheet) As clsTestingRequestCustomer

    Dim NameOnForm As String
    Dim AddressOnForm As String
    Dim Customer As New clsTestingRequestCustomer
    
    NameOnForm = WkS.Range("C10").Value
    AddressOnForm = WkS.Range("C11").Value & " " & WkS.Range("C12").Value
    
    Call Customer.SetFromDictionary(NameOnForm, AddressOnForm, m_Logger, m_AccessDB, m_SLIMDictManager.CustomerDict)

    ' Return the customer to the calling sub
    Set GetCustomer = Customer
    
End Function
Private Function GetCustomerContact(WkS As Worksheet) As String

    Dim CustomerContact As String
    
    CustomerContact = WkS.Range("C13").Value
    
    ' Return the customer to the calling sub
    GetCustomerContact = CustomerContact
    
End Function
Private Function GetCustomerPhone(WkS As Worksheet) As String

    Dim CustomerPhone As String
    
    ' Check which version of the testing request is being used and assign the phone number
    If WkS.Range("B14").Value = "Phone :" Then
        CustomerPhone = WkS.Range("C14").Value
    ElseIf WkS.Range("B15").Value = "Phone:" Then
        CustomerPhone = WkS.Range("C15").Value
    End If
    
    ' Return the customer to the calling sub
    GetCustomerPhone = CustomerPhone
    
End Function
Private Function GetPOInformation(WkS As Worksheet) As String
    Dim POInformation As String
    Dim RevisionLetter As String
    
    Const CHEM_DOC_REVISION_INFO As String = "P49"
    Const WATER_DOC_REVISION_INFO As String = "S51"
    Const WAFER_DOC_REVISION_INFO As String = "L45"
    Const ADDITIONAL_PURCHASING_INFORMATION As String = "F13"
    Const CHEM_PO_NUMBER_OLD As String = "I10"
    Const CHEM_PO_NUMBER_NEW As String = "H10"
    Const WATER_PO_NUMBER_OLD As String = "J10"
    Const WATER_PO_NUMBER_NEW As String = "H10"
    Const WAFER_PO_NUMBER As String = "G10"
    
    Const REVISION_THRESHOLD_CHEM As String = "T"
    Const REVISION_THRESHOLD_WATER As String = "I"
    
    ' Determine revision letter
    RevisionLetter = GetRevisionLetter(WkS, m_RequestType, CHEM_DOC_REVISION_INFO, WATER_DOC_REVISION_INFO, WAFER_DOC_REVISION_INFO)
    
    Select Case m_RequestType
        Case TestingRequestTypeEnum.Chemical
            
            If (m_Customer.CustomerName = "CUSTOMER A TX" Or m_Customer.CustomerName = "CUSTOMER A CA") _
                And UCase(RevisionLetter) >= REVISION_THRESHOLD_CHEM _
                And WkS.Range(ADDITIONAL_PURCHASING_INFORMATION).Value <> "" Then
                ' Customer A PO numbers are the request numbers and are in the 'additional purchasing information' cell
                POInformation = WkS.Range(ADDITIONAL_PURCHASING_INFORMATION).Value
            ElseIf UCase(RevisionLetter) >= REVISION_THRESHOLD_CHEM Then
                POInformation = WkS.Range(CHEM_PO_NUMBER_NEW).Value
            Else
                POInformation = WkS.Range(CHEM_PO_NUMBER_OLD).Value
            End If
            
        Case TestingRequestTypeEnum.Water
            If UCase(RevisionLetter) >= REVISION_THRESHOLD_WATER Then
                POInformation = WkS.Range(WATER_PO_NUMBER_NEW).Value
            Else
                POInformation = WkS.Range(WATER_PO_NUMBER_OLD).Value
            End If
        Case TestingRequestTypeEnum.Wafer
            POInformation = WkS.Range(WAFER_PO_NUMBER).Value
    End Select
    
    ' Sterilize the PO
    POInformation = Replace(POInformation, vbCr, "")
    POInformation = Replace(POInformation, vbLf, "")
    POInformation = Trim(POInformation)

    
    ' Return the PO information to the calling sub
    GetPOInformation = POInformation
    
End Function
Private Function GetRevisionLetter(WkS As Worksheet, RequestType As TestingRequestTypeEnum, ChemCell As String, WaterCell As String, WaferCell As String) As String
    Dim RawText As String
    Dim CellAddress As String

    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical
            CellAddress = ChemCell
        Case TestingRequestTypeEnum.Water
            CellAddress = WaterCell
        Case TestingRequestTypeEnum.Wafer
            CellAddress = WaferCell
    End Select

    RawText = WkS.Range(CellAddress).Value
    If InStr(UCase(RawText), "REV ") > 0 Then
        GetRevisionLetter = Trim(Mid(RawText, InStr(UCase(RawText), "REV ") + 4, 1))
    Else
        GetRevisionLetter = ""
    End If
End Function
Private Function GetCreditCardInformation(WkS As Worksheet) As String
    
    Dim CreditCardInformation As String
    Dim CreditCardRange As Range
    
    Select Case m_RequestType
        Case TestingRequestTypeEnum.Chemical
            Set CreditCardRange = WkS.Range("I11")
        Case TestingRequestTypeEnum.Water
            Set CreditCardRange = WkS.Range("J11")
        Case TestingRequestTypeEnum.Wafer
            Set CreditCardRange = WkS.Range("G11")
    End Select
    
    ' Get the value from the determined range
    CreditCardInformation = CreditCardRange.Value & IIf(CreditCardRange.Offset(1).Value = "", "", " " & CreditCardRange.Offset(1))

    
    ' Return the Credit Card information to the calling sub
    GetCreditCardInformation = CreditCardInformation
    
End Function
Private Function GetSampleRange(WkS As Worksheet) As Range
    
    Dim FirstRow As Integer
    Dim LastRow As Long
    Dim LastColumn As Long
    
    ' Determine the first row based on whether "Sample ID" is present in cell B20
    If WkS.Range("B20") = "Sample ID" Then
        FirstRow = 22
    Else
        FirstRow = 21
    End If
    
    LastRow = WkS.Cells(Rows.Count, 2).End(xlUp).Row - 11
    
    LastColumn = WkS.Cells(FirstRow - 2, WkS.Columns.Count).End(xlToLeft).Column - 1

    ' Compile the range and return to the calling sub
    Set GetSampleRange = WkS.Range(WkS.Cells(FirstRow, 2), WkS.Cells(LastRow, LastColumn))
    
End Function
Private Sub AddEmails(WkS As Worksheet)
    
    Dim InvoiceMainCellReference As String
    Dim InvoiceCCCellReference As String
    
    ' Specify where the emails are found by testing request type
    Select Case m_RequestType
        Case TestingRequestTypeEnum.Chemical
            InvoiceMainCellReference = "N10"
            InvoiceCCCellReference = "N11"
        Case TestingRequestTypeEnum.Wafer
            InvoiceMainCellReference = "J10"
            InvoiceCCCellReference = "J11"
        Case TestingRequestTypeEnum.Water
            InvoiceMainCellReference = "P10"
            InvoiceCCCellReference = "P11"
        Case Else
            Call m_Logger.LogMessage("clsTestingRequestSubmission.AddEmails", LogLevelEnum.LogError, "Testing request type not recognized")
            Err.Raise 2001, "clsTestingRequestSubmission.AddEmails", "Testing request type not recognized"
    End Select
    
    
    ' Add emails by type
    
    Call AddEmailsByType(WkS.Range("C15").Value, EmailTypeEnum.ResultsMain)
    Call AddEmailsByType(WkS.Range("C17").Value, EmailTypeEnum.ResultsCC)
    If Not AddDefaultInvoiceEmails() Then
        Call AddEmailsByType(WkS.Range(InvoiceMainCellReference).Value, EmailTypeEnum.InvoiceMain)
    End If
    Call AddEmailsByType(WkS.Range(InvoiceCCCellReference).Value, EmailTypeEnum.InvoiceCC)
    
    
End Sub

Private Sub AddEmailsByType(EmailString As String, EmailType As EmailTypeEnum)
    Dim EmailCollection As Collection
    Dim Email As clsTestingRequestEmail
    Dim iIndex As Integer
    
    ' Split the email string and create the list of email objects
    Set EmailCollection = CreateEmailList(EmailString)
    
    ' Add valid emails to the correct submission list
    For iIndex = 1 To EmailCollection.Count
        Set Email = EmailCollection(iIndex)
        Select Case EmailType
            Case EmailTypeEnum.ResultsMain
                m_ResultsEmailMain.Add Email
            Case EmailTypeEnum.ResultsCC
                m_ResultsEmailCC.Add Email
            Case EmailTypeEnum.InvoiceMain
                m_InvoiceEmailMain.Add Email
            Case EmailTypeEnum.InvoiceCC
                m_InvoiceEmailCC.Add Email
        End Select
    Next iIndex

End Sub
Private Function CreateEmailList(EmailString As String) As Collection
    Dim EmailCollection As Collection
    Set EmailCollection = New Collection
    Dim EmailArray() As String
    Dim iIndex As Integer
    Dim Email As clsTestingRequestEmail
    
    ' Replace different delimiters with a comma
    EmailString = Replace(EmailString, ";", ",")
    EmailString = Replace(EmailString, " ", ",")
    EmailString = Replace(EmailString, vbNewLine, ",")
    
    ' Split the string by comma
    EmailArray = Split(EmailString, ",")
    
    ' Iterate through the array and create email objects
    For iIndex = LBound(EmailArray) To UBound(EmailArray)
        If Trim(EmailArray(iIndex)) <> "" Then
            Set Email = New clsTestingRequestEmail
            Call Email.Initialize(Trim(EmailArray(iIndex)), m_Logger)
            
            ' Add only if it's a valid email
            If Email.IsValidEmail Then
                EmailCollection.Add Email
            End If
        End If
    Next iIndex
    
    ' Return the collection of emails
    Set CreateEmailList = EmailCollection
        
End Function
Private Function AddDefaultInvoiceEmails() As Boolean
    Dim SQLQuery As String
    Dim Results As Variant
    Dim EmailStr As String
    
    Call m_Logger.LogMessage("clsTestingRequestSubmission.AddDefaultInvoiceEmails", LogLevelEnum.LogDebug, "Checking for default invoice emails")
    
    If Int(Customer.CustomerID) And Customer.CustomerID > 0 Then
        SQLQuery = "SELECT contact_email FROM invoice_default_contacts WHERE customer_id = " & Customer.CustomerID & " AND is_default = TRUE AND is_active = TRUE"
        Results = m_AccessDB.ExecuteQuery(SQLQuery)
        If UBound(Results) <> -1 Then
            Call AddEmailsByType(Join2DArray(Results), EmailTypeEnum.InvoiceMain)
            AddDefaultInvoiceEmails = True
        Else
            AddDefaultInvoiceEmails = False
        End If
    Else
        AddDefaultInvoiceEmails = False
    End If
    
    Call m_Logger.LogMessage("clsTestingRequestSubmission.AddDefaultInvoiceEmails", LogLevelEnum.LogDebug, "Default invoice emails " & IIf(AddDefaultInvoiceEmails, "", "not ") & "found.")
    
End Function
Private Function Join2DArray(TwoDArray As Variant) As String
    Dim iIndex As Long
    Dim OneDArray() As String
    
    ReDim OneDArray(LBound(TwoDArray, 1) To UBound(TwoDArray, 1))
    
    For iIndex = LBound(TwoDArray, 1) To UBound(TwoDArray, 1)
        OneDArray(iIndex) = TwoDArray(iIndex, 0)
    Next iIndex
    
    Join2DArray = Join(OneDArray, ",")
End Function


Private Sub AddSamples(SampleRange As Range)

    Dim SampleCell As Range
    Dim Sample As clsTestingRequestSample

    ' Loop through the cells in the first column of the sample range
    For Each SampleCell In SampleRange.Columns(1).Cells
        If Not IsEmpty(SampleCell) Then
            ' Create and initialize the sample object
            Set Sample = New clsTestingRequestSample
            Call Sample.Initialize(SampleCell, SampleRange, m_RequestType, m_SLIMDictManager, m_Logger, m_AccessDB)
            ' Add the sample to the submission
            m_Samples.Add Sample
        End If
    Next SampleCell

End Sub
Public Sub LoadFromJson(Json As Dictionary, Logger As clsLoggingSystem)
    Call Logger.LogMessage("clsTestingRequestSubmission.LoadFromJson", LogLevelEnum.LogDebug, "Loading Json data into Submission object")
    
    Set m_Logger = Logger
    
    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler
    
    m_SubmissionID = Json("id")
    m_DateReceived = CDate(Replace(Json("date_received"), "T", " "))
    m_RequestType = ParseRequestType(Json("request_type"))
    m_DateSubmitted = CDate(Replace(Json("date_submitted"), "T", " "))
    If Json.Exists("form_contact") Then
        If Json("form_contact").Exists("name") Then m_CustomerContact = Json("form_contact")("name")
        If Json("form_contact").Exists("phone") Then m_CustomerPhone = Json("form_contact")("phone")
    End If
    m_POInformation = Json("po_information")
    m_CreditCardInformation = Json("credit_card_information")
    m_FileName = Json("file_name")
    
    
    Set m_Customer = New clsTestingRequestCustomer
    Call m_Customer.LoadFromJson(Json("form_customer"), Logger)
    
    Dim SampleJson As Dictionary
    Dim Sample As clsTestingRequestSample
    For Each SampleJson In Json("samples")
        Set Sample = New clsTestingRequestSample
        Call Sample.LoadFromJson(SampleJson, Logger)
        m_Samples.Add Sample
    Next SampleJson
    
    Call m_Logger.LogMessage("clsTestingRequestSubmission.LoadFromJson", LogLevelEnum.LogDebug, "Completed loading Json data into Submission object.")
    Exit Sub
ErrorHandler:
    Dim ErrorMessage As String
    ErrorMessage = "Error encountered in clsTestingRequestSubmission.LoadFromJson where the file name is: " & m_FileName
    Call m_Logger.LogMessage("clsTestingRequestSubmission.LoadFromJson", LogLevelEnum.LogError, ErrorMessage)
    Err.Raise vbObjectError + 1000, "clsTestingRequestSubmission.LoadFromJson", ErrorMessage
End Sub
Private Function ParseRequestType(RequestType As String) As TestingRequestTypeEnum
    Select Case RequestType
        Case "Chemical"
            m_RequestType = TestingRequestTypeEnum.Chemical
        Case "Water"
            m_RequestType = TestingRequestTypeEnum.Water
        Case "Wafer"
            m_RequestType = TestingRequestTypeEnum.Wafer
        Case Else
            Err.Raise vbObjectError + 1000, "clsTestingRequestSubmission.LoadFromJson", _
                      "Unknown request type: " & RequestType
    End Select
End Function
Public Sub Class_Initialize()
    ' Initialize the Sample collection
    Set m_Samples = New Collection
    Set m_ResultsEmailMain = New Collection
    Set m_ResultsEmailCC = New Collection
    Set m_InvoiceEmailMain = New Collection
    Set m_InvoiceEmailCC = New Collection

End Sub

Private Sub Class_Terminate()
    Set m_Samples = Nothing
    Set m_ResultsEmailMain = Nothing
    Set m_ResultsEmailCC = Nothing
    Set m_InvoiceEmailMain = Nothing
    Set m_InvoiceEmailCC = Nothing
    Set m_Customer = Nothing
End Sub
