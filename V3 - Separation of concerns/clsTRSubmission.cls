VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTRSubmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsTRSubmission
' Author: JT
' Created: 2026-04-10
' Description: Immutable-after-load entity representing one testing
'              request submission. Replaces clsTestingRequestSubmission.
' DependsOn: clsLoggingSystem, modEnums, modUtilities
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release
'
' ================================================================
' LAYER: Entity - no DB, no Worksheet, no service dependencies.
'
' LIFECYCLE:
'   1.  New clsTRSubmission
'   2.  .Initialize(Logger)
'   3.  .LoadFromDict(D)          <- scalar fields only
'   4.  .AddResultsEmail / .AddInvoiceEmail x N
'
' All steps 3-6 are performed by clsSubmissionFormParser.
' Callers should only ever receive a fully-hydrated object from
' clsSubmissionService - never call these methods directly.
'
' ID CONTRACT:
'   ID = 0 means unsaved. If DB persistence is added later,
'   clsSubmissionService will call .SetInternalID(newID) once.
'
' Dict keys for LoadFromDict:
'   submission_id, date_submitted, date_received,
'   request_type (Long = enum), customer_contact, customer_phone,
'   po_information, credit_card_information,
'   file_name, service_date, download_date, form_customer_name, form_customer_address
'   form_customer_address_2
'
' EFFECTIVE SERVICE DATE:
'   Computed lazily from m_ServiceDate. If the sentinel date
'   (1901-01-01) was stored, the fallback calculation runs over
'   the samples collection. Pure computation - no I/O.
' ================================================================

Private m_Logger                As clsLoggingSystem
Private m_ID                    As Long
Private m_DateSubmitted         As Date
Private m_DateReceived          As Date
Private m_RequestType           As TestingRequestTypeEnum
Private m_CustomerContact       As String
Private m_CustomerPhone         As String
Private m_POInformation         As String
Private m_CreditCardInfo        As String
Private m_FileName              As String
Private m_ServiceDate           As Date
Private m_DownloadDate          As Date
Private m_FormCustomerName      As String
Private m_FormCustomerAddress   As String
Private m_FormCustomerAddress2  As String

Private m_CustomerID            As Long
Private m_Samples               As Collection   ' clsSample

' Email collections store plain address strings.
' Parser validates before adding; entity just holds them.
Private m_ResultsEmailMain      As Collection   ' String
Private m_ResultsEmailCC        As Collection   ' String
Private m_InvoiceEmailMain      As Collection   ' String
Private m_InvoiceEmailCC        As Collection   ' String

Private Const SENTINEL_DATE   As String = "1901-01-01"

' ----------------------------------------------------------------
' Initialization
' ----------------------------------------------------------------
Public Sub Initialize(Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init

    Set m_Logger = Logger
    Set m_Samples = New Collection
    Set m_ResultsEmailMain = New Collection
    Set m_ResultsEmailCC = New Collection
    Set m_InvoiceEmailMain = New Collection
    Set m_InvoiceEmailCC = New Collection
End Sub

' ----------------------------------------------------------------
' Scalar hydration  (called by parser after Initialize)
' ----------------------------------------------------------------
' Dict keys: submission_id, customer_id, date_submitted, date_received,
'            request_type, customer_contact, customer_phone,
'            po_information, credit_card_information,
'            file_name, service_date, download_date, submission_ref
'
Public Sub LoadFromDict(D As Scripting.Dictionary)
    AssertInit "LoadFromDict"
    Debug.Assert Not D Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    RequireKey D, "submission_id"
    RequireKey D, "customer_id"
    RequireKey D, "date_submitted"
    RequireKey D, "date_received"
    RequireKey D, "request_type"
    RequireKey D, "customer_contact"
    RequireKey D, "customer_phone"
    RequireKey D, "po_information"
    RequireKey D, "credit_card_information"
    RequireKey D, "file_name"
    RequireKey D, "service_date"
    RequireKey D, "download_date"
    RequireKey D, "form_customer_name"
    RequireKey D, "form_customer_address"
    RequireKey D, "form_customer_address_2"

    m_ID = CLng(D("submission_id"))
    m_CustomerID = CLng(D("customer_id"))
    m_DateSubmitted = CDate(D("date_submitted"))
    m_DateReceived = CDate(D("date_received"))
    m_RequestType = CInt(D("request_type"))           ' stored as enum Long
    m_CustomerContact = CStr(D("customer_contact"))
    m_CustomerPhone = CStr(D("customer_phone"))
    m_POInformation = CStr(D("po_information"))
    m_CreditCardInfo = CStr(D("credit_card_information"))
    m_FileName = CStr(D("file_name"))
    m_ServiceDate = CDate(D("service_date"))
    m_DownloadDate = CDate(D("download_date"))
    m_FormCustomerName = CStr(D("form_customer_name"))
    m_FormCustomerAddress = CStr(D("form_customer_address"))
    m_FormCustomerAddress2 = CStr(D("form_customer_address_2"))

    Exit Sub
EH:                                                         HandleError "LoadFromDict", "Keys present: " & Join(D.Keys, ", ")
End Sub

' ----------------------------------------------------------------
' Child object population  (called by parser/service after LoadFromDict)
' ----------------------------------------------------------------

Public Sub AddSample(S As clsTRSample)
    AssertInit "AddSample"
    Debug.Assert Not S Is Nothing
    m_Samples.Add S
End Sub

Public Sub AddResultsEmail(Address As String, IsCC As Boolean)
    AssertInit "AddResultsEmail"
    If IsCC Then
        m_ResultsEmailCC.Add Address
    Else
        m_ResultsEmailMain.Add Address
    End If
End Sub

Public Sub AddInvoiceEmail(Address As String, IsCC As Boolean)
    AssertInit "AddInvoiceEmail"
    If IsCC Then
        m_InvoiceEmailCC.Add Address
    Else
        m_InvoiceEmailMain.Add Address
    End If
End Sub

' ----------------------------------------------------------------
' ID sealing  (called by service after INSERT, when persistence added)
' ----------------------------------------------------------------
Friend Sub SetInternalID(NewID As Long)
    AssertInit "SetInternalID"
    Debug.Assert NewID > 0
    Debug.Assert m_ID = 0

    If Not m_Logger.DebugMode Then On Error GoTo EH

    If NewID <= 0 Then Err.Raise vbObjectError + 400, "clsTRSubmission.SetInternalID", "ID must be positive, got: " & NewID
    If m_ID <> 0 Then Err.Raise vbObjectError + 401, "clsTRSubmission.SetInternalID", "ID already set to: " & m_ID

    m_ID = NewID
    Exit Sub
EH:
    HandleError "SetInternalID", "Failed to set ID: " & NewID
End Sub

' ----------------------------------------------------------------
' Read-only properties
' ----------------------------------------------------------------
Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Get DateSubmitted() As Date
    DateSubmitted = m_DateSubmitted
End Property

Public Property Get DateReceived() As Date
    DateReceived = m_DateReceived
End Property

Public Property Get RequestType() As TestingRequestTypeEnum
    RequestType = m_RequestType
End Property

Public Property Get CustomerID() As Long
    CustomerID = m_CustomerID
End Property

Public Property Get CustomerContact() As String
    CustomerContact = m_CustomerContact
End Property

Public Property Get CustomerPhone() As String
    CustomerPhone = m_CustomerPhone
End Property

Public Property Get POInformation() As String
    POInformation = m_POInformation
End Property

Public Property Get CreditCardInformation() As String
    CreditCardInformation = m_CreditCardInfo
End Property

Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Get ServiceDate() As Date
    ServiceDate = m_ServiceDate
End Property

Public Property Get DownloadDate() As Date
    DownloadDate = m_DownloadDate
End Property

Public Property Get Samples() As Collection
    Set Samples = m_Samples
End Property

Public Property Get ResultsEmailMain() As Collection
    Set ResultsEmailMain = m_ResultsEmailMain
End Property

Public Property Get ResultsEmailCC() As Collection
    Set ResultsEmailCC = m_ResultsEmailCC
End Property

Public Property Get InvoiceEmailMain() As Collection
    Set InvoiceEmailMain = m_InvoiceEmailMain
End Property

Public Property Get InvoiceEmailCC() As Collection
    Set InvoiceEmailCC = m_InvoiceEmailCC
End Property

Public Property Get ResultsEmailMainFirst() As String
    If m_ResultsEmailMain.Count > 0 Then ResultsEmailMainFirst = m_ResultsEmailMain(1)
End Property

Public Property Get InvoiceEmailMainFirst() As String
    If m_InvoiceEmailMain.Count > 0 Then InvoiceEmailMainFirst = m_InvoiceEmailMain(1)
End Property

Public Property Get FormCustomerName() As String
    FormCustomerName = m_FormCustomerName
End Property

Public Property Get FormCustomerAddress() As String
    FormCustomerAddress = m_FormCustomerAddress
End Property

' ----------------------------------------------------------------
' Computed: EffectiveServiceDate
' Pure calculation over own state - no I/O.
' ----------------------------------------------------------------
Public Property Get EffectiveServiceDate() As Date
    If IsServiceDateMissing Then
        EffectiveServiceDate = GetFallbackDateOfService(m_DateReceived, GetMaxProcessingTime())
    Else
        EffectiveServiceDate = m_ServiceDate
    End If
End Property

' ----------------------------------------------------------------
' Private: EffectiveServiceDate helpers
' ----------------------------------------------------------------
Private Property Get IsServiceDateMissing() As Boolean
    IsServiceDateMissing = (m_ServiceDate = CDate(SENTINEL_DATE))
End Property

Private Function GetMaxProcessingTime() As ProcessingTimeEnum
    Dim S           As clsTRSample
    Dim MaxDays     As Long
    Dim MaxPT       As ProcessingTimeEnum
    Dim SDays       As Long

    MaxDays = -1

    For Each S In m_Samples
        SDays = modUtilities.ProcessingTimeToDays(S.ProcessingTime)
        If SDays > MaxDays Then
            MaxDays = SDays
            MaxPT = S.ProcessingTime
        End If
    Next S

    GetMaxProcessingTime = MaxPT
End Function

Private Function GetFallbackDateOfService(ReceivedDate As Date, ProcessingTime As ProcessingTimeEnum) As Date
    Dim Days As Long
    Days = modUtilities.ProcessingTimeToDays(ProcessingTime)
    If Days = 0 Then
        GetFallbackDateOfService = ReceivedDate
    Else
        GetFallbackDateOfService = AddWorkingDays(ReceivedDate, Days)
    End If
End Function

Private Function AddWorkingDays(StartDate As Date, DaysToAdd As Long) As Date
    Dim Result As Date
    Dim i      As Long
    Result = StartDate
    For i = 1 To DaysToAdd
        Result = DateAdd("d", 1, Result)
        Do While Weekday(Result, vbMonday) > 5
            Result = DateAdd("d", 1, Result)
        Loop
    Next i
    AddWorkingDays = Result
End Function

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub RequireKey(D As Scripting.Dictionary, Key As String)
    Debug.Assert D.Exists(Key)
    If Not D.Exists(Key) Then
        Err.Raise vbObjectError + 200, "clsTRSubmission.LoadFromDict", _
    "Missing required key: " & Key
    End If
End Sub

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsTRSubmission." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Optional Context As String = "")
    m_Logger.LogError "clsTRSubmission." & Src, Err.Number, Err.Description & " | " & Context
End Sub
