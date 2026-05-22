VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTRSubmissionFormParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsSubmissionFormParser
' Author: JT
' Created: 2026-04-10
' Description: Reads an Excel testing request worksheet and returns
'              a fully-hydrated clsSubmission. Owns ALL cell-reading
'              and form-structure logic for submissions.
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsCustomerService,
'            clsChemicalService, clsAnalysisService, clsElementService,
'            clsTRSubmission, clsTRSampleFormParser
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Extracted from clsTestingRequestSubmission.Initialize
'
' ================================================================
' LAYER: Parser - form reading and service lookups only.
'
' DB ACCESS:
'   The one query here (default invoice contacts) will move to
'   clsSubmissionRepository when DB persistence is added.
'   TODO: move AddDefaultInvoiceEmails SQL to clsSubmissionRepository.
'
' CELL MAP (request-type dependent, all 1-based):
'
'   Field             Chemical   Water   Wafer
'   ---------------   --------   -----   -----
'   Date submitted    I4         I4      G4
'   Date received     P5         S5      L5
'   Invoice email     N10/N11    P10/P11 J10/J11
'   Credit card       I11        J11     G11
'   Doc revision      P49        S51     L45
'   PO number (old)   I10        J10     -
'   PO number (new)   H10        H10     G10
'
'   Fixed cells (all types):
'     Customer name   C10
'     Customer addr   C11
'     Customer addr 2 C12
'     Contact name    C13
'     Phone           C14 or C15 (version-dependent)
'     Results main    C15
'     Results CC      C17
' ================================================================

Private m_Logger      As clsLoggingSystem
Private m_AccessDB    As clsAccessDatabase
Private m_CustomerSvc As clsCustomerService
Private m_SampleParser As clsTRSampleFormParser
Private m_InputResolver As clsTRFormInputResolver

' ----------------------------------------------------------------
' Initialization
' ----------------------------------------------------------------
Public Sub Initialize(AccessDB As clsAccessDatabase, _
                      CustomerSvc As clsCustomerService, _
                      ChemicalSvc As clsChemicalService, _
                      AnalysisSvc As clsAnalysisService, _
                      ElementSvc As clsElementService, _
                      InputResolver As clsTRFormInputResolver, _
                      Logger As clsLoggingSystem)
    Debug.Assert Not AccessDB Is Nothing
    Debug.Assert Not CustomerSvc Is Nothing
    Debug.Assert Not ChemicalSvc Is Nothing
    Debug.Assert Not AnalysisSvc Is Nothing
    Debug.Assert Not ElementSvc Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert Not InputResolver Is Nothing
    Debug.Assert m_Logger Is Nothing         ' enforce single-init

    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    Set m_CustomerSvc = CustomerSvc
    Set m_InputResolver = InputResolver

    Set m_SampleParser = New clsTRSampleFormParser
    m_SampleParser.Initialize ChemicalSvc, AnalysisSvc, ElementSvc, InputResolver, Logger
End Sub

' ----------------------------------------------------------------
' Public API
' ----------------------------------------------------------------
Public Function BuildFromWorksheet(WkS As Worksheet) As clsTRSubmission
    Dim RequestType As TestingRequestTypeEnum
    Dim C As clsCustomer
    Dim FormCustomerName As String
    Dim FormCustomerAddress As String
    Dim FormCustomerAddress2 As String
    Dim CustomerID   As Long
    Dim CustomerName As String
    Dim D As Scripting.Dictionary
    Dim S As New clsTRSubmission
    
    AssertInit "BuildFromWorksheet"
    Debug.Assert Not WkS Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    ' 1. Determine request type first - everything else depends on it.
    RequestType = ReadRequestType(WkS)

    ' 2. Resolve Customer first - ID is required in the scalar dict
    Set C = ResolveCustomer(WkS, FormCustomerName, FormCustomerAddress, FormCustomerAddress2)
    If Not C Is Nothing Then
        CustomerID = C.ID
        CustomerName = C.Name
    End If

    ' 3. Build scalar dict and hydrate the entity.
    Set D = BuildSubmissionDict(WkS, RequestType, CustomerID, FormCustomerName, FormCustomerAddress, FormCustomerAddress2)
    S.Initialize m_Logger
    S.LoadFromDict D

    ' 4. Parse and attach emails.
    AttachEmails S, WkS, RequestType

    ' 5. Parse and attach samples.
    AttachSamples S, WkS, RequestType

    Set BuildFromWorksheet = S
    Exit Function
EH:
    HandleError "BuildFromWorksheet", "File: " & WkS.Parent.Name
End Function

' ----------------------------------------------------------------
' Private: scalar dict builder
' ----------------------------------------------------------------
Private Function BuildSubmissionDict(WkS As Worksheet, _
                                     RequestType As TestingRequestTypeEnum, _
                                     CustomerID As Long, _
                                     FormCustomerName As String, _
                                     FormCustomerAddress As String, _
                                     FormCustomerAddress2 As String) As Scripting.Dictionary
    Dim D As New Scripting.Dictionary

    If Not m_Logger.DebugMode Then On Error GoTo EH

    D.Add "submission_id", 0
    D.Add "customer_id", CustomerID
    D.Add "date_submitted", ReadDateSubmitted(WkS, RequestType)
    D.Add "date_received", ReadDateReceived(WkS, RequestType)
    D.Add "request_type", CLng(RequestType)
    D.Add "customer_contact", CStr(WkS.Range("C13").Value)
    D.Add "customer_phone", ReadPhone(WkS)
    D.Add "po_information", ReadPOInformation(WkS, RequestType)
    D.Add "credit_card_information", ReadCreditCardInfo(WkS, RequestType)
    D.Add "file_name", WkS.Parent.Name
    D.Add "service_date", ExtractServiceDate(WkS.Parent.Name)
    D.Add "download_date", FileDateTime(WkS.Parent.FullName)
    D.Add "form_customer_name", FormCustomerName
    D.Add "form_customer_address", FormCustomerAddress
    D.Add "form_customer_address_2", FormCustomerAddress2

    Set BuildSubmissionDict = D
    Exit Function
EH:
    HandleError "BuildSubmissionDict", "Failed to build submission dict for " & WkS.Parent.Name
End Function

' ----------------------------------------------------------------
' Private: request type
' ----------------------------------------------------------------
Private Function ReadRequestType(WkS As Worksheet) As TestingRequestTypeEnum
    Select Case WkS.Range("D3").Value
        Case "Chemical Testing Request Form"
            ReadRequestType = TestingRequestTypeEnum.Chemical
        Case "Water Testing Request Form"
            ReadRequestType = TestingRequestTypeEnum.Water
        Case "Wafer Testing Request Form"
            ReadRequestType = TestingRequestTypeEnum.Wafer
        Case Else
            Err.Raise vbObjectError + 500, "clsSubmissionFormParser.ReadRequestType", _
                "Unrecognised form title in D3: " & WkS.Range("D3").Value
    End Select
End Function

' ----------------------------------------------------------------
' Private: date readers
' ----------------------------------------------------------------
Private Function ReadDateSubmitted(WkS As Worksheet, RequestType As TestingRequestTypeEnum) As Date
    Dim Cell As Range
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical, TestingRequestTypeEnum.Water
            Set Cell = WkS.Range("I4")
        Case TestingRequestTypeEnum.Wafer
            Set Cell = WkS.Range("G4")
    End Select
    If IsDate(Cell.Value) Then ReadDateSubmitted = CDate(Cell.Value)
End Function

Private Function ReadDateReceived(WkS As Worksheet, RequestType As TestingRequestTypeEnum) As Date
    Dim Cell As Range
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical: Set Cell = WkS.Range("P5")
        Case TestingRequestTypeEnum.Water:    Set Cell = WkS.Range("S5")
        Case TestingRequestTypeEnum.Wafer:    Set Cell = WkS.Range("L5")
    End Select
    If IsDate(Cell.Value) Then ReadDateReceived = CDate(Cell.Value)
End Function

' ----------------------------------------------------------------
' Private: Customer
' ----------------------------------------------------------------
Private Function ResolveCustomer(WkS As Worksheet, ByRef NameOnForm As String, ByRef AddressOnForm As String, ByRef AddressOnForm2 As String) As clsCustomer
    ' Note that blank strings are passed from the calling sub so that NameOnForm and AddressOnForm can be passed back
    Dim C As clsCustomer

    If Not m_Logger.DebugMode Then On Error GoTo EH

    NameOnForm = CStr(Trim(WkS.Range("C10").Value))
    AddressOnForm = CStr(Trim(WkS.Range("C11").Value))
    AddressOnForm2 = CStr(Trim(WkS.Range("C12").Value))

    Set C = m_InputResolver.ResolveCustomer(NameOnForm, AddressOnForm, AddressOnForm2)

    If C Is Nothing Then
        Err.Raise vbObjectError + 502, "clsTRSubmissionFormParser.ResolveCustomer", _
            "Could not resolve customer '" & NameOnForm & "' - user cancelled or no match found."
    End If

    Set ResolveCustomer = C
    Exit Function
EH:
    HandleError "ResolveCustomer"
End Function

' ----------------------------------------------------------------
' Private: phone
' ----------------------------------------------------------------
Private Function ReadPhone(WkS As Worksheet) As String
    If WkS.Range("B14").Value = "Phone :" Then
        ReadPhone = CStr(WkS.Range("C14").Value)
    ElseIf WkS.Range("B15").Value = "Phone:" Then
        ReadPhone = CStr(WkS.Range("C15").Value)
    End If
End Function

' ----------------------------------------------------------------
' Private: PO information  (Customer A/revision/version-aware)
' ----------------------------------------------------------------
Private Function ReadPOInformation(WkS As Worksheet, _
                                    RequestType As TestingRequestTypeEnum) As String
    Const CHEM_REV_CELL  As String = "P49"
    Const WATER_REV_CELL As String = "S51"
    Const WAFER_REV_CELL As String = "L45"
    Const ADD_PO_CELL    As String = "F13"
    Const CHEM_OLD_PO    As String = "I10"
    Const CHEM_NEW_PO    As String = "H10"
    Const WATER_OLD_PO   As String = "J10"
    Const WATER_NEW_PO   As String = "H10"
    Const WAFER_PO       As String = "G10"
    Const CHEM_REV_THRESH  As String = "T"
    Const WATER_REV_THRESH As String = "I"

    Dim RevLetter As String
    RevLetter = ReadRevisionLetter(WkS, RequestType, CHEM_REV_CELL, WATER_REV_CELL, WAFER_REV_CELL)

    Dim POInfo As String
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical
            If UCase(RevLetter) >= CHEM_REV_THRESH Then
                POInfo = CStr(WkS.Range(CHEM_NEW_PO).Value)
            Else
                POInfo = CStr(WkS.Range(CHEM_OLD_PO).Value)
            End If
        Case TestingRequestTypeEnum.Water
            If UCase(RevLetter) >= WATER_REV_THRESH Then
                POInfo = CStr(WkS.Range(WATER_NEW_PO).Value)
            Else
                POInfo = CStr(WkS.Range(WATER_OLD_PO).Value)
            End If
        Case TestingRequestTypeEnum.Wafer
            POInfo = CStr(WkS.Range(WAFER_PO).Value)
    End Select

    POInfo = Trim$(Replace(Replace(POInfo, vbCr, ""), vbLf, ""))
    ReadPOInformation = POInfo
End Function

Private Function ReadRevisionLetter(WkS As Worksheet, _
                                    RequestType As TestingRequestTypeEnum, _
                                    ChemCell As String, _
                                    WaterCell As String, _
                                    WaferCell As String) As String
    Dim CellAddr As String
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical: CellAddr = ChemCell
        Case TestingRequestTypeEnum.Water:    CellAddr = WaterCell
        Case TestingRequestTypeEnum.Wafer:    CellAddr = WaferCell
    End Select

    Dim RawText As String
    RawText = CStr(WkS.Range(CellAddr).Value)
    Dim Pos As Long
    Pos = InStr(UCase(RawText), "REV ")
    If Pos > 0 Then
        ReadRevisionLetter = Trim$(Mid$(RawText, Pos + 4, 1))
    Else
        ReadRevisionLetter = ""
    End If
End Function

' ----------------------------------------------------------------
' Private: Credit card
' ----------------------------------------------------------------
Private Function ReadCreditCardInfo(WkS As Worksheet, RequestType As TestingRequestTypeEnum) As String
    Dim Cell As Range
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical: Set Cell = WkS.Range("I11")
        Case TestingRequestTypeEnum.Water:    Set Cell = WkS.Range("J11")
        Case TestingRequestTypeEnum.Wafer:    Set Cell = WkS.Range("G11")
    End Select
    ReadCreditCardInfo = CStr(Cell.Value) & _
        IIf(CStr(Cell.Offset(1).Value) = "", "", " " & CStr(Cell.Offset(1).Value))
End Function

' ----------------------------------------------------------------
' Private: Service date extraction from filename
' ----------------------------------------------------------------
Private Function ExtractServiceDate(FileName As String) As Date
    Dim StartPos As Long
    Dim EndPos   As Long

    StartPos = InStr(1, FileName, "[")
    EndPos = InStr(StartPos + 1, FileName, "]")

    If Not m_Logger.DebugMode Then On Error GoTo Fallback

    If StartPos > 0 And EndPos > StartPos Then
        Dim Raw As String
        Raw = Mid$(FileName, StartPos + 1, EndPos - StartPos - 1)
        Raw = Replace(Raw, "_", " ")
        Raw = Replace(Raw, "-", "/")
        ExtractServiceDate = CDate(Raw)
        Exit Function
    End If

Fallback:
    ExtractServiceDate = DateSerial(1901, 1, 1)  ' sentinel: no date in filename
End Function

' ----------------------------------------------------------------
' Private: Emails
' ----------------------------------------------------------------
Private Sub AttachEmails(Submission As clsTRSubmission, WkS As Worksheet, RequestType As TestingRequestTypeEnum)
    Dim InvMainCell As String
    Dim InvCCCell   As String

    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical
            InvMainCell = "N10": InvCCCell = "N11"
        Case TestingRequestTypeEnum.Wafer
            InvMainCell = "J10": InvCCCell = "J11"
        Case TestingRequestTypeEnum.Water
            InvMainCell = "P10": InvCCCell = "P11"
        Case Else
            Err.Raise vbObjectError + 501, "clsSubmissionFormParser.AttachEmails", _
                "Unrecognised request type: " & RequestType
    End Select

    ' Results emails (fixed cells, all form types).
    ParseAndAddEmails Submission, CStr(WkS.Range("C15").Value), isInvoice:=False, IsCC:=False
    ParseAndAddEmails Submission, CStr(WkS.Range("C17").Value), isInvoice:=False, IsCC:=True

    ' Invoice emails - DB defaults take priority over form cells.
    If Not AddDefaultInvoiceEmails(Submission) Then
        ParseAndAddEmails Submission, CStr(WkS.Range(InvMainCell).Value), isInvoice:=True, IsCC:=False
    End If
    ParseAndAddEmails Submission, CStr(WkS.Range(InvCCCell).Value), isInvoice:=True, IsCC:=True
End Sub

Private Sub ParseAndAddEmails(Submission As clsTRSubmission, EmailString As String, _
                               isInvoice As Boolean, IsCC As Boolean)
    Dim S As String
    S = EmailString
    S = Replace(S, ";", ",")
    S = Replace(S, " ", ",")
    S = Replace(S, vbNewLine, ",")

    Dim Tokens() As String
    Tokens = Split(S, ",")

    Dim i     As Long
    Dim Addr  As String
    For i = LBound(Tokens) To UBound(Tokens)
        Addr = Trim$(Tokens(i))
        If Len(Addr) > 0 And IsValidEmailAddress(Addr) Then
            If isInvoice Then
                Submission.AddInvoiceEmail Addr, IsCC
            Else
                Submission.AddResultsEmail Addr, IsCC
            End If
        End If
    Next i
End Sub

' Basic email validation - same rule as old clsTestingRequestEmail.IsValidEmail.
Private Function IsValidEmailAddress(Addr As String) As Boolean
    IsValidEmailAddress = (InStr(Addr, "@") > 1 And InStr(Addr, ".") > 2)
End Function

' TODO: Move this query to clsSubmissionRepository when DB persistence is added.
Private Function AddDefaultInvoiceEmails(Sub_ As clsTRSubmission) As Boolean
    If Sub_.CustomerID <= 0 Then Exit Function

    Dim SQL As String
    SQL = "SELECT contact_email FROM invoice_default_contacts " & _
          "WHERE customer_id = " & Sub_.CustomerID & _
          "  AND is_default = TRUE AND is_active = TRUE"

    Dim Results As Variant
    Results = m_AccessDB.ExecuteQuery(SQL)

    If UBound(Results) < 0 Then Exit Function
    If UBound(Results, 2) < LBound(Results, 2) Then Exit Function

    Dim i As Long
    For i = LBound(Results, 2) To UBound(Results, 2)
        Dim Addr As String
        Addr = Trim$(CStr(Results(0, i)))
        If Len(Addr) > 0 Then Sub_.AddInvoiceEmail Addr, IsCC:=False
    Next i

    AddDefaultInvoiceEmails = True
End Function

' ----------------------------------------------------------------
' Private: Sample range and sample building
' ----------------------------------------------------------------
Private Sub AttachSamples(Sub_ As clsTRSubmission, WkS As Worksheet, RequestType As TestingRequestTypeEnum)
    Dim SampleRange As Range
    Set SampleRange = GetSampleRange(WkS)
    If SampleRange Is Nothing Then Exit Sub

    Dim SampleCell As Range
    Dim S          As clsTRSample

    For Each SampleCell In SampleRange.Columns(1).Cells
        If Not IsEmpty(SampleCell) And Len(Trim$(CStr(SampleCell.Value))) > 0 Then
            Set S = m_SampleParser.BuildFromRow(SampleCell, SampleRange, RequestType)
            Sub_.AddSample S
        End If
    Next SampleCell
End Sub

Private Function GetSampleRange(WkS As Worksheet) As Range
    Dim FirstRow  As Long
    Dim LastRow   As Long
    Dim LastCol   As Long
    Const FOOTER_ROWS As Long = 11

    If WkS.Range("B20").Value = "Sample ID" Then
        FirstRow = 22
    Else
        FirstRow = 21
    End If

    LastRow = WkS.Cells(WkS.Rows.Count, 2).End(xlUp).Row - FOOTER_ROWS
    LastCol = WkS.Cells(FirstRow - 2, WkS.Columns.Count).End(xlToLeft).Column - 1

    If LastRow < FirstRow Or LastCol < 2 Then Exit Function  ' empty form

    Set GetSampleRange = WkS.Range(WkS.Cells(FirstRow, 2), WkS.Cells(LastRow, LastCol))
End Function

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsSubmissionFormParser." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Optional Context As String = "")
    m_Logger.LogError "clsSubmissionFormParser." & Src, Err.Number, Err.Description & " | " & Context
End Sub
