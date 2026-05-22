VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceSubmissionManager1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceSubmissionManager
' Author: JT
' Created: 2026-01-22
' Description: Manager class to coordinate collecting submission properties, ready to process towards invoicing
' DependsOn: clsLoggingSystem, clsInvoiceReportGatekeeper, clsSubmissionService, clsTRSubmissionService, clsTRSubmission,
'            clsCustomerService, clsChemicalService, clsAnalysisService, clsElementService,
' ChangeLog:
'   - 1.0.0 - 2026-01-22 - Initial release
'   - 2.0.0 - 2026-03-04 - Expanded to replace clsInvoiceSubmissionManagerBatch and clsInvoiceSubmissionBase
'   - 3.0.0 - 2026-04-14 - Replumbed to use clsSubmissionService / clsTRSubmission

' Responsibilities:
' - Filters files by invoice eligibility
' - Extracts file dates for batch processing

Private m_Logger        As clsLoggingSystem
Private m_Gatekeeper    As clsInvoiceReportGatekeeper
Private m_SubmissionSvc As clsTRSubmissionService
Private m_Submissions   As Collection
Private m_FilesScanned As Long
Private m_FilesEligible As Long
Private m_SubmissionsCreated As Long
Private m_FilesSkipped As Long
Private m_Errors As Long

' ==============
' Get Properties
' ==============
Public Property Get Submissions() As Collection
    If m_Submissions Is Nothing Then
        Set m_Submissions = New Collection
    End If
    Set Submissions = m_Submissions
End Property

' ==============
' Initialization
' ==============
Public Sub Initialize(Logger As clsLoggingSystem)

    Debug.Assert Not Logger Is Nothing
    Debug.Assert Not AccessDB Is Nothing
    Debug.Assert m_Logger Is Nothing

    If Not Logger.DebugMode Then On Error GoTo ErrorHandler

    If Not m_Logger Is Nothing Then Err.Raise vbObjectError + 100, , "The manager is already initialized."

    Set m_Logger = Logger
    Set m_Gatekeeper = New clsInvoiceReportGatekeeper

    ' Build the SLIMDict used by analysis and element services.
    Dim SLIMDict As New clsSLIMDictionaryManager
    SLIMDict.Initialize Logger, AccessDB

    ' Build the four domain services required by clsSubmissionService.
    Dim CustomerSvc As New clsCustomerService
    CustomerSvc.Initialize AccessDB, Logger

    Dim ChemicalSvc As New clsChemicalService
    ChemicalSvc.Initialize AccessDB, Logger

    Dim AnalysisSvc As New clsAnalysisService
    AnalysisSvc.Initialize SLIMDict, Logger

    Dim ElementSvc As New clsElementService
    ElementSvc.Initialize SLIMDict, Logger

    Set m_SubmissionSvc = New clsTRSubmissionService
    m_SubmissionSvc.Initialize AccessDB, CustomerSvc, ChemicalSvc, AnalysisSvc, ElementSvc, Logger

    Exit Sub
ErrorHandler:
    Call HandleError("Initialize", "Error initializing the manager.")
End Sub

' ====
' APIs
' ====
Public Sub LoadSingle(FilePath As String)
    Dim Submission As clsTRSubmission
    Dim ErrorMessage As String

    Debug.Assert FilePath <> ""
    Debug.Assert m_Submissions Is Nothing

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    Call AssertInitialized("LoadSingle")

    If Not m_Gatekeeper.ValidateBeforeInvoice(FilePath, ErrorMessage) Then Err.Raise vbObjectError + 1, , ErrorMessage

    Set m_Submissions = New Collection

    Set Submission = BuildSubmission(FilePath)

    If Not Submission Is Nothing Then
        m_Submissions.Add Submission
    End If


    Exit Sub
ErrorHandler:
    Call HandleError("LoadSingle")

End Sub

Public Sub LoadByDateRange(FolderPath As String, StartDate As Date, EndDate As Date)
    Dim FileName As String

    Debug.Assert FolderPath <> ""
    Debug.Assert m_Submissions Is Nothing
    Debug.Assert EndDate > StartDate

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    Call AssertInitialized("LoadByDateRange")

    If StartDate > EndDate Then
        Err.Raise vbObjectError + 10, "clsInvoiceSubmissionManager.LoadByDateRange", "StartDate cannot be greater than EndDate."
    End If

    Set m_Submissions = New Collection

    m_FilesScanned = 0
    m_FilesEligible = 0
    m_SubmissionsCreated = 0
    m_FilesSkipped = 0
    m_Errors = 0

    FolderPath = NormalizeFolderPath(FolderPath)

    FileName = Dir(FolderPath & "*.xlsx")

    Call m_Logger.LogMessage("clsInvoiceSubmissionManager.LoadByDateRange", LogLevelEnum.LogInfo, "Scanning folder: " & FolderPath)

    'Loops through each file in the specified folder
    Do While FileName <> ""

        Call ProcessFile(FolderPath, FileName, StartDate, EndDate)

        FileName = Dir()
    Loop

    Call LogBatchSummary

    Exit Sub
ErrorHandler:
    Call HandleError("LoadByDateRange", Err.Description)

End Sub

' ===============
' Private Helpers
' ===============
Private Function NormalizeFolderPath(FolderPath As String) As String

    ' Normalize FilePath
    If Len(FolderPath) = 0 Then
        Err.Raise vbObjectError + 2, "NormalizeFolderPath", "FolderPath cannot be empty."
    End If

    If Right$(FolderPath, 1) <> "\" Then
        FolderPath = FolderPath & "\"
    End If

    NormalizeFolderPath = FolderPath

End Function
Private Sub ProcessFile(FolderPath As String, FileName As String, StartDate As Date, EndDate As Date)

    Dim FilePath As String
    Dim ErrorMessage As String
    Dim FileDate As Date
    Dim Submission As clsTRSubmission

    Call m_Logger.LogMessage("clsInvoiceSubmissionManager.ProcessFile", LogLevelEnum.LogDebug, "Evaluating file: " & FileName)

    m_FilesScanned = m_FilesScanned + 1

    FilePath = FolderPath & FileName

    If Not IsValidInvoiceFile(FilePath) Then
        m_FilesSkipped = m_FilesSkipped + 1
        Exit Sub
    End If
    If Not TryExtractFileDate(FileName, FileDate) Then
        m_FilesSkipped = m_FilesSkipped + 1
        Exit Sub
    End If
    If Not IsWithinRange(FileDate, StartDate, EndDate) Then
        m_FilesSkipped = m_FilesSkipped + 1
        Exit Sub
    End If

    m_FilesEligible = m_FilesEligible + 1
    Call m_Logger.LogMessage("clsInvoiceSubmissionManager.ProcessFile", LogLevelEnum.LogDebug, "File passed eligibility checks: " & FileName)


    Set Submission = BuildSubmission(FilePath)

    If Not Submission Is Nothing Then
        m_SubmissionsCreated = m_SubmissionsCreated + 1
        m_Submissions.Add Submission
        Call m_Logger.LogMessage("clsInvoiceSubmissionManager.ProcessFile", LogLevelEnum.LogInfo, "Submission added: " & FileName)
    Else
        Call m_Logger.LogMessage("clsInvoiceSubmissionManager.ProcessFile", LogLevelEnum.LogWarning, "Submission could not be created: " & FileName)
    End If

End Sub
Private Function IsValidInvoiceFile(FilePath As String) As Boolean
    Dim ErrorMessage As String

    If Not m_Gatekeeper.ValidateBeforeInvoice(FilePath, ErrorMessage) Then
        Call m_Logger.LogMessage("clsInvoiceSubmissionManager.ProcessFile", LogLevelEnum.LogWarning, "Invalid file skipped: " & FilePath & " | " & ErrorMessage)
        Exit Function
    End If

    IsValidInvoiceFile = True

End Function
Private Function TryExtractFileDate(FileName As String, ByRef FileDate As Date) As Boolean
    Dim MonthPart As Integer
    Dim DayPart As Integer
    Dim YearPart As Integer

    If Len(Trim(FileName)) < 6 Then Exit Function

    If Not IsNumeric(Left$(FileName, 6)) Then Exit Function

    MonthPart = CInt(Left$(FileName, 2))
    DayPart = CInt(Mid$(FileName, 3, 2))
    YearPart = CInt(Mid$(FileName, 5, 2)) + 2000

    If MonthPart < 1 Or MonthPart > 12 Then Exit Function
    If DayPart < 1 Or DayPart > 31 Then Exit Function

    FileDate = DateSerial(YearPart, MonthPart, DayPart)

    TryExtractFileDate = True

End Function
Private Function IsWithinRange(FileDate As Date, StartDate As Date, EndDate As Date) As Boolean
    IsWithinRange = FileDate >= StartDate And FileDate <= EndDate
End Function
Private Function BuildSubmission(FilePath As String) As clsTRSubmission
    Dim TestingRequest As Worksheet
    Dim Submission As clsTRSubmission
    Dim WBk As Workbook

    Call m_Logger.LogMessage("clsInvoiceSubmissionManager.BuildSubmission", LogLevelEnum.LogDebug, "Opening workbook: " & FilePath)

    On Error GoTo ErrorHandler
    Set WBk = Workbooks.Open(FilePath, ReadOnly:=True)

    Set TestingRequest = FindTestingRequest(WBk)

    If Not TestingRequest Is Nothing Then
        Set Submission = m_SubmissionSvc.BuildFromWorksheet(TestingRequest)
        Set BuildSubmission = Submission
    End If

CleanUp:
    If Not WBk Is Nothing Then
        WBk.Close SaveChanges:=False
    End If

    Exit Function
ErrorHandler:
    m_Errors = m_Errors + 1
    Call m_Logger.LogError("clsInvoiceSubmissionManager.BuildSubmission", Err.Number, Err.Description)
    Err.Clear
    Resume CleanUp
End Function

Private Function FindTestingRequest(WBk As Workbook) As Worksheet
    Dim WkS As Worksheet

    Call m_Logger.LogMessage("clsInvoiceSubmissionManager.FindTestingRequest", LogLevelEnum.LogDebug, "Finding testing request for: " & WBk.Name)

    For Each WkS In WBk.Sheets
        If CStr(WkS.Range("D3").Value) Like "*Testing Request Form" Then
            Set FindTestingRequest = WkS
            Call m_Logger.LogMessage("clsInvoiceSubmissionManager.FindTestingRequest", LogLevelEnum.LogInfo, "Testing request found in workbook: " & WBk.Name)
            Exit Function
        End If
    Next WkS

    ' Log failure
    Call m_Logger.LogMessage("clsInvoiceSubmissionManager.FindTestingRequest", LogLevelEnum.LogWarning, "No testing request found. Closing workbook.")


End Function
Private Sub LogBatchSummary()

    Call m_Logger.LogMessage("clsInvoiceSubmissionManager.LoadByDateRange", LogLevelEnum.LogInfo, "Batch summary | Files scanned: " & m_FilesScanned & _
        " | Eligible: " & m_FilesEligible & " | Submissions created: " & m_SubmissionsCreated & " | Skipped: " & m_FilesSkipped & " | Errors: " & m_Errors)

End Sub
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsInvoiceSubmissionManager." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsInvoiceSubmissionManager." & MethodName, "The manager must be initialized before use."
    End If
End Sub

