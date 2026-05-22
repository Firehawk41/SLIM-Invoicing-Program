Attribute VB_Name = "modInvoiceSystem"
Option Explicit
' ================================================================
' Name: modInvoiceSystem
' Author: JT
' Created: 2022-09-08
' Description: Composition root and macro entry point for the
'              invoice generation pipeline. Builds all services
'              once and injects them into consumers.
' DependsOn: clsLoggingSystem, clsAccessDatabase, modConfig,
'            modEnums, frmInvoiceTypeSelector, clsStalenessChecker,
'            clsCustomerService, clsChemicalService,
'            clsAnalysisService, clsElementService,
'            clsTRFormInputResolver, clsInvoiceSubmissionManager,
'            clsQuoteRepository, clsQuoteCache, clsQuoteService,
'            clsSalesOrderPricingEngine, clsSalesOrderLineItemBuilder,
'            clsSalesOrderBuilder, clsSalesOrder,
'            clsSalesOrderWriterCSV
' ChangeLog:
'   - 5.0.1 - 2025-06-05 - Added metadata
'   - 5.1.0 - 2025-11-07 - Added CSV/Word output toggle
'   - 5.2.0 - 2025-11-18 - Integrated output choice with userform
'   - 6.0.0 - 2026-03-04 - Removed Word support
'   - 7.0.0 - 2026-05-08 - V3 port: removed legacy managers,
'                           wired QuoteService directly into
'                           PricingEngine, moved sensitive paths
'                           to modConfig, collapsed scheduler
'                           entry point into single pipeline,
'                           refactored construction into helpers
'   - 7.0.1 - 2026-05-19 - Fixed BuildSOBuilder: QuoteCache
'                           Initialize missing ChemicalSvc.
'                           QuoteService Initialize updated to
'                           inject Repo and Cache separately.
'                           QuoteCache.Build called before
'                           QuoteService construction.
'
' ================================================================
' LAYER: Entry point and composition root. All services are built
'        once here and injected into consumers. No business logic
'        lives in this module.
'
' PIPELINE:
'   CreateInvoice
'     1. Initialise logger and staleness check
'     2. Show frmInvoiceTypeSelector, collect user input
'     3. Build infrastructure and domain services
'     4. Load submissions (Individual or Batch)
'     5. Build sales orders via pricing engine and builder
'     6. Write CSV via clsSalesOrderWriterCSV
'
' HELPERS:
'   BuildDomainServices   — wires AccessDB into all four domain
'                           service stacks
'   BuildSubmissionManager — wires domain services and resolver
'                            into submission manager
'   BuildSOBuilder        — wires quote service and pricing engine
'                           into sales order builder
'
' CONSTANTS:
'   DEFAULT_PAYMENT_TERMS — interim fallback pending
'                           clsCustomerBillingProfile implementation.
'   DEBUG_MODE            — set in modConfig. True for development,
'                           False for production.
' ================================================================

Private Const DEFAULT_PAYMENT_TERMS As Long = 30

' ============
' Entry Point
' ============

Public Sub CreateInvoice()

    Dim Logger As clsLoggingSystem
    Dim Checker As clsStalenessChecker
    Dim InvoiceUserForm As frmInvoiceTypeSelector
    Dim AccessDB As clsAccessDatabase
    Dim ChemicalSvc As clsChemicalService
    Dim AnalysisSvc As clsAnalysisService
    Dim CustomerSvc As clsCustomerService
    Dim SubmissionManager As clsInvoiceSubmissionManager
    Dim SOBuilder As clsSalesOrderBuilder
    Dim Writer As clsSalesOrderWriterCSV
    Dim Submissions As Collection
    Dim Submission As clsTRSubmission
    Dim SO As clsSalesOrder
    Dim StartDate As Date
    Dim EndDate As Date
    Dim FilePath As String
    Dim InvoiceSelection As InvoiceSelectionEnum
    Dim StartTime As Single
    Dim PipelineRan As Boolean
    Dim Results As New Collection

    PipelineRan = False

    ' --- Logger ---
    Set Logger = New clsLoggingSystem
    Call Logger.Initialize("CreateInvoice", modConfig.DEBUG_MODE)

    If Not Logger.DebugMode Then
        On Error GoTo ErrorHandler
    End If

    ' --- Staleness check ---
    Set Checker = New clsStalenessChecker
    Call Checker.Initialize( _
        ThisWorkbook.Name, _
        ThisWorkbook.FullName, _
        FileDateTime(ThisWorkbook.FullName), _
        Logger)
    If Not Checker.IsCurrent Then
        Err.Raise 1984, "modInvoiceSystem.CreateInvoice", Checker.IsObsoleteMessage
    End If

    ' --- User input ---
    Set InvoiceUserForm = New frmInvoiceTypeSelector
    Call InvoiceUserForm.Initialize(Logger)
    InvoiceUserForm.Show

    If InvoiceUserForm.Cancelled Then GoTo CleanUp

    StartDate = InvoiceUserForm.StartDate
    EndDate = InvoiceUserForm.EndDate
    FilePath = InvoiceUserForm.SelectedFilePath
    InvoiceSelection = InvoiceUserForm.InvoiceSelection
    Set InvoiceUserForm = Nothing

    ' --- Application settings ---
    If Not Logger.DebugMode Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    End If

    StartTime = Timer
    PipelineRan = True

    ' --- Infrastructure ---
    Set AccessDB = New clsAccessDatabase
    Call AccessDB.Initialize(modConfig.DB_PATH, Logger)

    ' --- Shared domain services ---
    Set AnalysisSvc = New clsAnalysisService
    Call AnalysisSvc.Initialize(AccessDB, Logger)
    
    Set ChemicalSvc = New clsChemicalService
    Call ChemicalSvc.Initialize(AccessDB, Logger)
    
    Set CustomerSvc = New clsCustomerService
    Call CustomerSvc.Initialize(AccessDB, Logger)

    ' --- Submission manager ---
    Set SubmissionManager = BuildSubmissionManager(AccessDB, AnalysisSvc, ChemicalSvc, CustomerSvc, Logger)

    Select Case InvoiceSelection
        Case InvoiceSelectionEnum.Individual
            Call SubmissionManager.LoadSingle(FilePath)
        Case InvoiceSelectionEnum.Batch
            Call SubmissionManager.LoadByDateRange(FilePath, StartDate, EndDate)
        Case Else
            Err.Raise 911, "modInvoiceSystem.CreateInvoice", "Unknown InvoiceSelectionEnum value."
    End Select

    Set Submissions = SubmissionManager.Submissions

    If Submissions Is Nothing Then
        Err.Raise 1001, "modInvoiceSystem.CreateInvoice", "No submissions loaded."
    End If
    If Submissions.Count = 0 Then
        Err.Raise 1001, "modInvoiceSystem.CreateInvoice", "No submissions found for the selected criteria."
    End If

    ' --- SO builder ---
    Set SOBuilder = BuildSOBuilder(AccessDB, AnalysisSvc, ChemicalSvc, CustomerSvc, Logger)

    ' --- Pipeline loop ---
    For Each Submission In Submissions
        Set SO = SOBuilder.BuildFromSubmission(Submission)
        If Not SO Is Nothing Then
            Results.Add SO
        Else
            Call Logger.LogMessage( _
                "modInvoiceSystem.CreateInvoice", _
                LogLevelEnum.LogWarning, _
                "BuildFromSubmission returned Nothing for submission: " & Submission.FileName)
        End If
    Next Submission
        
    ' --- Writer ---
    Set Writer = New clsSalesOrderWriterCSV
    Call Writer.Initialize(Logger)
    Call Writer.BeginOutput
    For Each SO In Results
        Call Writer.WriteInvoice(SO)
    Next SO

    Call Writer.SaveInvoice( _
        modConfig.OUTPUT_PATH & "Invoice_" & Format(Now, "yyyymmdd_hhmmss") & ".csv")

CleanUp:
    If PipelineRan Then
        Call Logger.LogMessage( _
            "modInvoiceSystem.CreateInvoice", _
            LogLevelEnum.LogInfo, _
            "Execution time: " & Round(Timer - StartTime, 1) & " seconds.")
    End If

    If Not Writer Is Nothing Then Call Writer.CloseInvoice

    Set Writer = Nothing
    Set SOBuilder = Nothing
    Set SubmissionManager = Nothing
    Set ChemicalSvc = Nothing
    Set AccessDB = Nothing
    Set InvoiceUserForm = Nothing

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    If Not Logger Is Nothing Then
        Call Logger.LogMessage( _
            "modInvoiceSystem.CreateInvoice", _
            LogLevelEnum.LogInfo, "Pipeline complete.")
        Logger.CloseLogFile
        Set Logger = Nothing
    End If

    Exit Sub
ErrorHandler:
    Call Logger.LogError("modInvoiceSystem.CreateInvoice", Err.Number, Err.Description, False)
    GoTo CleanUp
End Sub

' ================
' Private Helpers
' ================

Private Function BuildSubmissionManager( _
    AccessDB As clsAccessDatabase, _
    AnalysisSvc As clsAnalysisService, _
    ChemicalSvc As clsChemicalService, _
    CustomerSvc As clsCustomerService, _
    Logger As clsLoggingSystem) As clsInvoiceSubmissionManager

    Dim ElementSvc As clsElementService
    Dim Resolver As clsTRFormInputResolver
    Dim SubmissionManager As clsInvoiceSubmissionManager

    Set ElementSvc = New clsElementService
    Call ElementSvc.Initialize(AccessDB, Logger)

    Set Resolver = New clsTRFormInputResolver
    Call Resolver.Initialize(CustomerSvc, ChemicalSvc, ElementSvc, Logger)

    Set SubmissionManager = New clsInvoiceSubmissionManager
    Call SubmissionManager.Initialize( _
        AccessDB, CustomerSvc, ChemicalSvc, _
        AnalysisSvc, ElementSvc, Resolver, Logger)

    Set BuildSubmissionManager = SubmissionManager
End Function

Private Function BuildSOBuilder( _
    AccessDB As clsAccessDatabase, _
    AnalysisSvc As clsAnalysisService, _
    ChemicalSvc As clsChemicalService, _
    CustomerSvc As clsCustomerService, _
    Logger As clsLoggingSystem) As clsSalesOrderBuilder

    Dim QuoteRepo As clsQuoteRepository
    Dim QuoteCache As clsQuoteCache
    Dim QuoteSvc As clsQuoteService
    Dim PricingEngine As clsSalesOrderPricingEngine
    Dim LineItemBuilder As clsSalesOrderLineItemBuilder
    Dim SOBuilder As clsSalesOrderBuilder
    Dim SalesOrderSvc As clsSalesOrderService

    Set QuoteRepo = New clsQuoteRepository
    Call QuoteRepo.Initialize(AccessDB, Logger)

    Set QuoteCache = New clsQuoteCache
    Call QuoteCache.Initialize(QuoteRepo, ChemicalSvc, Logger)
    Call QuoteCache.Build

    Set QuoteSvc = New clsQuoteService
    Call QuoteSvc.Initialize(QuoteRepo, QuoteCache, Logger)

    Set PricingEngine = New clsSalesOrderPricingEngine
    Call PricingEngine.Initialize(QuoteSvc, DEFAULT_PAYMENT_TERMS, Logger)

    Set LineItemBuilder = New clsSalesOrderLineItemBuilder
    Call LineItemBuilder.Initialize(PricingEngine, AnalysisSvc, Logger)

    Set SalesOrderSvc = New clsSalesOrderService
    Call SalesOrderSvc.Initialize(AccessDB, Logger)

    Set SOBuilder = New clsSalesOrderBuilder
    Call SOBuilder.Initialize(LineItemBuilder, CustomerSvc, SalesOrderSvc, Logger)

    Set BuildSOBuilder = SOBuilder
End Function

