VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTRSubmissionService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsTRSubmissionService
' Author: JT
' Created: 2026-04-10
' Description: Orchestrates the submission domain.
'              THE only public API callers should interact with.
'              Replaces clsTRSubmissionBuilder.
' DependsOn: clsLoggingSystem, clsAccessDatabase,
'            clsCustomerService, clsChemicalService,
'            clsAnalysisService, clsElementService,
'            clsTRSubmissionFormParser, clsTRSubmission, clsTRFormInputResolver
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release (replaces clsTRSubmissionBuilder v0.1.0)
'
' ================================================================
' LAYER: Service - owns the parser; callers never touch it.
'
' CURRENT SCOPE:
'   Form parsing only. No cache or repository yet.
'   When DB persistence is added:
'     1. Add clsSubmissionRepository and clsSampleRepository
'     2. Add SaveSubmission / LoadSubmission / DeleteSubmission
'     3. Move AddDefaultInvoiceEmails SQL from parser to repo
' ================================================================

Private m_Logger     As clsLoggingSystem
Private m_FormParser As clsTRSubmissionFormParser

' ----------------------------------------------------------------
' Initialization  -  builds FormParser internally
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
    Debug.Assert Not InputResolver Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing         ' enforce single-init

    If Not Logger.DebugMode Then On Error GoTo EH
    Set m_Logger = Logger

    Set m_FormParser = New clsTRSubmissionFormParser
    m_FormParser.Initialize AccessDB, CustomerSvc, ChemicalSvc, AnalysisSvc, ElementSvc, InputResolver, Logger

    Exit Sub
EH:                                                         HandleError "Initialize", "Failed"
End Sub

' ----------------------------------------------------------------
' Public API
' ----------------------------------------------------------------
' Builds and returns a fully-hydrated clsSubmission from a worksheet.
' This is the primary entry point for all form processing.
'
Public Function BuildFromWorksheet(WkS As Worksheet) As clsTRSubmission
    Dim S As clsTRSubmission
    
    AssertInit "BuildFromWorksheet"
    Debug.Assert Not WkS Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    
    Set S = m_FormParser.BuildFromWorksheet(WkS)

    m_Logger.LogMessage "clsSubmissionService.BuildFromWorksheet", LogLevelEnum.LogInfo, "Built submission from: " & WkS.Parent.Name & " | Samples: " & S.Samples.Count

    Set BuildFromWorksheet = S
    Exit Function
EH:
    HandleError "BuildFromWorksheet", "File: " & WkS.Parent.Name
End Function

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsSubmissionService." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Optional Context As String = "")
    m_Logger.LogError "clsSubmissionService." & Src, Err.Number, Err.Description & " | " & Context
End Sub
