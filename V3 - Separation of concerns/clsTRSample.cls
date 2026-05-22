VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTRSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsSample
' Author: JT
' Created: 2026-04-10
' Description: Immutable-after-load entity for one sample row in a
'              testing request submission. Replaces clsTestingRequestSample.
' DependsOn: clsLoggingSystem, modEnums, modUtilities
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release
'
' ================================================================
' LAYER: Entity - no DB, no Worksheet, no service dependencies.
'
' LIFECYCLE:
'   1.  New clsSample
'   2.  .Initialize(Logger)
'   3.  .LoadFromDict(D)          <- scalar fields only
'   4.  .AddAnalysis(A) x N       <- in-order, parser deduplicates
'   5.  .AddElement(E) x N        <- additional elements
'
' All steps 3-6 are performed by clsSampleFormParser.
'
' Dict keys for LoadFromDict:
'   sample_id (Long), sample_name, form_chemical_name,
'   processing_time_id (Long = enum), additional_notes, requested_time,
'   chemical_id (long)
' ================================================================

Private m_Logger            As clsLoggingSystem
Private m_ID                As Long
Private m_SampleName        As String
Private m_FormChemicalName  As String
Private m_ProcessingTime    As ProcessingTimeEnum
Private m_AdditionalNotes   As String
Private m_RequestedTime     As String

Private m_ChemicalID        As Long
Private m_AnalysisIDs       As Collection           ' long, insertion-ordered
Private m_AnalysesIDDict      As Scripting.Dictionary ' long -> True, for dedup
Private m_AdditionalElementIDs As Collection          ' long
Private m_ElementIDDict     As Scripting.Dictionary ' long -> True, for dedup

' ----------------------------------------------------------------
' Initialization
' ----------------------------------------------------------------
Public Sub Initialize(Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init

    Set m_Logger = Logger
    Set m_AnalysisIDs = New Collection
    Set m_AnalysesIDDict = New Scripting.Dictionary
    Set m_AdditionalElementIDs = New Collection
    Set m_ElementIDDict = New Scripting.Dictionary
    
End Sub

' ----------------------------------------------------------------
' Scalar hydration  (called by parser after Initialize)
' ----------------------------------------------------------------
Public Sub LoadFromDict(D As Scripting.Dictionary)
    AssertInit "LoadFromDict"
    Debug.Assert Not D Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    RequireKey D, "sample_id"
    RequireKey D, "sample_name"
    RequireKey D, "form_chemical_name"
    RequireKey D, "processing_time_id"
    RequireKey D, "additional_notes"
    RequireKey D, "requested_time"
    RequireKey D, "chemical_id"

    m_ID = CLng(D("sample_id"))
    m_SampleName = CStr(D("sample_name"))
    m_FormChemicalName = CStr(D("form_chemical_name"))
    m_ProcessingTime = CInt(D("processing_time_id"))     ' stored as enum Long
    m_AdditionalNotes = CStr(D("additional_notes"))
    m_RequestedTime = CStr(D("requested_time"))
    m_ChemicalID = CLng(D("chemical_id"))

    Exit Sub
EH:                                                         HandleError "LoadFromDict", "Keys present: " & Join(D.Keys, ", ")
End Sub

' ----------------------------------------------------------------
' Child object population  (called by parser after LoadFromDict)
' ----------------------------------------------------------------
' Silently deduplicates by analysis ID - parser need not guard.
Public Sub AddAnalysis(AnalysisID As Long)
    AssertInit "AddAnalysis"
    Debug.Assert AnalysisID > 0

    If Not m_AnalysesIDDict.Exists(AnalysisID) Then
        m_AnalysesIDDict.Add AnalysisID, True
        m_AnalysisIDs.Add AnalysisID
    End If
End Sub

Public Sub AddElement(ElementID As Long)
    AssertInit "AddElement"
    Debug.Assert ElementID > 0
    
    If Not m_ElementIDDict.Exists(ElementID) Then
        m_ElementIDDict.Add ElementID, True
        m_AdditionalElementIDs.Add ElementID
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

    If NewID <= 0 Then Err.Raise vbObjectError + 400, "clsSample.SetInternalID", "ID must be positive, got: " & NewID
    If m_ID <> 0 Then Err.Raise vbObjectError + 401, "clsSample.SetInternalID", "ID already set to: " & m_ID

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

Public Property Get SampleName() As String
    SampleName = m_SampleName
End Property

Public Property Get FormChemicalName() As String
    FormChemicalName = m_FormChemicalName
End Property

Public Property Get ProcessingTime() As ProcessingTimeEnum
    ProcessingTime = m_ProcessingTime
End Property

Public Property Get ProcessingTimeStr() As String
    ProcessingTimeStr = modUtilities.ProcessingTimeToString(m_ProcessingTime)
End Property

Public Property Get AdditionalNotes() As String
    AdditionalNotes = m_AdditionalNotes
End Property

Public Property Get RequestedTime() As String
    RequestedTime = m_RequestedTime
End Property

Public Property Get ChemicalID() As Long
    ChemicalID = m_ChemicalID
End Property

Public Property Get AnalysisIDs() As Collection
    Set AnalysisIDs = modUtilities.CloneCollection(m_AnalysisIDs)
End Property

Public Property Get AdditionalElementIDs() As Collection
    Set AdditionalElementIDs = modUtilities.CloneCollection(m_AdditionalElementIDs)
End Property

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub RequireKey(D As Scripting.Dictionary, Key As String)
    Debug.Assert D.Exists(Key)
    If Not D.Exists(Key) Then
        Err.Raise vbObjectError + 200, "clsTRSample.LoadFromDict", "Missing required key: " & Key
    End If
End Sub

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsTRSample." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Optional Context As String = "")
    m_Logger.LogError "clsTRSample." & Src, Err.Number, Err.Description & " | " & Context
End Sub
