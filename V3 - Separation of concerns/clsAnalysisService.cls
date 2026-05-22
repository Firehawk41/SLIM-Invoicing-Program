VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnalysisService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsAnalysisService
' Author: JT
' Created: 2026-04-10
' Description: Orchestrates the analyses domain.
'              THE only public API callers should interact with.
' DependsOn: clsLoggingSystem, clsAccessDatabase,
'            clsAnalysisRepository, clsAnalysisCache, clsAnalysis
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release
'
' ================================================================
' LAYER: Service - owns Repo and Cache; callers never touch them.
'
' READ PATTERN:
'   All reads go through the Cache - no DB round-trips.
'
' WRITE PATTERN:
'   Validate -> Repo write -> assign ID (if new) -> Cache.Upsert
'   Cache is updated incrementally; no full rebuild after writes.
'
' VALIDATION:
'   IsValid* and IsUnique* methods are public so UI controls can
'   pre-validate individual fields before a form is submitted.
'   The same checks are enforced inside CreateAnalysis.
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_Repo   As clsAnalysisRepository
Private m_Cache  As clsAnalysisCache

' ----------------------------------------------------------------
' Initialization  -  builds Repo and Cache internally
' ----------------------------------------------------------------
Public Sub Initialize(AccessDB As clsAccessDatabase, Logger As clsLoggingSystem)
    Debug.Assert Not AccessDB Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init

    If Not Logger.DebugMode Then On Error GoTo EH
    Set m_Logger = Logger

    Set m_Repo = New clsAnalysisRepository
    m_Repo.Initialize AccessDB, Logger

    Set m_Cache = New clsAnalysisCache
    m_Cache.Initialize m_Repo, Logger      ' Cache.Build() is called inside Initialize

    Exit Sub
EH:
    HandleError "Initialize", "Failed"
End Sub

' ----------------------------------------------------------------
' Reads  (all via Cache - zero DB round-trips)
' ----------------------------------------------------------------
Public Function LoadAnalysis(AnalysisID As Long) As clsAnalysis
    AssertInit "LoadAnalysis"
    Set LoadAnalysis = m_Cache.GetByID(AnalysisID)
End Function

Public Function GetByName(Name As String) As clsAnalysis
    AssertInit "GetByName"
    Set GetByName = m_Cache.GetByName(Name)
End Function

Public Function AllAnalyses() As Collection
    AssertInit "AllAnalyses"
    Set AllAnalyses = m_Cache.AllAnalyses
End Function
Public Function GetIDsByFormName(FormName As String) As Collection
    AssertInit "GetIDsByFormName"
    Set GetIDsByFormName = m_Cache.GetIDsByFormName(FormName)
End Function
' ----------------------------------------------------------------
' Writes
' ----------------------------------------------------------------
Public Function CreateAnalysis(D As Scripting.Dictionary) As clsAnalysis
    AssertInit "CreateAnalysis"
    Debug.Assert Not D Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    RequireValidName CStr(D("analysis_name"))
    RequireUniqueName CStr(D("analysis_name"))

    Dim A As New clsAnalysis
    A.Initialize m_Logger
    A.LoadFromDict BuildNewRecord(D)

    SaveAnalysis A
    Set CreateAnalysis = A
    Exit Function
EH:                                                         HandleError "CreateAnalysis", "Failed on: " & D("analysis_name")
End Function

Public Sub SaveAnalysis(A As clsAnalysis)
    AssertInit "SaveAnalysis"
    Debug.Assert Not A Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    If A.ID = 0 Then
        Dim NewID As Long
        NewID = m_Repo.Insert(A)
        A.SetInternalID NewID       ' seal the ID on the object (Friend access)
    Else
        m_Repo.Update A
    End If

    m_Cache.Upsert A                ' incremental update - no full rebuild
    Exit Sub
EH:                                                         HandleError "SaveAnalysis", "Failed on: " & A.Name
End Sub

Public Sub DeleteAnalysis(AnalysisID As Long)
    AssertInit "DeleteAnalysis"
    Debug.Assert AnalysisID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_Repo.Delete AnalysisID
    m_Cache.Remove AnalysisID
    Exit Sub
EH:                                                         HandleError "DeleteAnalysis", "Failed on ID: " & AnalysisID
End Sub

' ----------------------------------------------------------------
' Validation  (public so UI can check fields before submit)
' ----------------------------------------------------------------
Public Function IsValidName(V As String) As Boolean
    IsValidName = (Len(Trim$(V)) > 0)
End Function

Public Function IsUniqueName(V As String) As Boolean
    AssertInit "IsUniqueName"
    IsUniqueName = Not m_Cache.ExistsByName(V)
End Function

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Function BuildNewRecord(D As Scripting.Dictionary) As Scripting.Dictionary
    Dim R As New Scripting.Dictionary
    R.Add "analysis_id", 0
    R.Add "analysis_name", D("analysis_name")
    R.Add "analysis_description", IIf(D.Exists("analysis_description"), D("analysis_description"), "")
    Set BuildNewRecord = R
End Function

Private Sub RequireValidName(V As String)
    If Not IsValidName(V) Then
        Err.Raise vbObjectError + 400, , "Analysis name cannot be blank"
    End If
End Sub

Private Sub RequireUniqueName(V As String)
    If Not IsUniqueName(V) Then
        Err.Raise vbObjectError + 401, , "Analysis name already exists: " & V
    End If
End Sub

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsAnalysisService." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Optional Context As String)
    m_Logger.LogError "clsAnalysisService." & Src, Err.Number, Err.Description & " | " & Context
End Sub

