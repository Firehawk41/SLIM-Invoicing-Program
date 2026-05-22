VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnalysisRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsAnalysisRepository
' Author: JT
' Created: 2026-04-10
' Description: Owns ALL SQL for the analyses domain.
'              No business logic. No object graph traversal.
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsAnalysis
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release
'
' ================================================================
' LAYER: Repository - translates between objects and SQL.
'
' ASSUMED clsAccessDatabase interface:
'   .ExecuteQuery(SQL)   As Variant  SELECT -> 2D array(field,row)
'                                    or Empty if no rows
'   .ExecuteNonQuery(SQL)            INSERT / UPDATE / DELETE
'   .Executenonquery(SQL)  As Long     INSERT, returns autonumber ID
'   .BeginTransaction
'   .CommitTransaction
'   .RollbackTransaction
'
' SelectAll returns fields in this column order (0-based):
'   0 ID  |  1 analysis_name  |  2 analysis_description
' ================================================================

Private m_Logger   As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase

' ----------------------------------------------------------------
' Initialization
' ----------------------------------------------------------------
Public Sub Initialize(AccessDB As clsAccessDatabase, Logger As clsLoggingSystem)
    Debug.Assert Not AccessDB Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init

    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
End Sub

' ----------------------------------------------------------------
' Bulk read  (called ONLY by clsAnalysisCache.Build)
' ----------------------------------------------------------------
Public Function SelectAll() As Variant
    AssertInit "SelectAll"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAll = m_AccessDB.ExecuteQuery( _
        "SELECT ID, analysis_name, analysis_description " & _
        "FROM   analyses " & _
        "ORDER BY analysis_name")
    Exit Function
EH:                                                         HandleError "SelectAll", "Failed"
End Function
' ----------------------------------------------------------------
' Form analysis lookup  (called ONLY by clsAnalysisCache)
' Returns 2D array: col 0 = form_name, col 1 = analysis_id
' Returns Empty if table is empty
' ----------------------------------------------------------------
Public Function SelectAllFormAnalyses() As Variant
    AssertInit "SelectAllFormAnalyses"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllFormAnalyses = m_AccessDB.ExecuteQuery( _
        "SELECT fa.form_analysis_name, fam.database_analysis_id " & _
        "FROM   form_analyses fa " & _
        "INNER JOIN form_analysis_mapping fam " & _
        "    ON fa.ID = fam.form_analysis_id " & _
        "ORDER BY fa.form_analysis_name")
    Exit Function
EH:                                         HandleError "SelectAllFormAnalyses", "Failed"
End Function
' ----------------------------------------------------------------
' Writes
' ----------------------------------------------------------------
' Inserts a new analysis row. Returns the new autonumber ID.
' The Service calls A.SetInternalID(result) after this returns.
'
Public Function Insert(A As clsAnalysis) As Long
    AssertInit "Insert"
    Debug.Assert Not A Is Nothing
    Debug.Assert A.ID = 0               ' must be unsaved
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim SQL As String
    SQL = "INSERT INTO analyses (analysis_name, analysis_description) " & _
          "VALUES (" & SqlStr(A.Name) & ", " & SqlStr(A.Description) & ")"

    Insert = m_AccessDB.ExecuteNonQuery(SQL)
    Exit Function
EH:                                                         HandleError "Insert", "Failed on: " & A.Name
End Function

Public Sub Update(A As clsAnalysis)
    AssertInit "Update"
    Debug.Assert Not A Is Nothing
    Debug.Assert A.ID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim SQL As String
    SQL = "UPDATE analyses SET " & _
          "    analysis_name        = " & SqlStr(A.Name) & ", " & _
          "    analysis_description = " & SqlStr(A.Description) & " " & _
          "WHERE ID = " & A.ID

    m_AccessDB.ExecuteNonQuery SQL
    Exit Sub
EH:                                                         HandleError "Update", "Failed on ID: " & A.ID
End Sub

Public Sub Delete(AnalysisID As Long)
    AssertInit "Delete"
    Debug.Assert AnalysisID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM analyses WHERE ID = " & AnalysisID
    Exit Sub
EH:                                                         HandleError "Delete", "Failed on ID: " & AnalysisID
End Sub

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Function SqlStr(V As String) As String
    SqlStr = "'" & Replace(V, "'", "''") & "'"
End Function

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsAnalysisRepository." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Context As String)
    m_Logger.LogError "clsAnalysisRepository." & Src, Err.Number, Err.Description & " | " & Context
End Sub


