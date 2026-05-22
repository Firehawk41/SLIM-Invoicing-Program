VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsElementRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsElementRepository
' Author: JT
' Created: 2026-04-10
' Description: Owns ALL SQL for the elements domain.
'              No business logic. No object graph traversal.
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsElement
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release
'
' ================================================================
' LAYER: Repository — translates between objects and SQL.
'
' SelectAll returns fields in this column order (0-based):
'   0 ID  |  1 symbol  |  2 name
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
' Bulk read  (called ONLY by clsElementCache.Build)
' ----------------------------------------------------------------
Public Function SelectAll() As Variant
    AssertInit "SelectAll"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAll = m_AccessDB.ExecuteQuery( _
        "SELECT ID, element_symbol, element_name " & _
        "FROM   elements " & _
        "ORDER BY element_symbol")
    Exit Function
EH:                                                         HandleError "SelectAll", "Failed"
End Function

' ----------------------------------------------------------------
' Writes
' ----------------------------------------------------------------
' Inserts a new element row. Returns the new autonumber ID.
'
Public Function Insert(E As clsElement) As Long
    AssertInit "Insert"
    Debug.Assert Not E Is Nothing
    Debug.Assert E.ID = 0               ' must be unsaved
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim SQL As String
    SQL = "INSERT INTO elements (element_symbol, element_name) " & _
          "VALUES (" & SqlStr(E.Symbol) & ", " & SqlStr(E.Name) & ")"

    Insert = m_AccessDB.ExecuteNonQuery(SQL)
    Exit Function
EH:                                                         HandleError "Insert", "Failed on: " & E.Symbol
End Function

Public Sub Update(E As clsElement)
    AssertInit "Update"
    Debug.Assert Not E Is Nothing
    Debug.Assert E.ID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim SQL As String
    SQL = "UPDATE elements SET " & _
          "    element_symbol = " & SqlStr(E.Symbol) & ", " & _
          "    element_name   = " & SqlStr(E.Name) & " " & _
          "WHERE ID = " & E.ID

    m_AccessDB.ExecuteNonQuery SQL
    Exit Sub
EH:                                                         HandleError "Update", "Failed on ID: " & E.ID
End Sub

Public Sub Delete(ElementID As Long)
    AssertInit "Delete"
    Debug.Assert ElementID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM elements WHERE ID = " & ElementID
    Exit Sub
EH:                                                         HandleError "Delete", "Failed on ID: " & ElementID
End Sub

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Function SqlStr(V As String) As String
    SqlStr = "'" & Replace(V, "'", "''") & "'"
End Function

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsElementRepository." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Context As String)
    m_Logger.LogError "clsElementRepository." & Src, Err.Number, Err.Description & " | " & Context
End Sub

