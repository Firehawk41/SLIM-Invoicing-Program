VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChemicalRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsChemicalRepository
' Author: JT
' Created: 2026-04-10
' Description: Owns ALL SQL for the chemicals domain.
'              No business logic. No object graph traversal.
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsChemical
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release
'   - 1.1.0 - 2026-04-26 - Added InsertFormName
'
' ================================================================
' LAYER: Repository — translates between objects and SQL.
'
' All write methods (Insert/Update/Delete) are fully transactional.
' On any error the transaction is rolled back before the error
' is logged, so the DB is never left in a partial state.
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
'   0 chemical_id  |  1 chemical_name  |  2 metals_prep
'   3 silicon_prep |  4 ions_prep      |  5 database_entry_date
'
' SelectAllKEDElements returns:
'   0 chemical_id  |  1 element_id
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
' Bulk reads  (called ONLY by clsChemicalCache.Build)
' ----------------------------------------------------------------
Public Function SelectAll() As Variant
    AssertInit "SelectAll"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAll = m_AccessDB.ExecuteQuery("SELECT ID, chemical_name, metals_prep, silicon_prep, ions_prep, database_entry_date " & _
                "FROM chemicals " & _
                "ORDER BY chemical_name")
    Exit Function
EH:                                                                 HandleError "SelectAll", "Failed"
End Function

Public Function SelectAllKEDElements() As Variant
    AssertInit "SelectAllKEDElements"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllKEDElements = m_AccessDB.ExecuteQuery( _
        "SELECT chemical_id, element_id " & _
        "FROM   ked_elements " & _
        "ORDER BY chemical_id")
    Exit Function
EH:                                                                 HandleError "SelectAllKEDElements", "Failed"
End Function
Public Function SelectAllFormChemicals() As Variant
    AssertInit "SelectAllFormChemicals"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllFormChemicals = m_AccessDB.ExecuteQuery( _
        "SELECT form_chemical_name, chemical_id " & _
        "FROM   form_chemicals " & _
        "ORDER BY form_chemical_name")
    Exit Function
EH:                             HandleError "SelectAllFormChemicals", "Failed"
End Function

' ----------------------------------------------------------------
' Form name mapping writes
' ----------------------------------------------------------------
' Persists a new form_chemical_name -> chemical_id mapping.
' Called by clsChemicalService.AddFormName after UI resolution.
' No transaction needed — single row insert, no FK children.
'
Public Sub InsertFormName(FormName As String, ChemicalID As Long)
    AssertInit "InsertFormName"
    Debug.Assert Len(Trim(FormName)) > 0
    Debug.Assert ChemicalID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_AccessDB.ExecuteNonQuery _
        "INSERT INTO form_chemicals (form_chemical_name, chemical_id) " & _
        "VALUES (" & SqlStr(FormName) & ", " & ChemicalID & ")"
    Exit Sub
EH:
    HandleError "InsertFormName", "FormName: " & FormName & " ChemicalID: " & ChemicalID
End Sub
' ----------------------------------------------------------------
' Writes
' ----------------------------------------------------------------
' Inserts a new chemical row + all KED element rows in one
' transaction. Returns the new autonumber ID.
' The Service calls C.SetInternalID(result) after this returns.
'
Public Function Insert(C As clsChemical) As Long
    AssertInit "Insert"
    Debug.Assert Not C Is Nothing
    Debug.Assert C.ID = 0               ' must be unsaved
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim SQL As String
    SQL = "INSERT INTO chemicals " & _
          "    (chemical_name, metals_prep, silicon_prep, ions_prep, database_entry_date) " & _
          "VALUES (" & _
              SqlStr(C.Name) & ", " & _
              SqlStr(C.MetalsPrep) & ", " & _
              SqlStr(C.SiliconPrep) & ", " & _
              SqlStr(C.IonsPrep) & ", " & _
              SqlDate(C.EntryDate) & ")"

    Dim NewID As Long
    NewID = m_AccessDB.ExecuteNonQuery(SQL, True)

    InsertKEDElements NewID, C.KEDElementIDs

    m_AccessDB.CommitTransaction
    Insert = NewID
    Exit Function
EH:
    m_AccessDB.RollbackTransaction
    HandleError "Insert", "Failed on: " & C.Name
End Function

' Updates the chemical row and syncs KED elements in one transaction.
'
Public Sub Update(C As clsChemical)
    AssertInit "Update"
    Debug.Assert Not C Is Nothing
    Debug.Assert C.ID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim SQL As String
    SQL = "UPDATE chemicals SET " & _
          "    chemical_name = " & SqlStr(C.Name) & ", " & _
          "    metals_prep   = " & SqlStr(C.MetalsPrep) & ", " & _
          "    silicon_prep  = " & SqlStr(C.SiliconPrep) & ", " & _
          "    ions_prep     = " & SqlStr(C.IonsPrep) & " " & _
          "WHERE chemical_id = " & C.ID

    m_AccessDB.ExecuteNonQuery SQL, True
    SyncKEDElements C.ID, C.KEDElementIDs

    m_AccessDB.CommitTransaction
    Exit Sub
EH:
    m_AccessDB.RollbackTransaction
    HandleError "Update", "Failed on ID: " & C.ID
End Sub

' Deletes KED element rows first, then the chemical row,
' in one transaction. Order matters for FK constraints.
'
Public Sub Delete(ChemicalID As Long)
    AssertInit "Delete"
    Debug.Assert ChemicalID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM ked_elements WHERE chemical_id = " & ChemicalID, True
    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM chemicals WHERE chemical_id = " & ChemicalID
    m_AccessDB.CommitTransaction
    Exit Sub
EH:
    m_AccessDB.RollbackTransaction
    HandleError "Delete", "Failed on ID: " & ChemicalID
End Sub

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub InsertKEDElements(ChemicalID As Long, ElementIDs As Scripting.Dictionary)
    Dim Key As Variant
    For Each Key In ElementIDs.Keys
        m_AccessDB.ExecuteNonQuery _
            "INSERT INTO ked_elements (chemical_id, element_id) " & _
            "VALUES (" & ChemicalID & ", " & CLng(Key) & ")"
    Next Key
End Sub

Private Sub SyncKEDElements(ChemicalID As Long, Desired As Scripting.Dictionary)
    Dim Existing As Scripting.Dictionary
    Set Existing = LoadKEDFromDB(ChemicalID)

    Dim Key As Variant

    For Each Key In Desired.Keys
        If Not Existing.Exists(Key) Then
            m_AccessDB.ExecuteNonQuery _
                "INSERT INTO ked_elements (chemical_id, element_id) " & _
                "VALUES (" & ChemicalID & ", " & CLng(Key) & ")"
        End If
    Next Key

    For Each Key In Existing.Keys
        If Not Desired.Exists(Key) Then
            m_AccessDB.ExecuteNonQuery _
                "DELETE FROM ked_elements " & _
                "WHERE chemical_id = " & ChemicalID & _
                "  AND element_id  = " & CLng(Key)
        End If
    Next Key
End Sub

Private Function LoadKEDFromDB(ChemicalID As Long) As Scripting.Dictionary
    Dim D      As New Scripting.Dictionary
    Dim Result As Variant
    Dim i      As Long
    Dim EID    As Long

    Result = m_AccessDB.ExecuteQuery( _
        "SELECT element_id FROM ked_elements WHERE chemical_id = " & ChemicalID)

    If Not IsEmpty(Result) Then
        For i = LBound(Result, 2) To UBound(Result, 2)
            EID = CLng(Result(0, i))
            If Not D.Exists(EID) Then D.Add EID, True
        Next i
    End If

    Set LoadKEDFromDB = D
End Function

Private Function SqlStr(V As String) As String
    SqlStr = "'" & Replace(V, "'", "''") & "'"
End Function

Private Function SqlDate(V As Date) As String
    SqlDate = "#" & Format(V, "yyyy-mm-dd hh:mm:ss") & "#"
End Function

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsChemicalRepository." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Context As String)
    m_Logger.LogError "clsChemicalRepository." & Src, Err.Number, Err.Description & " | " & Context
End Sub


