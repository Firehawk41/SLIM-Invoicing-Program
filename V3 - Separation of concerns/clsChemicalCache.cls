VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChemicalCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsChemicalCache
' Author: JT
' Created: 2026-04-10
' Description: Materialised in-memory cache of all chemicals.
'              No dependency on clsChemicalService.
' DependsOn: clsLoggingSystem, clsChemicalRepository, clsChemical,
'               modUtilities
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release
'
' ================================================================
' LAYER: Cache — depends on Repo only. No upward coupling to Service.
'
' BUILD STRATEGY (two queries total, no N+1):
'   Build() calls:
'     1. Repo.SelectAll()             -> all chemical rows
'     2. Repo.SelectAllKEDElements()  -> all junction rows
'   BuildKEDMap() pre-indexes the junction rows by chemical_id.
'   Each Chemical is then hydrated in a single pass.
'
' WRITE STRATEGY (incremental, no full rebuild after writes):
'   After each Service write, the Service calls Upsert() or Remove().
'   Build() is only ever called once at init, or explicitly to refresh
'   if the underlying DB was modified from outside the application.
'
' LOOKUP CONTRACT:
'   GetByID / GetByName return Nothing if not found (safe for callers
'   that want to handle the miss themselves).
'   GetIDByName / GetNameByID raise if not found (use Exists* to guard).
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_Repo   As clsChemicalRepository
Private m_ByID   As Scripting.Dictionary   ' Long      -> clsChemical
Private m_ByName As Scripting.Dictionary   ' LCase key -> clsChemical
Private m_ByFormName As Scripting.Dictionary  ' form_name -> Long (single chemical ID)
' ----------------------------------------------------------------
' Initialization
' ----------------------------------------------------------------
Public Sub Initialize(Repo As clsChemicalRepository, Logger As clsLoggingSystem)
    Debug.Assert Not Repo Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init

    Set m_Logger = Logger
    Set m_Repo = Repo
    Set m_ByID = New Scripting.Dictionary
    Set m_ByName = New Scripting.Dictionary
    Set m_ByFormName = New Scripting.Dictionary

    Build
End Sub

' ----------------------------------------------------------------
' Build  —  full rebuild from Repo
' Called from Initialize; public to allow explicit refresh.
' ----------------------------------------------------------------
Public Sub Build()
    AssertInit "Build"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_ByID.RemoveAll
    m_ByName.RemoveAll
    m_ByFormName.RemoveAll

    Dim Chemicals As Variant
    Dim KEDRows   As Variant
    Chemicals = m_Repo.SelectAll
    KEDRows = m_Repo.SelectAllKEDElements

    ' No chemicals in DB is valid — leave the cache empty.
    If IsNull(Chemicals) Then Exit Sub

    ' Pre-index all KED rows by chemical_id before the main loop.
    Dim KEDMap As Scripting.Dictionary
    Set KEDMap = BuildKEDMap(KEDRows)

    ' Hoist loop variables — VBA has procedure scope, not block scope.
    Dim i      As Long
    Dim ChemID As Long
    Dim D      As Scripting.Dictionary
    Dim C      As clsChemical

    For i = LBound(Chemicals, 2) To UBound(Chemicals, 2)
        ChemID = CLng(Chemicals(0, i))

        Set D = New Scripting.Dictionary
        D.Add "chemical_id", Chemicals(0, i)
        D.Add "chemical_name", Chemicals(1, i)
        D.Add "metals_prep", Chemicals(2, i)
        D.Add "silicon_prep", Chemicals(3, i)
        D.Add "ions_prep", Chemicals(4, i)
        D.Add "database_entry_date", Chemicals(5, i)
        If KEDMap.Exists(ChemID) Then
            D.Add "ked_elements", KEDMap(ChemID)
        End If

        Set C = New clsChemical
        C.Initialize m_Logger
        C.LoadFromDict D

        StoreInIndex C
    Next i


    Dim FormRows As Variant
    FormRows = m_Repo.SelectAllFormChemicals()

    If Not IsEmpty(FormRows) Then
        Dim j As Long
        For j = LBound(FormRows, 2) To UBound(FormRows, 2)
            Dim FormName As String
            FormName = modUtilities.NullToDefault(FormRows(0, j), "")

            Dim ChemicalID As Long
            ChemicalID = modUtilities.NullToDefault(FormRows(1, j), 0)
            
            If FormName <> "" And ChemicalID <> 0 Then
                Dim Key As String
                Key = NormName(FormName)  ' normalize before storing
                If Not m_ByFormName.Exists(Key) Then
                    m_ByFormName.Add Key, ChemicalID
                End If
            End If

        Next j
    End If

    Exit Sub
EH:                                                                 HandleError "Build", "Failed at row " & i
End Sub

' ----------------------------------------------------------------
' Incremental updates  (called by Service after each write)
' ----------------------------------------------------------------
Public Sub Upsert(C As clsChemical)
    AssertInit "Upsert"
    Debug.Assert Not C Is Nothing
    Debug.Assert C.ID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    ' Remove stale name key if the chemical was renamed.
    If m_ByID.Exists(C.ID) Then
        Dim OldKey As String
        OldKey = NormName(m_ByID(C.ID).Name)
        If OldKey <> NormName(C.Name) Then
            If m_ByName.Exists(OldKey) Then m_ByName.Remove OldKey
        End If
    End If

    StoreInIndex C
    Exit Sub
EH:                                                                 HandleError "Upsert", "Failed on ID: " & C.ID
End Sub

Public Sub Remove(ChemicalID As Long)
    AssertInit "Remove"
    Debug.Assert ChemicalID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    If m_ByID.Exists(ChemicalID) Then
        Dim NameKey As String
        NameKey = NormName(m_ByID(ChemicalID).Name)
        If m_ByName.Exists(NameKey) Then m_ByName.Remove NameKey
        m_ByID.Remove ChemicalID
    End If
    Exit Sub
EH:                                                                 HandleError "Remove", "Failed on ID: " & ChemicalID
End Sub

' ----------------------------------------------------------------
' Read API
' ----------------------------------------------------------------
Public Function ExistsByID(ChemicalID As Long) As Boolean
    AssertInit "ExistsByID"
    ExistsByID = m_ByID.Exists(ChemicalID)
End Function

Public Function ExistsByName(Name As String) As Boolean
    AssertInit "ExistsByName"
    ExistsByName = m_ByName.Exists(NormName(Name))
End Function

Public Function GetByID(ChemicalID As Long) As clsChemical
    AssertInit "GetByID"
    If m_ByID.Exists(ChemicalID) Then Set GetByID = m_ByID(ChemicalID)
End Function

Public Function GetByName(Name As String) As clsChemical
    AssertInit "GetByName"
    Dim Key As String: Key = NormName(Name)
    If m_ByName.Exists(Key) Then Set GetByName = m_ByName(Key)
End Function

Public Function GetIDByName(Name As String) As Long
    AssertInit "GetIDByName"
    Dim Key As String: Key = NormName(Name)
    If m_ByName.Exists(Key) Then
        GetIDByName = m_ByName(Key).ID
    Else
        Err.Raise vbObjectError + 103, "clsChemicalCache.GetIDByName", _
            "Chemical not found: " & Name
    End If
End Function

Public Function GetNameByID(ChemicalID As Long) As String
    AssertInit "GetNameByID"
    If m_ByID.Exists(ChemicalID) Then
        GetNameByID = m_ByID(ChemicalID).Name
    Else
        Err.Raise vbObjectError + 103, "clsChemicalCache.GetNameByID", _
            "Chemical ID not found: " & ChemicalID
    End If
End Function

Public Function AllChemicals() As Collection
    AssertInit "AllChemicals"
    Dim Coll As New Collection
    Dim Key  As Variant
    For Each Key In m_ByID.Keys
        Coll.Add m_ByID(Key)
    Next Key
    Set AllChemicals = Coll
End Function
Public Function GetIDByFormName(FormName As String) As Long
    AssertInit "GetIDByFormName"
    Dim Key As String
    Key = NormName(FormName)  ' use the private function consistently

    If m_ByFormName.Exists(Key) Then
        GetIDByFormName = m_ByFormName(Key)
    Else
        GetIDByFormName = 0
    End If
End Function
Public Function ExistsByFormName(FormName As String) As Boolean
    AssertInit "ExistsByFormName"
    ExistsByFormName = m_ByFormName.Exists(NormName(FormName))
End Function

Public Sub AddFormNameIndex(FormName As String, ChemicalID As Long)
    AssertInit "AddFormNameIndex"
    Debug.Assert Len(Trim(FormName)) > 0
    Debug.Assert ChemicalID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim Key As String
    Key = NormName(FormName)

    ' Silently skip if already indexed — duplicate form names
    ' are a data issue, not a cache issue
    If Not m_ByFormName.Exists(Key) Then
        m_ByFormName.Add Key, ChemicalID
    End If
    Exit Sub
EH:                             HandleError "AddFormNameIndex", "FormName: " & FormName
End Sub
' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub StoreInIndex(C As clsChemical)
    Set m_ByID(C.ID) = C
    Set m_ByName(NormName(C.Name)) = C
End Sub

Private Function BuildKEDMap(KEDRows As Variant) As Scripting.Dictionary
    Dim Map     As New Scripting.Dictionary
    Dim ElemSet As Scripting.Dictionary
    Dim i       As Long
    Dim CID     As Long
    Dim EID     As Long

    If IsEmpty(KEDRows) Then
        Set BuildKEDMap = Map
        Exit Function
    End If

    For i = LBound(KEDRows, 2) To UBound(KEDRows, 2)
        CID = CLng(KEDRows(0, i))
        EID = CLng(KEDRows(1, i))

        If Not Map.Exists(CID) Then
            Map.Add CID, New Scripting.Dictionary
        End If

        Set ElemSet = Map(CID)
        If Not ElemSet.Exists(EID) Then ElemSet.Add EID, True
    Next i

    Set BuildKEDMap = Map
End Function

Private Function NormName(V As String) As String
    NormName = LCase$(Trim$(V))
End Function

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsChemicalCache." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Context As String)
    m_Logger.LogError "clsChemicalCache." & Src, Err.Number, Err.Description & " | " & Context
End Sub


