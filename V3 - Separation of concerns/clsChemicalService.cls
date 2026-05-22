VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChemicalService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsChemicalService
' Author: JT
' Created: 2026-04-10
' Description: Orchestrates the chemicals domain.
'              THE only public API callers should interact with.
' DependsOn: clsLoggingSystem, clsAccessDatabase,
'            clsChemicalRepository, clsChemicalCache, clsChemical
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release
'
' ================================================================
' LAYER: Service — owns Repo and Cache; callers never touch them.
'
' READ PATTERN:
'   All reads go through the Cache — no DB round-trips.
'
' WRITE PATTERN:
'   Validate -> Repo write -> assign ID (if new) -> Cache.Upsert
'   Cache is updated incrementally; no full rebuild after writes.
'
' VALIDATION:
'   IsValid* and IsUnique* methods are public so UI controls can
'   pre-validate individual fields before a form is submitted.
'   The same checks are enforced inside CreateChemical.
'
' PREP LISTS:
'   The valid strings for MetalsPrep / SiliconPrep / IonsPrep live
'   here as private constants. Expose them via the List properties
'   for UI comboboxes so the source of truth is never duplicated.
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_Repo   As clsChemicalRepository
Private m_Cache  As clsChemicalCache

Private Const METALS_PREP  As String = "N/A,Evaporation,Dilute and Shoot,Organic Dilute and Shoot,Standard Addition"
Private Const SILICON_PREP As String = "N/A,Evaporation,Dilute and Shoot"
Private Const IONS_PREP    As String = "N/A,Dilute and Shoot,Evaporation,Neutralizing,Extraction,Standard Addition"

' ----------------------------------------------------------------
' Initialization  —  builds Repo and Cache internally
' ----------------------------------------------------------------
Public Sub Initialize(AccessDB As clsAccessDatabase, Logger As clsLoggingSystem)
    Debug.Assert Not AccessDB Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init

    If Not Logger.DebugMode Then On Error GoTo EH
    Set m_Logger = Logger

    Set m_Repo = New clsChemicalRepository
    m_Repo.Initialize AccessDB, Logger

    Set m_Cache = New clsChemicalCache
    m_Cache.Initialize m_Repo, Logger      ' Cache.Build() is called inside Initialize

    Exit Sub
EH:                                                                 HandleError "Initialize", "Failed"
End Sub

' ----------------------------------------------------------------
' Prep list accessors  (for UI comboboxes)
' ----------------------------------------------------------------
Public Property Get MetalsPrepList() As Variant
    MetalsPrepList = Split(METALS_PREP, ",")
End Property

Public Property Get SiliconPrepList() As Variant
    SiliconPrepList = Split(SILICON_PREP, ",")
End Property

Public Property Get IonsPrepList() As Variant
    IonsPrepList = Split(IONS_PREP, ",")
End Property

' ----------------------------------------------------------------
' Reads  (all via Cache — zero DB round-trips)
' ----------------------------------------------------------------
Public Function LoadChemical(ChemicalID As Long) As clsChemical
    AssertInit "LoadChemical"
    Set LoadChemical = m_Cache.GetByID(ChemicalID)
End Function

Public Function GetByName(Name As String) As clsChemical
    AssertInit "GetByName"
    Set GetByName = m_Cache.GetByName(Name)
End Function

Public Function AllChemicals() As Collection
    AssertInit "AllChemicals"
    Set AllChemicals = m_Cache.AllChemicals
End Function
Public Function GetIDByFormName(FormName As String) As Long
    AssertInit "GetIDByFormName"
    GetIDByFormName = m_Cache.GetIDByFormName(FormName)
End Function
Public Function GetByFormName(FormName As String) As clsChemical
    AssertInit "GetByFormName"
    Dim ChemicalID As Long
    ChemicalID = m_Cache.GetIDByFormName(FormName)
    If ChemicalID > 0 Then
        Set GetByFormName = m_Cache.GetByID(ChemicalID)
    End If
    ' Returns Nothing if not found — caller handles miss
End Function

Public Function FormNameExists(FormName As String) As Boolean
    AssertInit "FormNameExists"
    FormNameExists = m_Cache.ExistsByFormName(FormName)
End Function

Public Sub AddFormName(FormName As String, ChemicalID As Long)
    AssertInit "AddFormName"
    Debug.Assert Len(Trim(FormName)) > 0
    Debug.Assert ChemicalID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_Repo.InsertFormName FormName, ChemicalID
    m_Cache.AddFormNameIndex FormName, ChemicalID
    Exit Sub
EH:                             HandleError "AddFormName", "FormName: " & FormName
End Sub

' ----------------------------------------------------------------
' Writes
' ----------------------------------------------------------------
Public Function CreateChemical(D As Scripting.Dictionary) As clsChemical
    AssertInit "CreateChemical"
    Debug.Assert Not D Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    RequireValidName CStr(D("chemical_name"))
    RequireUniqueName CStr(D("chemical_name"))
    RequireValidMetals CStr(D("metals_prep"))
    RequireValidSilicon CStr(D("silicon_prep"))
    RequireValidIons CStr(D("ions_prep"))

    Dim C As New clsChemical
    C.Initialize m_Logger
    C.LoadFromDict BuildNewRecord(D)

    SaveChemical C
    Set CreateChemical = C
    Exit Function
EH:                                                                 HandleError "CreateChemical", "Failed on: " & D("chemical_name")
End Function

Public Sub SaveChemical(C As clsChemical)
    AssertInit "SaveChemical"
    Debug.Assert Not C Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    If C.ID = 0 Then
        Dim NewID As Long
        NewID = m_Repo.Insert(C)
        C.SetInternalID NewID       ' seal the ID on the object (Friend access)
    Else
        m_Repo.Update C
    End If

    m_Cache.Upsert C                ' incremental update — no full rebuild
    Exit Sub
EH:                                                                 HandleError "SaveChemical", "Failed on: " & C.Name
End Sub

Public Sub DeleteChemical(ChemicalID As Long)
    AssertInit "DeleteChemical"
    Debug.Assert ChemicalID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_Repo.Delete ChemicalID
    m_Cache.Remove ChemicalID
    Exit Sub
EH:                                                                 HandleError "DeleteChemical", "Failed on ID: " & ChemicalID
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

Public Function IsValidMetalsPrep(V As String) As Boolean
    IsValidMetalsPrep = IsInList(V, METALS_PREP)
End Function

Public Function IsValidSiliconPrep(V As String) As Boolean
    IsValidSiliconPrep = IsInList(V, SILICON_PREP)
End Function

Public Function IsValidIonsPrep(V As String) As Boolean
    IsValidIonsPrep = IsInList(V, IONS_PREP)
End Function

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Function BuildNewRecord(D As Scripting.Dictionary) As Scripting.Dictionary
    Dim R As New Scripting.Dictionary
    R.Add "chemical_id", 0
    R.Add "chemical_name", D("chemical_name")
    R.Add "metals_prep", D("metals_prep")
    R.Add "silicon_prep", D("silicon_prep")
    R.Add "ions_prep", D("ions_prep")
    R.Add "database_entry_date", Now
    If D.Exists("ked_elements") Then
        R.Add "ked_elements", D("ked_elements")
    End If
    Set BuildNewRecord = R
End Function

Private Function IsInList(V As String, List As String) As Boolean
    Dim Items() As String
    Dim i       As Long
    Items = Split(List, ",")
    For i = LBound(Items) To UBound(Items)
        If Trim$(Items(i)) = Trim$(V) Then
            IsInList = True
            Exit Function
        End If
    Next i
    IsInList = False
End Function

Private Sub RequireValidName(V As String)
    If Not IsValidName(V) Then
        Err.Raise vbObjectError + 400, , "Chemical name cannot be blank"
    End If
End Sub

Private Sub RequireUniqueName(V As String)
    If Not IsUniqueName(V) Then
        Err.Raise vbObjectError + 401, , "Chemical name already exists: " & V
    End If
End Sub

Private Sub RequireValidMetals(V As String)
    If Not IsValidMetalsPrep(V) Then
        Err.Raise vbObjectError + 402, , "Invalid MetalsPrep: " & V
    End If
End Sub

Private Sub RequireValidSilicon(V As String)
    If Not IsValidSiliconPrep(V) Then
        Err.Raise vbObjectError + 403, , "Invalid SiliconPrep: " & V
    End If
End Sub

Private Sub RequireValidIons(V As String)
    If Not IsValidIonsPrep(V) Then
        Err.Raise vbObjectError + 404, , "Invalid IonsPrep: " & V
    End If
End Sub

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsChemicalService." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Context As String)
    m_Logger.LogError "clsChemicalService." & Src, Err.Number, Err.Description & " | " & Context
End Sub


