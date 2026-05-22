VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChemical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsChemical
' Author: JT
' Created: 2026-04-10
' Description: Immutable-after-load entity representing one row
'              from the chemicals table plus its KED element IDs
'              from the ked_elements junction table.
' DependsOn: clsLoggingSystem
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release
'
' ================================================================
' LAYER: Entity — no DB or Service dependencies.
'
' LIFECYCLE:
'   1.  New clsChemical
'   2.  .Initialize(Logger)
'   3.  .LoadFromDict(D)      <- hydrate from a repo row dict
'
' ID CONTRACT:
'   ID = 0 means unsaved. After INSERT the Service calls
'   .SetInternalID(newID) exactly once; ID is then sealed forever.
'
' KED ELEMENTS:
'   Stored internally as Long -> True (set semantics, no duplicates).
'   .KEDElementIDs always returns a defensive copy so callers
'   cannot mutate internal state.
' ================================================================

Private m_Logger        As clsLoggingSystem
Private m_ID            As Long
Private m_Name          As String
Private m_MetalsPrep    As String
Private m_SiliconPrep   As String
Private m_IonsPrep      As String
Private m_EntryDate     As Date
Private m_KEDElementIDs As Scripting.Dictionary   ' Long -> True

' ----------------------------------------------------------------
' Initialization
' ----------------------------------------------------------------
Public Sub Initialize(Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init

    Set m_Logger = Logger
    Set m_KEDElementIDs = New Scripting.Dictionary
End Sub

' ----------------------------------------------------------------
' Read-only properties
' ----------------------------------------------------------------
Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get MetalsPrep() As String
    MetalsPrep = m_MetalsPrep
End Property

Public Property Get SiliconPrep() As String
    SiliconPrep = m_SiliconPrep
End Property

Public Property Get IonsPrep() As String
    IonsPrep = m_IonsPrep
End Property

Public Property Get EntryDate() As Date
    EntryDate = m_EntryDate
End Property

' Returns a defensive copy — callers cannot mutate internal state.
Public Property Get KEDElementIDs() As Scripting.Dictionary
    Set KEDElementIDs = CopyDict(m_KEDElementIDs)
End Property

' ----------------------------------------------------------------
' Hydration  (repo row dict -> object)
' ----------------------------------------------------------------
' Required keys: chemical_id, chemical_name, metals_prep,
'                silicon_prep, ions_prep, entry_date
' Optional key:  ked_elements (Scripting.Dictionary, Long -> True)
'
Public Sub LoadFromDict(D As Scripting.Dictionary)
    AssertInit "LoadFromDict"
    Debug.Assert Not D Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    RequireKey D, "chemical_id"
    RequireKey D, "chemical_name"
    RequireKey D, "metals_prep"
    RequireKey D, "silicon_prep"
    RequireKey D, "ions_prep"
    RequireKey D, "database_entry_date"

    m_ID = CLng(D("chemical_id"))
    m_Name = CStr(D("chemical_name"))
    m_MetalsPrep = CStr(D("metals_prep"))
    m_SiliconPrep = CStr(D("silicon_prep"))
    m_IonsPrep = CStr(D("ions_prep"))
    If Not IsNull(D("database_entry_date")) Then
        m_EntryDate = CDate(D("database_entry_date"))
    End If

    If D.Exists("ked_elements") Then
        If TypeOf D("ked_elements") Is Scripting.Dictionary Then
            LoadElementIDs D("ked_elements")
        End If
    End If

    Exit Sub
EH:                                                                 HandleError "LoadFromDict", "Keys present: " & Join(D.Keys, ", ")
End Sub

' ----------------------------------------------------------------
' Friend — called ONCE by clsChemicalService after INSERT
' ----------------------------------------------------------------
Friend Sub SetInternalID(NewID As Long)
    AssertInit "SetInternalID"
    Debug.Assert NewID > 0
    Debug.Assert m_ID = 0               ' must not already be set

    If Not m_Logger.DebugMode Then On Error GoTo EH

    If NewID <= 0 Then Err.Raise vbObjectError + 400, "clsChemical.SetInternalID", "ID must be positive, got: " & NewID
    If m_ID <> 0 Then Err.Raise vbObjectError + 401, "clsChemical.SetInternalID", "ID already set to: " & m_ID

    m_ID = NewID

    Exit Sub
EH:
    HandleError "SetInternalID", "Failed to set ID: " & NewID & " for chemical: " & m_Name
End Sub

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub LoadElementIDs(Source As Scripting.Dictionary)
    Dim Key     As Variant
    Dim LongKey As Long

    Set m_KEDElementIDs = New Scripting.Dictionary
    For Each Key In Source.Keys
        LongKey = CLng(Key)
        If Not m_KEDElementIDs.Exists(LongKey) Then
            m_KEDElementIDs.Add LongKey, True
        End If
    Next Key
End Sub

Private Function CopyDict(Source As Scripting.Dictionary) As Scripting.Dictionary
    Dim D   As New Scripting.Dictionary
    Dim Key As Variant
    For Each Key In Source.Keys
        D.Add Key, Source(Key)
    Next Key
    Set CopyDict = D
End Function

Private Sub RequireKey(D As Scripting.Dictionary, Key As String)
    Debug.Assert D.Exists(Key)
    If Not D.Exists(Key) Then
        Err.Raise vbObjectError + 200, "clsChemical.LoadFromDict", _
    "Missing required key: " & Key
    End If
End Sub

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsChemical." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Optional Context As String = "")
    m_Logger.LogError "clsChemical." & Src, Err.Number, Err.Description & " | " & Context
End Sub


