VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsElementService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsElementService
' Author: JT
' Created: 2026-04-10
' Description: Orchestrates the elements domain.
'              THE only public API callers should interact with.
' DependsOn: clsLoggingSystem, clsAccessDatabase,
'            clsElementRepository, clsElementCache, clsElement
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
'
' VALIDATION:
'   IsValid* and IsUnique* are public so UI can pre-validate.
'   Symbol uniqueness and name uniqueness are both enforced.
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_Repo   As clsElementRepository
Private m_Cache  As clsElementCache

' ----------------------------------------------------------------
' Initialization  -  builds Repo and Cache internally
' ----------------------------------------------------------------
Public Sub Initialize(AccessDB As clsAccessDatabase, Logger As clsLoggingSystem)
    Debug.Assert Not AccessDB Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init


    Set m_Logger = Logger

    If Not Logger.DebugMode Then On Error GoTo EH
    
    Set m_Repo = New clsElementRepository
    m_Repo.Initialize AccessDB, Logger

    Set m_Cache = New clsElementCache
    m_Cache.Initialize m_Repo, Logger      ' Cache.Build() is called inside Initialize

    Exit Sub
EH:                                                         HandleError "Initialize", "Failed"
End Sub

' ----------------------------------------------------------------
' Reads  (all via Cache - zero DB round-trips)
' ----------------------------------------------------------------
Public Function LoadElement(ElementID As Long) As clsElement
    AssertInit "LoadElement"
    Set LoadElement = m_Cache.GetByID(ElementID)
End Function

Public Function GetByName(Name As String) As clsElement
    AssertInit "GetByName"
    Set GetByName = m_Cache.GetByName(Name)
End Function

Public Function GetBySymbol(Symbol As String) As clsElement
    AssertInit "GetBySymbol"
    Set GetBySymbol = m_Cache.GetBySymbol(Symbol)
End Function

Public Function AllElements() As Collection
    AssertInit "AllElements"
    Set AllElements = m_Cache.AllElements
End Function

' ----------------------------------------------------------------
' Writes
' ----------------------------------------------------------------
Public Function CreateElement(D As Scripting.Dictionary) As clsElement
    AssertInit "CreateElement"
    Debug.Assert Not D Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    RequireValidSymbol CStr(D("element_symbol"))
    RequireUniqueSymbol CStr(D("element_symbol"))
    RequireValidName CStr(D("element_name"))
    RequireUniqueName CStr(D("element_name"))

    Dim E As New clsElement
    E.Initialize m_Logger
    E.LoadFromDict BuildNewRecord(D)

    SaveElement E
    Set CreateElement = E
    Exit Function
EH:                                                         HandleError "CreateElement", "Failed on: " & D("element_symbol")
End Function

Public Sub SaveElement(E As clsElement)
    AssertInit "SaveElement"
    Debug.Assert Not E Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    If E.ID = 0 Then
        Dim NewID As Long
        NewID = m_Repo.Insert(E)
        E.SetInternalID NewID       ' seal the ID on the object (Friend access)
    Else
        m_Repo.Update E
    End If

    m_Cache.Upsert E                ' incremental update - no full rebuild
    Exit Sub
EH:                                                         HandleError "SaveElement", "Failed on: " & E.Symbol
End Sub

Public Sub DeleteElement(ElementID As Long)
    AssertInit "DeleteElement"
    Debug.Assert ElementID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_Repo.Delete ElementID
    m_Cache.Remove ElementID
    Exit Sub
EH:                                                         HandleError "DeleteElement", "Failed on ID: " & ElementID
End Sub

' ----------------------------------------------------------------
' Validation  (public so UI can check fields before submit)
' ----------------------------------------------------------------
Public Function IsValidSymbol(V As String) As Boolean
    IsValidSymbol = (Len(Trim$(V)) > 0)
End Function

Public Function IsUniqueSymbol(V As String) As Boolean
    AssertInit "IsUniqueSymbol"
    IsUniqueSymbol = Not m_Cache.ExistsBySymbol(V)
End Function

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
    R.Add "element_id", 0
    R.Add "element_symbol", D("element_symbol")
    R.Add "element_name", D("element_name")
    Set BuildNewRecord = R
End Function

Private Sub RequireValidSymbol(V As String)
    If Not IsValidSymbol(V) Then
        Err.Raise vbObjectError + 400, , "Element symbol cannot be blank"
    End If
End Sub

Private Sub RequireUniqueSymbol(V As String)
    If Not IsUniqueSymbol(V) Then
        Err.Raise vbObjectError + 401, , "Element symbol already exists: " & V
    End If
End Sub

Private Sub RequireValidName(V As String)
    If Not IsValidName(V) Then
        Err.Raise vbObjectError + 402, , "Element name cannot be blank"
    End If
End Sub

Private Sub RequireUniqueName(V As String)
    If Not IsUniqueName(V) Then
        Err.Raise vbObjectError + 403, , "Element name already exists: " & V
    End If
End Sub

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsElementService." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Context As String)
    m_Logger.LogError "clsElementService." & Src, Err.Number, Err.Description & " | " & Context
End Sub

