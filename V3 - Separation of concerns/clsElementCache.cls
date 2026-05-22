VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsElementCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsElementCache
' Author: JT
' Created: 2026-04-10
' Description: Materialised in-memory cache of all elements.
'              No dependency on clsElementService.
' DependsOn: clsLoggingSystem, clsElementRepository, clsElement
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release
'
' ================================================================
' LAYER: Cache — depends on Repo only. No upward coupling to Service.
'
' Three indexes are maintained:
'   m_ByID     Long      -> clsElement
'   m_ByName   LCase     -> clsElement
'   m_BySymbol LCase     -> clsElement
'
' LOOKUP CONTRACT:
'   GetBy* return Nothing if not found.
'   GetIDBy* / GetNameByID / GetSymbolByID raise if not found.
' ================================================================

Private m_Logger   As clsLoggingSystem
Private m_Repo     As clsElementRepository
Private m_ByID     As Scripting.Dictionary   ' Long      -> clsElement
Private m_ByName   As Scripting.Dictionary   ' LCase key -> clsElement
Private m_BySymbol As Scripting.Dictionary   ' LCase key -> clsElement

' ----------------------------------------------------------------
' Initialization
' ----------------------------------------------------------------
Public Sub Initialize(Repo As clsElementRepository, Logger As clsLoggingSystem)
    Debug.Assert Not Repo Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init

    Set m_Logger = Logger
    Set m_Repo = Repo
    Set m_ByID = New Scripting.Dictionary
    Set m_ByName = New Scripting.Dictionary
    Set m_BySymbol = New Scripting.Dictionary

    Build
End Sub

' ----------------------------------------------------------------
' Build  —  full rebuild from Repo
' ----------------------------------------------------------------
Public Sub Build()
    AssertInit "Build"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_ByID.RemoveAll
    m_ByName.RemoveAll
    m_BySymbol.RemoveAll

    Dim Elements As Variant
    Elements = m_Repo.SelectAll

    ' No elements in DB is valid — leave the cache empty.
    If IsEmpty(Elements) Then Exit Sub

    Dim i As Long
    Dim D As Scripting.Dictionary
    Dim E As clsElement

    ' SelectAll column order: 0=ID, 1=symbol, 2=name
    For i = LBound(Elements, 2) To UBound(Elements, 2)
        Set D = New Scripting.Dictionary
        D.Add "element_id", Elements(0, i)
        D.Add "element_symbol", Elements(1, i)
        D.Add "element_name", Elements(2, i)

        Set E = New clsElement
        E.Initialize m_Logger
        E.LoadFromDict D

        StoreInIndex E
    Next i

    Exit Sub
EH:                                                         HandleError "Build", "Failed at row " & i
End Sub

' ----------------------------------------------------------------
' Incremental updates  (called by Service after each write)
' ----------------------------------------------------------------
Public Sub Upsert(E As clsElement)
    AssertInit "Upsert"
    Debug.Assert Not E Is Nothing
    Debug.Assert E.ID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    ' Remove stale name/symbol keys if the element was renamed.
    If m_ByID.Exists(E.ID) Then
        Dim OldNameKey As String
        Dim OldSymKey  As String
        OldNameKey = NormName(m_ByID(E.ID).Name)
        OldSymKey = NormName(m_ByID(E.ID).Symbol)

        If OldNameKey <> NormName(E.Name) Then
            If m_ByName.Exists(OldNameKey) Then m_ByName.Remove OldNameKey
        End If
        If OldSymKey <> NormName(E.Symbol) Then
            If m_BySymbol.Exists(OldSymKey) Then m_BySymbol.Remove OldSymKey
        End If
    End If

    StoreInIndex E
    Exit Sub
EH:                                                         HandleError "Upsert", "Failed on ID: " & E.ID
End Sub

Public Sub Remove(ElementID As Long)
    AssertInit "Remove"
    Debug.Assert ElementID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    If m_ByID.Exists(ElementID) Then
        Dim NameKey As String
        Dim SymKey  As String
        NameKey = NormName(m_ByID(ElementID).Name)
        SymKey = NormName(m_ByID(ElementID).Symbol)
        If m_ByName.Exists(NameKey) Then m_ByName.Remove NameKey
        If m_BySymbol.Exists(SymKey) Then m_BySymbol.Remove SymKey
        m_ByID.Remove ElementID
    End If
    Exit Sub
EH:                                                         HandleError "Remove", "Failed on ID: " & ElementID
End Sub

' ----------------------------------------------------------------
' Read API
' ----------------------------------------------------------------
Public Function ExistsByID(ElementID As Long) As Boolean
    AssertInit "ExistsByID"
    ExistsByID = m_ByID.Exists(ElementID)
End Function

Public Function ExistsByName(Name As String) As Boolean
    AssertInit "ExistsByName"
    ExistsByName = m_ByName.Exists(NormName(Name))
End Function

Public Function ExistsBySymbol(Symbol As String) As Boolean
    AssertInit "ExistsBySymbol"
    ExistsBySymbol = m_BySymbol.Exists(NormName(Symbol))
End Function

Public Function GetByID(ElementID As Long) As clsElement
    AssertInit "GetByID"
    If m_ByID.Exists(ElementID) Then Set GetByID = m_ByID(ElementID)
End Function

Public Function GetByName(Name As String) As clsElement
    AssertInit "GetByName"
    Dim Key As String: Key = NormName(Name)
    If m_ByName.Exists(Key) Then Set GetByName = m_ByName(Key)
End Function

Public Function GetBySymbol(Symbol As String) As clsElement
    AssertInit "GetBySymbol"
    Dim Key As String: Key = NormName(Symbol)
    If m_BySymbol.Exists(Key) Then Set GetBySymbol = m_BySymbol(Key)
End Function

Public Function GetIDByName(Name As String) As Long
    AssertInit "GetIDByName"
    Dim Key As String: Key = NormName(Name)
    If m_ByName.Exists(Key) Then
        GetIDByName = m_ByName(Key).ID
    Else
        Err.Raise vbObjectError + 103, "clsElementCache.GetIDByName", _
            "Element not found: " & Name
    End If
End Function

Public Function GetIDBySymbol(Symbol As String) As Long
    AssertInit "GetIDBySymbol"
    Dim Key As String: Key = NormName(Symbol)
    If m_BySymbol.Exists(Key) Then
        GetIDBySymbol = m_BySymbol(Key).ID
    Else
        Err.Raise vbObjectError + 103, "clsElementCache.GetIDBySymbol", _
            "Element symbol not found: " & Symbol
    End If
End Function

Public Function GetNameByID(ElementID As Long) As String
    AssertInit "GetNameByID"
    If m_ByID.Exists(ElementID) Then
        GetNameByID = m_ByID(ElementID).Name
    Else
        Err.Raise vbObjectError + 103, "clsElementCache.GetNameByID", _
            "Element ID not found: " & ElementID
    End If
End Function

Public Function GetSymbolByID(ElementID As Long) As String
    AssertInit "GetSymbolByID"
    If m_ByID.Exists(ElementID) Then
        GetSymbolByID = m_ByID(ElementID).Symbol
    Else
        Err.Raise vbObjectError + 103, "clsElementCache.GetSymbolByID", _
            "Element ID not found: " & ElementID
    End If
End Function

Public Function AllElements() As Collection
    AssertInit "AllElements"
    Dim Coll As New Collection
    Dim Key  As Variant
    For Each Key In m_ByID.Keys
        Coll.Add m_ByID(Key)
    Next Key
    Set AllElements = Coll
End Function

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub StoreInIndex(E As clsElement)
    Set m_ByID(E.ID) = E
    Set m_ByName(NormName(E.Name)) = E
    Set m_BySymbol(NormName(E.Symbol)) = E
End Sub

Private Function NormName(V As String) As String
    NormName = LCase$(Trim$(V))
End Function

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsElementCache." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Context As String)
    m_Logger.LogError "clsElementCache." & Src, Err.Number, Err.Description & " | " & Context
End Sub

