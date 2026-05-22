VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsElement
' Author: JT
' Created: 2026-04-10
' Description: Immutable-after-load entity representing one row
'              from the elements table.
' DependsOn: clsLoggingSystem
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release
'
' ================================================================
' LAYER: Entity — no DB or Service dependencies.
'
' LIFECYCLE:
'   1.  New clsElement
'   2.  .Initialize(Logger)
'   3.  .LoadFromDict(D)      <- hydrate from a repo row dict
'
' ID CONTRACT:
'   ID = 0 means unsaved. After INSERT the Service calls
'   .SetInternalID(newID) exactly once; ID is then sealed forever.
'
' Dict keys: element_id, element_symbol, element_name
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_ID     As Long
Private m_Symbol As String
Private m_Name   As String

' ----------------------------------------------------------------
' Initialization
' ----------------------------------------------------------------
Public Sub Initialize(Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init

    Set m_Logger = Logger
End Sub

' ----------------------------------------------------------------
' Read-only properties
' ----------------------------------------------------------------
Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Get Symbol() As String
    Symbol = m_Symbol
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

' ----------------------------------------------------------------
' Hydration  (repo row dict -> object)
' ----------------------------------------------------------------
' Required keys: element_id, element_symbol, element_name
'
Public Sub LoadFromDict(D As Scripting.Dictionary)
    AssertInit "LoadFromDict"
    Debug.Assert Not D Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    RequireKey D, "element_id"
    RequireKey D, "element_symbol"
    RequireKey D, "element_name"

    m_ID = CLng(D("element_id"))
    m_Symbol = CStr(D("element_symbol"))
    m_Name = CStr(D("element_name"))

    Exit Sub
EH:                                                         HandleError "LoadFromDict", "Keys present: " & Join(D.Keys, ", ")
End Sub

' ----------------------------------------------------------------
' Friend — called ONCE by clsElementService after INSERT
' ----------------------------------------------------------------
Friend Sub SetInternalID(NewID As Long)
    AssertInit "SetInternalID"
    Debug.Assert NewID > 0
    Debug.Assert m_ID = 0               ' must not already be set

    If NewID <= 0 Then Err.Raise vbObjectError + 400, "clsElement.SetInternalID", "ID must be positive, got: " & NewID
    If m_ID <> 0 Then Err.Raise vbObjectError + 401, "clsElement.SetInternalID", "ID already set to: " & m_ID
    
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_ID = NewID
    Exit Sub
EH:                                                                                                                                                                         HandleError "SetInternalID", "Failed to set ID: " & NewID & " for element: " & m_Symbol
End Sub

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub RequireKey(D As Scripting.Dictionary, Key As String)
    Debug.Assert D.Exists(Key)
    If Not D.Exists(Key) Then
        Err.Raise vbObjectError + 200, "clsElement.LoadFromDict", _
    "Missing required key: " & Key
    End If
End Sub

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsElement." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Optional Context As String = "")
    m_Logger.LogError "clsElement." & Src, Err.Number, Err.Description & " | " & Context
End Sub

