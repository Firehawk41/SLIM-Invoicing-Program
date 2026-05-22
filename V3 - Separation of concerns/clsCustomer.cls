VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsCustomer
' Author: JT
' Created: 2026-04-10
' Description: Immutable-after-load entity representing one row
'              from the customers table.
' DependsOn: clsLoggingSystem
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release (replaces v2.0.0 ToDict/FromDict version)
'
' ================================================================
' LAYER: Entity - no DB or Service dependencies.
'
' LIFECYCLE:
'   1.  New clsCustomer
'   2.  .Initialize(Logger)
'   3.  .LoadFromDict(D)      <- hydrate from a repo row dict
'
' ID CONTRACT:
'   ID = 0 means unsaved. After INSERT the Service calls
'   .SetInternalID(newID) exactly once; ID is then sealed forever.
'
' Dict keys: customer_id, customer_name, street_address,
'            city, state, postal_code, country
' ================================================================

Private m_Logger        As clsLoggingSystem
Private m_ID            As Long
Private m_Name          As String
Private m_StreetAddress As String
Private m_City          As String
Private m_State         As String
Private m_PostalCode    As String
Private m_Country       As String

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

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get StreetAddress() As String
    StreetAddress = m_StreetAddress
End Property

Public Property Get City() As String
    City = m_City
End Property

Public Property Get State() As String
    State = m_State
End Property

Public Property Get PostalCode() As String
    PostalCode = m_PostalCode
End Property

Public Property Get Country() As String
    Country = m_Country
End Property

Public Property Get FormattedAddress() As String
    Dim Addr As String
    Addr = m_StreetAddress & ", " & m_City
    If m_PostalCode <> "" Then Addr = Addr & " " & m_PostalCode
    If m_Country <> "" Then Addr = Addr & ", " & m_Country
    FormattedAddress = Addr
End Property

' ----------------------------------------------------------------
' Hydration  (repo row dict -> object)
' ----------------------------------------------------------------
' Required keys: customer_id, customer_name, street_address,
'                city, postal_code, country
'
Public Sub LoadFromDict(D As Scripting.Dictionary)
    AssertInit "LoadFromDict"
    Debug.Assert Not D Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    RequireKey D, "customer_id"
    RequireKey D, "customer_name"
    RequireKey D, "street_address"
    RequireKey D, "city"
    RequireKey D, "state"
    RequireKey D, "postal_code"
    RequireKey D, "country"

    m_ID = CLng(D("customer_id"))
    m_Name = CStr(D("customer_name"))
    m_StreetAddress = CStr(D("street_address"))
    m_City = CStr(D("city"))
    m_State = CStr(D("state"))
    m_PostalCode = CStr(D("postal_code"))
    m_Country = CStr(D("country"))

    Exit Sub
EH:                                                             HandleError "LoadFromDict", "Keys present: " & Join(D.Keys, ", ")
End Sub

' ----------------------------------------------------------------
' Friend - called ONCE by clsCustomerService after INSERT
' ----------------------------------------------------------------
Friend Sub SetInternalID(NewID As Long)
    AssertInit "SetInternalID"
    Debug.Assert NewID > 0
    Debug.Assert m_ID = 0               ' must not already be set

    

    If NewID <= 0 Then Err.Raise vbObjectError + 400, "clsCustomer.SetInternalID", "ID must be positive, got: " & NewID
    If m_ID <> 0 Then Err.Raise vbObjectError + 401, "clsCustomer.SetInternalID", "ID already set to: " & m_ID
    
    If Not m_Logger.DebugMode Then On Error GoTo EH
    
    m_ID = NewID
    Exit Sub
EH:
    HandleError "SetInternalID", "Failed to set ID: " & NewID & " for customer: " & m_Name
End Sub

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub RequireKey(D As Scripting.Dictionary, Key As String)
    Debug.Assert D.Exists(Key)
    If Not D.Exists(Key) Then
        Err.Raise vbObjectError + 200, "clsCustomer.LoadFromDict", _
"Missing required key: " & Key
    End If
End Sub

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsCustomer." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Optional Context As String = "")
    m_Logger.LogError "clsCustomer." & Src, Err.Number, Err.Description & " | " & Context
End Sub
