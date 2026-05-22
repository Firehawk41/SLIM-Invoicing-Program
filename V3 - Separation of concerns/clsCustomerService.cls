VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomerService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsCustomerService
' Author: JT
' Created: 2026-04-10
' Description: Orchestrates the customers domain.
'              THE only public API callers should interact with.
' DependsOn: clsLoggingSystem, clsAccessDatabase,
'            clsCustomerRepository, clsCustomerCache, clsCustomer
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release (replaces v0.1.0 prototype)
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
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_Repo   As clsCustomerRepository
Private m_Cache  As clsCustomerCache

' ----------------------------------------------------------------
' Initialization  -  builds Repo and Cache internally
' ----------------------------------------------------------------
Public Sub Initialize(AccessDB As clsAccessDatabase, Logger As clsLoggingSystem)
    Debug.Assert Not AccessDB Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init

    If Not Logger.DebugMode Then On Error GoTo EH
    Set m_Logger = Logger

    Set m_Repo = New clsCustomerRepository
    m_Repo.Initialize AccessDB, Logger

    Set m_Cache = New clsCustomerCache
    m_Cache.Initialize m_Repo, Logger      ' Cache.Build() is called inside Initialize

    Exit Sub
EH:                                                             HandleError "Initialize", "Failed"
End Sub

' ----------------------------------------------------------------
' Reads  (all via Cache - zero DB round-trips)
' ----------------------------------------------------------------
Public Function LoadCustomer(CustomerID As Long) As clsCustomer
    AssertInit "LoadCustomer"
    Set LoadCustomer = m_Cache.GetByID(CustomerID)
End Function

Public Function GetByName(Name As String) As clsCustomer
    AssertInit "GetByName"
    Set GetByName = m_Cache.GetByName(Name)
End Function

Public Function AllCustomers() As Collection
    AssertInit "AllCustomers"
    Set AllCustomers = m_Cache.AllCustomers
End Function
Public Function GetIDByFormCustomer(FormName As String, FormAddress As String, FormAddress2 As String) As Long
    AssertInit "GetIDByFormCustomer"
    GetIDByFormCustomer = m_Cache.GetIDByFormCustomer(FormName, FormAddress, FormAddress2)
End Function
Public Function FormKeyExists(Name As String, Address As String, Address2 As String) As Boolean
    AssertInit "FormKeyExists"
    FormKeyExists = m_Cache.ExistsByFormCustomer(Name, Address, Address2)
End Function
Public Function GetNameByID(ID As Long) As String
    AssertInit "GetNameByID"
    GetNameByID = m_Cache.GetNameByID(ID)
End Function
Public Sub AddFormKey(Name As String, Address As String, Address2 As String, CustomerID As Long)
    AssertInit "AddFormKey"
    Debug.Assert Len(Trim(Name)) > 0
    Debug.Assert CustomerID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_Repo.InsertFormCustomer Name, Address, Address2, CustomerID
    m_Cache.AddFormCustomerIndex Name, Address, Address2, CustomerID
    Exit Sub
EH:                 HandleError "AddFormKey", "Name: " & Name & " | Address: " & Address & " | Address2: " & Address2
End Sub

' ----------------------------------------------------------------
' Writes
' ----------------------------------------------------------------
Public Function CreateCustomer(D As Scripting.Dictionary) As clsCustomer
    AssertInit "CreateCustomer"
    Debug.Assert Not D Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    RequireValidName CStr(D("customer_name"))
    RequireUniqueName CStr(D("customer_name"))
    RequireNonBlank CStr(D("street_address")), "Street address"
    RequireNonBlank CStr(D("city")), "City"
    RequireNonBlank CStr(D("state")), "State"
    RequireNonBlank CStr(D("postal_code")), "Postal code"
    RequireNonBlank CStr(D("country")), "Country"

    Dim C As New clsCustomer
    C.Initialize m_Logger
    C.LoadFromDict BuildNewRecord(D)

    SaveCustomer C
    Set CreateCustomer = C
    Exit Function
EH:                                                             HandleError "CreateCustomer", "Failed on: " & D("customer_name")
End Function

Public Sub SaveCustomer(C As clsCustomer)
    AssertInit "SaveCustomer"
    Debug.Assert Not C Is Nothing
    If Not m_Logger.DebugMode Then On Error GoTo EH

    If C.ID = 0 Then
        Dim NewID As Long
        NewID = m_Repo.Insert(C)
        C.SetInternalID NewID       ' seal the ID on the object (Friend access)
    Else
        m_Repo.Update C
    End If

    m_Cache.Upsert C                ' incremental update - no full rebuild
    Exit Sub
EH:                                                             HandleError "SaveCustomer", "Failed on: " & C.Name
End Sub

Public Sub DeleteCustomer(CustomerID As Long)
    AssertInit "DeleteCustomer"
    Debug.Assert CustomerID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_Repo.Delete CustomerID
    m_Cache.Remove CustomerID
    Exit Sub
EH:                                                             HandleError "DeleteCustomer", "Failed on ID: " & CustomerID
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

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Function BuildNewRecord(D As Scripting.Dictionary) As Scripting.Dictionary
    Dim R As New Scripting.Dictionary
    R.Add "customer_id", 0
    R.Add "customer_name", D("customer_name")
    R.Add "street_address", D("street_address")
    R.Add "city", D("city")
    R.Add "state", D("state")
    R.Add "postal_code", D("postal_code")
    R.Add "country", D("country")
    Set BuildNewRecord = R
End Function

Private Sub RequireValidName(V As String)
    If Not IsValidName(V) Then
        Err.Raise vbObjectError + 400, , "Customer name cannot be blank"
    End If
End Sub

Private Sub RequireUniqueName(V As String)
    If Not IsUniqueName(V) Then
        Err.Raise vbObjectError + 401, , "Customer name already exists: " & V
    End If
End Sub

Private Sub RequireNonBlank(V As String, FieldName As String)
    If Len(Trim$(V)) = 0 Then
        Err.Raise vbObjectError + 402, , FieldName & " cannot be blank"
    End If
End Sub

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsCustomerService." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Context As String)
    m_Logger.LogError "clsCustomerService." & Src, Err.Number, Err.Description & " | " & Context
End Sub
