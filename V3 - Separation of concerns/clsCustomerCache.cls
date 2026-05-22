VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomerCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsCustomerCache
' Author: JT
' Created: 2026-04-10
' Description: Materialised in-memory cache of all customers.
'              No dependency on clsCustomerService.
' DependsOn: clsLoggingSystem, clsCustomerRepository, clsCustomer
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release (replaces v0.1.0 AccessDB-direct version)
'
' ================================================================
' LAYER: Cache - depends on Repo only. No upward coupling to Service.
'
' LOOKUP CONTRACT:
'   GetByID / GetByName return Nothing if not found.
'   GetIDByName / GetNameByID raise if not found (use Exists* to guard).
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_Repo   As clsCustomerRepository
Private m_ByID   As Scripting.Dictionary   ' Long      -> clsCustomer
Private m_ByName As Scripting.Dictionary   ' LCase key -> clsCustomer
Private m_ByFormCustomer As Scripting.Dictionary ' composite key -> Long (customer_id)

' ----------------------------------------------------------------
' Initialization
' ----------------------------------------------------------------
Public Sub Initialize(Repo As clsCustomerRepository, Logger As clsLoggingSystem)
    Debug.Assert Not Repo Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init

    Set m_Logger = Logger
    Set m_Repo = Repo
    Set m_ByID = New Scripting.Dictionary
    Set m_ByName = New Scripting.Dictionary
    Set m_ByFormCustomer = New Scripting.Dictionary

    Build
End Sub

' ----------------------------------------------------------------
' Build  -  full rebuild from Repo
' ----------------------------------------------------------------

Public Sub Build()
    AssertInit "Build"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_ByID.RemoveAll
    m_ByName.RemoveAll
    m_ByFormCustomer.RemoveAll

    ' -------------------------------------------------------
    ' Main customer index
    ' -------------------------------------------------------
    Dim Customers As Variant
    Customers = m_Repo.SelectAll

    ' No customers in DB is valid — leave the cache empty.
    If IsEmpty(Customers) Then Exit Sub

    Dim i As Long
    Dim D As Scripting.Dictionary
    Dim C As clsCustomer

    ' SelectAll column order: 0=ID, 1=name, 2=street_address,
    '                         3=city, 4=state, 5=postal_code, 6=country
    For i = LBound(Customers, 2) To UBound(Customers, 2)
        Set D = New Scripting.Dictionary
        D.Add "customer_id", modUtilities.NullToDefault(Customers(0, i), 0)
        D.Add "customer_name", modUtilities.NullToDefault(Customers(1, i), "")
        D.Add "street_address", modUtilities.NullToDefault(Customers(2, i), "")
        D.Add "city", modUtilities.NullToDefault(Customers(3, i), "")
        D.Add "state", modUtilities.NullToDefault(Customers(4, i), "")
        D.Add "postal_code", modUtilities.NullToDefault(Customers(5, i), "")
        D.Add "country", modUtilities.NullToDefault(Customers(6, i), "")
        Set C = New clsCustomer
        C.Initialize m_Logger
        C.LoadFromDict D
        StoreInIndex C
    Next i

    ' -------------------------------------------------------
    ' Form customer index — form_name + address -> customer_id
    ' -------------------------------------------------------
    Dim FormCustomers As Variant
    FormCustomers = m_Repo.SelectAllFormCustomers

    ' No form mappings is valid — leave the form index empty.
    If IsEmpty(FormCustomers) Then Exit Sub

    Dim j As Long
    Dim FormName        As String
    Dim FormAddress     As String
    Dim FormAddress2    As String
    Dim CustomerID      As Long
    Dim Key             As String

    ' SelectAllFormCustomers column order:
    '   0=form_customer_name, 1=form_address, 2=form_address2, 3=customer_id
    For j = LBound(FormCustomers, 2) To UBound(FormCustomers, 2)
        FormName = modUtilities.NullToDefault(FormCustomers(0, j), "")
        FormAddress = modUtilities.NullToDefault(FormCustomers(1, j), "")
        FormAddress2 = modUtilities.NullToDefault(FormCustomers(2, j), "")
        CustomerID = modUtilities.NullToDefault(FormCustomers(3, j), 0)

        ' Skip malformed rows
        If FormName <> "" And CustomerID <> 0 Then

            Key = BuildCustomerKey(FormName, FormAddress, FormAddress2)
    
            ' Duplicate keys indicate a data problem — log and skip
            If m_ByFormCustomer.Exists(Key) Then
                m_Logger.LogMessage "clsCustomerCache.Build", LogWarning, "Duplicate form customer key skipped: " & Key
            Else
                m_ByFormCustomer.Add Key, CustomerID
            End If
            
        End If

    Next j

    Exit Sub
EH:
    HandleError "Build", "Failed at customer row " & i & " or form row " & j
End Sub

' ----------------------------------------------------------------
' Incremental updates  (called by Service after each write)
' ----------------------------------------------------------------
Public Sub Upsert(C As clsCustomer)
    AssertInit "Upsert"
    Debug.Assert Not C Is Nothing
    Debug.Assert C.ID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    ' Remove stale name key if the customer was renamed.
    If m_ByID.Exists(C.ID) Then
        Dim OldKey As String
        OldKey = NormName(m_ByID(C.ID).Name)
        If OldKey <> NormName(C.Name) Then
            If m_ByName.Exists(OldKey) Then m_ByName.Remove OldKey
        End If
    End If

    StoreInIndex C
    Exit Sub
EH:
    HandleError "Upsert", "Failed on ID: " & C.ID
End Sub

Public Sub Remove(CustomerID As Long)
    AssertInit "Remove"
    Debug.Assert CustomerID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    If m_ByID.Exists(CustomerID) Then
        Dim NameKey As String
        NameKey = NormName(m_ByID(CustomerID).Name)
        If m_ByName.Exists(NameKey) Then m_ByName.Remove NameKey
        m_ByID.Remove CustomerID
    End If
    Exit Sub
EH:
    HandleError "Remove", "Failed on ID: " & CustomerID
End Sub

' ----------------------------------------------------------------
' Read API
' ----------------------------------------------------------------
Public Function ExistsByID(CustomerID As Long) As Boolean
    AssertInit "ExistsByID"
    ExistsByID = m_ByID.Exists(CustomerID)
End Function

Public Function ExistsByName(Name As String) As Boolean
    AssertInit "ExistsByName"
    ExistsByName = m_ByName.Exists(NormName(Name))
End Function

Public Function GetByID(CustomerID As Long) As clsCustomer
    AssertInit "GetByID"
    If m_ByID.Exists(CustomerID) Then Set GetByID = m_ByID(CustomerID)
End Function

Public Function GetByName(Name As String) As clsCustomer
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
        Err.Raise vbObjectError + 103, "clsCustomerCache.GetIDByName", _
        "Customer not found: " & Name
    End If
End Function

Public Function GetNameByID(CustomerID As Long) As String
    AssertInit "GetNameByID"
    If m_ByID.Exists(CustomerID) Then
        GetNameByID = m_ByID(CustomerID).Name
    Else
        Err.Raise vbObjectError + 103, "clsCustomerCache.GetNameByID", _
        "Customer ID not found: " & CustomerID
    End If
End Function

Public Function AllCustomers() As Collection
    AssertInit "AllCustomers"
    Dim Coll As New Collection
    Dim Key  As Variant
    For Each Key In m_ByID.Keys
        Coll.Add m_ByID(Key)
    Next Key
    Set AllCustomers = Coll
End Function
Public Function GetIDByFormCustomer(Name As String, Address As String, Address2 As String) As Long
    AssertInit "GetIDByFormCustomer"
    Dim Key As String
    Key = BuildCustomerKey(Name, Address, Address2)
    If m_ByFormCustomer.Exists(Key) Then
        GetIDByFormCustomer = m_ByFormCustomer(Key)
    Else
        GetIDByFormCustomer = 0
    End If
End Function

Public Function ExistsByFormCustomer(Name As String, Address As String, Address2 As String) As Boolean
    AssertInit "ExistsByFormCustomer"
    ExistsByFormCustomer = m_ByFormCustomer.Exists(BuildCustomerKey(Name, Address, Address2))
End Function

Public Sub AddFormCustomerIndex(Name As String, Address As String, Address2 As String, CustomerID As Long)
    AssertInit "AddFormCustomerIndex"
    Debug.Assert Len(Trim(Name)) > 0
    Debug.Assert CustomerID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim Key As String
    Key = BuildCustomerKey(Name, Address, Address2)

    If Not m_ByFormCustomer.Exists(Key) Then
        m_ByFormCustomer.Add Key, CustomerID
    End If
    Exit Sub
EH:                             HandleError "AddFormCustomerIndex", "Name: " & Name
End Sub
' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub StoreInIndex(C As clsCustomer)
    Set m_ByID(C.ID) = C
    Set m_ByName(NormName(C.Name)) = C
End Sub

Private Function NormName(V As String) As String
    NormName = LCase$(Trim$(V))
End Function
Private Function BuildCustomerKey(Name As String, Address As String, Address2 As String) As String
    BuildCustomerKey = LCase$(Trim$(Name)) & "|" & LCase$(Trim$(Address)) & "|" & LCase$(Trim$(Address2))
End Function

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsCustomerCache." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Context As String)
    m_Logger.LogError "clsCustomerCache." & Src, Err.Number, Err.Description & " | " & Context
End Sub
