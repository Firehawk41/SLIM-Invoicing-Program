VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsQuote
' Author: JT
' Created: 2025-11-21
' Description: Entity representing one quote record. Stores
'              effective/expiry dates, payment info, creator,
'              a collection of customer IDs, and a collection
'              of clsQuoteLineItem objects.
' DependsOn: clsLoggingSystem, clsQuoteLineItem, modEnums
' ChangeLog:
'   - 1.3.0 - 2026-04-14 - Removed CustomerManager
'   - 1.4.0 - 2026-04-28 - FromDict renamed LoadFromDict. ToDict
'                           removed — repo reads properties directly.
'                           ValidateStateDebug removed — rules moved
'                           to Validate. Validate returns Boolean.
'                           SetInternalID added. RequireKey added.
'                           Full Property Get surface added.
'                           ID and CustomerIDs exposed.
'                           PaymentType typed to PaymentTypeEnum.
'                           CollectionToArray/ArrayToCollection
'                           removed. Header updated to new format.
'
' ================================================================
' LAYER: Entity - no DB, no Worksheet, no service dependencies.
'
' LIFECYCLE:
'   1. New clsQuote
'   2. .Initialize(Logger)
'   3. .LoadFromDict(D)           <- scalar fields only
'   4. .UpsertLineItem(LI) x N    <- loader attaches line items
'   5. .AddCustomerID(ID) x N     <- loader attaches customer IDs
'
'   Steps 3-5 performed by clsQuoteLoader.
'
'   ID starts at 0. Sealed by clsQuoteRepository via
'   Friend SetInternalID after INSERT.
'
'   PaymentTerms is stored on the quote as a per-customer
'   negotiated value. A future clsCustomerBillingProfile entity
'   would be the canonical home — retained here as a pragmatic
'   interim solution.
'
'   IsDefaultQuote — default quotes apply to all customers as
'   fallback pricing when no customer-specific quote matches.
'
'   Repo reads properties directly — no ToDict.
'
' LoadFromDict keys:
'   ID, effective_date, expiry_date, payment_type, payment_value,
'   created_by, comments, created_date, modified_date, valid,
'   is_default_quote
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_ID As Long
Private m_CustomerIDs As Collection
Private m_LineItems As Collection
Private m_EffectiveDate As Date
Private m_ExpireDate As Date
Private m_PaymentType As PaymentTypeEnum
Private m_PaymentTerms As Long
Private m_CreatedBy As String
Private m_Comments As String
Private m_CreatedDate As Date
Private m_ModifiedDate As Date
Private m_IsValid As Boolean
Private m_IsDefaultQuote As Boolean

' ================================================================
' Get Properties
' ================================================================
Public Property Get ID() As Long
    ID = m_ID
End Property
Public Property Get EffectiveDate() As Date
    EffectiveDate = m_EffectiveDate
End Property
Public Property Get ExpireDate() As Date
    ExpireDate = m_ExpireDate
End Property
Public Property Get PaymentType() As PaymentTypeEnum
    PaymentType = m_PaymentType
End Property
Public Property Get PaymentTerms() As Long
    PaymentTerms = m_PaymentTerms
End Property
Public Property Get CreatedBy() As String
    CreatedBy = m_CreatedBy
End Property
Public Property Get Comments() As String
    Comments = m_Comments
End Property
Public Property Get CreatedDate() As Date
    CreatedDate = m_CreatedDate
End Property
Public Property Get ModifiedDate() As Date
    ModifiedDate = m_ModifiedDate
End Property
Public Property Get IsValid() As Boolean
    IsValid = m_IsValid
End Property
Public Property Get IsDefaultQuote() As Boolean
    IsDefaultQuote = m_IsDefaultQuote
End Property
Public Property Get LineItems() As Collection
    Dim Copy As New Collection
    Dim LI As clsQuoteLineItem
    For Each LI In m_LineItems
        Copy.Add LI
    Next LI
    Set LineItems = Copy
End Property
Public Property Get CustomerIDs() As Collection
    Dim Copy As New Collection
    Dim ID As Variant
    For Each ID In m_CustomerIDs
        Copy.Add ID
    Next ID
    Set CustomerIDs = Copy
End Property

' ================================================================
' Initialization
' ================================================================
Public Sub Initialize(Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    If Not Logger.DebugMode Then On Error GoTo EH

    Set m_Logger = Logger
    Set m_CustomerIDs = New Collection
    Set m_LineItems = New Collection

    Exit Sub
EH:
    HandleError "Initialize"
End Sub

' ================================================================
' APIs
' ================================================================
Public Sub LoadFromDict(D As Scripting.Dictionary)
    Debug.Assert Not D Is Nothing

    AssertInitialized "LoadFromDict"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    RequireKey D, "ID"
    RequireKey D, "effective_date"
    RequireKey D, "expiry_date"
    RequireKey D, "payment_type"
    RequireKey D, "payment_value"
    RequireKey D, "created_by"
    RequireKey D, "created_date"
    RequireKey D, "valid"
    RequireKey D, "is_default_quote"

    m_ID = NullToDefault(D("ID"), 0)
    m_EffectiveDate = NullToDefault(D("effective_date"), 0)
    m_ExpireDate = NullToDefault(D("expiry_date"), 0)
    m_PaymentType = NullToDefault(D("payment_value"), 0)
    m_PaymentTerms = NullToDefault(D("payment_value"), 0)
    m_CreatedBy = NullToDefault(D("created_by"), "")
    m_Comments = NullToDefault(D("comments"), "")
    m_CreatedDate = NullToDefault(D("created_date"), 0)
    m_IsValid = NullToDefault(D("valid"), False)
    m_IsDefaultQuote = NullToDefault(D("is_default_quote"), False)

    If D.Exists("modified_date") Then
        If Not IsNull(D("modified_date")) Then
            m_ModifiedDate = D("modified_date")
        End If
    End If

    Exit Sub
EH:
    HandleError "LoadFromDict", "Failed on ID: " & D("ID")
End Sub

Public Sub UpsertLineItem(LI As clsQuoteLineItem)
    Debug.Assert Not LI Is Nothing

    AssertInitialized "UpsertLineItem"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    If LI.ID = 0 Then
        m_LineItems.Add LI
    Else
        Dim Existing As clsQuoteLineItem
        Set Existing = FindLineItemByID(LI.ID)
        If Not Existing Is Nothing Then
            ReplaceLineItem Existing, LI
        Else
            m_LineItems.Add LI
        End If
    End If

    Exit Sub
EH:
    HandleError "UpsertLineItem", "Failed on ID: " & LI.ID
End Sub

Public Sub AddCustomerID(CustomerID As Long)
    AssertInitialized "AddCustomerID"

    Dim ID As Variant
    For Each ID In m_CustomerIDs
        If ID = CustomerID Then Exit Sub
    Next ID
    m_CustomerIDs.Add CustomerID
End Sub

Public Function Validate() As Boolean
    AssertInitialized "Validate"

    Validate = True

    If m_ID < 0 Then Validate = False
    If Not m_IsDefaultQuote Then
        If m_CustomerIDs.Count = 0 Then Validate = False
    End If
    If m_LineItems.Count = 0 Then Validate = False
    If m_PaymentType < PaymentTypeEnum.PONumber Or _
       m_PaymentType > PaymentTypeEnum.CreditCard Then Validate = False
    If m_PaymentTerms < 0 Then Validate = False
    If m_CreatedBy = "" Then Validate = False
    If m_EffectiveDate = 0 Then Validate = False
    If m_ExpireDate = 0 Then Validate = False
    If m_CreatedDate = 0 Then Validate = False
    If m_ExpireDate < m_EffectiveDate Then Validate = False
End Function

Friend Sub SetInternalID(NewID As Long)
    Debug.Assert m_ID = 0
    m_ID = NewID
End Sub

' ================================================================
' Private Helpers
' ================================================================
Private Function FindLineItemByID(LineItemID As Long) As clsQuoteLineItem
    Dim LI As clsQuoteLineItem
    For Each LI In m_LineItems
        If LI.ID = LineItemID Then
            Set FindLineItemByID = LI
            Exit Function
        End If
    Next LI
    Set FindLineItemByID = Nothing
End Function

Private Sub ReplaceLineItem(OldLI As clsQuoteLineItem, NewLI As clsQuoteLineItem)
    Dim i As Long
    Dim LI As clsQuoteLineItem
    i = 1
    For Each LI In m_LineItems
        If LI.ID = OldLI.ID Then
            m_LineItems.Remove i
            m_LineItems.Add NewLI, , i
            Exit Sub
        End If
        i = i + 1
    Next LI
    m_LineItems.Add NewLI
End Sub

Private Sub RequireKey(D As Scripting.Dictionary, Key As String)
    Debug.Assert D.Exists(Key)
    If Not D.Exists(Key) Then
        Err.Raise vbObjectError + 200, "clsQuote.LoadFromDict", _
            "Missing required key: " & Key
    End If
End Sub

Private Sub HandleError(MethodName As String, Optional Context As String = "")
    m_Logger.LogError "clsQuote." & MethodName, Err.Number, _
        Err.Description & IIf(Context <> "", " | " & Context, "")
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsQuote." & MethodName, _
            "The object must be initialized before use."
    End If
End Sub

