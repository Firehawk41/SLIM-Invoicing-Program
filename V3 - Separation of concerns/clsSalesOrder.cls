VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalesOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsSalesOrder
' Author: JT
' Created: 2026-03-13
' Description: Entity representing one sales order. Stores header-
'              level properties and a collection of
'              clsSalesOrderLineItem objects. Persisted to the
'              sales_orders table via clsSalesOrderRepository.
' DependsOn: clsLoggingSystem, clsSalesOrderLineItem,
'            clsSalesOrderData, modEnums
' ChangeLog:
'   - 1.0.0 - 2026-04-28 - Revised from clsInvoiceSalesOrder.
'                           SSACustomerID removed — Netsuite coupling.
'                           InvoiceEmail removed — owned by customer
'                           domain, not sales order.
'                           PaymentTerms typed to Long.
'                           LoadFromData added as creation path.
'                           LoadFromDict added as rehydration path.
'                           ToDict/FromDict removed.
'                           Full Property Get surface added.
'                           Validate implemented with real rules.
'                           TotalValue computed property added.
'
' ================================================================
' LAYER: Entity - no DB, no Worksheet, no service dependencies.
'
' LIFECYCLE:
'   1. New clsSalesOrder
'   2. .Initialize(Logger)
'   3. .LoadFromData(Data)      <- creation path
'                                   called by clsSalesOrderService
'                                   .CreateFromData after builder
'                                   produces clsSalesOrderData
'
'   ID starts at 0. Sealed by clsSalesOrderRepository via
'   Friend SetInternalID after INSERT.
'
'   SubmissionRef carries the source filename for audit traceability.
'   SubmissionID is the FK to the TR tree (persistence not yet active).
'
'   PaymentType drives Validate — if PONumber then PONumber must
'   be non-blank.
'
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_ID As Long
Private m_CustomerName As String
Private m_CustomerID As Long
Private m_SubmissionID As Long
Private m_SubmissionRef As String
Private m_DateReceived As Date
Private m_PONumber As String
Private m_PaymentType As PaymentTypeEnum
Private m_PaymentTerms As Long
Private m_RequesterName As String
Private m_RequesterEmail As String
Private m_ServiceDate As Date
Private m_LineItems As Collection

' ================================================================
' Get Properties
' ================================================================
Public Property Get ID() As Long
    ID = m_ID
End Property
Public Property Get CustomerName() As String
    CustomerName = m_CustomerName
End Property
Public Property Get CustomerID() As Long
    CustomerID = m_CustomerID
End Property
Public Property Get SubmissionID() As Long
    SubmissionID = m_SubmissionID
End Property
Public Property Get SubmissionRef() As String
    SubmissionRef = m_SubmissionRef
End Property
Public Property Get DateReceived() As Date
    DateReceived = m_DateReceived
End Property
Public Property Get PONumber() As String
    PONumber = m_PONumber
End Property
Public Property Get PaymentType() As PaymentTypeEnum
    PaymentType = m_PaymentType
End Property
Public Property Get PaymentTerms() As Long
    PaymentTerms = m_PaymentTerms
End Property
Public Property Get RequesterName() As String
    RequesterName = m_RequesterName
End Property
Public Property Get RequesterEmail() As String
    RequesterEmail = m_RequesterEmail
End Property
Public Property Get ServiceDate() As Date
    ServiceDate = m_ServiceDate
End Property
Public Property Get LineItems() As Collection
    Set LineItems = m_LineItems
End Property
Public Property Get TotalValue() As Currency
    Dim LI As clsSalesOrderLineItem
    Dim Total As Currency
    Total = 0
    For Each LI In m_LineItems
        Total = Total + LI.Total
    Next LI
    TotalValue = Total
End Property

' ================================================================
' Initialization
' ================================================================
Public Sub Initialize(Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    Set m_Logger = Logger
    Set m_LineItems = New Collection
End Sub

' ================================================================
' APIs
' ================================================================
Public Sub LoadFromData(Data As clsSalesOrderData)
    Debug.Assert Not Data Is Nothing

    AssertInitialized "LoadFromData"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_ID = 0
    m_CustomerName = Data.CustomerName
    m_CustomerID = Data.CustomerID
    m_SubmissionID = Data.SubmissionID
    m_SubmissionRef = Data.SubmissionRef
    m_DateReceived = Data.DateReceived
    m_PONumber = Data.PONumber
    m_PaymentType = Data.PaymentType
    m_PaymentTerms = Data.PaymentTerms
    m_RequesterName = Data.RequesterName
    m_RequesterEmail = Data.RequesterEmail
    m_ServiceDate = Data.ServiceDate

    Exit Sub
EH:
    HandleError "LoadFromData"
End Sub

Public Sub AddLineItem(LI As clsSalesOrderLineItem)
    Debug.Assert Not LI Is Nothing

    AssertInitialized "AddLineItem"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_LineItems.Add LI

    Exit Sub
EH:
    HandleError "AddLineItem"
End Sub

Public Function Validate() As Boolean
    AssertInitialized "Validate"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    Validate = True

    If m_CustomerID = 0 Then Validate = False
    If m_CustomerName = "" Then Validate = False
    If m_ServiceDate = 0 Then Validate = False
    If m_DateReceived = 0 Then Validate = False
    If m_PaymentType = PONumber And m_PONumber = "" Then Validate = False
    If m_PaymentTerms < 0 Then Validate = False

    Exit Function
EH:
    HandleError "Validate"
End Function

Friend Sub SetInternalID(NewID As Long)
    Debug.Assert m_ID = 0
    m_ID = NewID
End Sub

' ================================================================
' Private Helpers
' ================================================================
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsSalesOrder." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsSalesOrder." & MethodName, "The sales order must be initialized before use."
    End If
End Sub
