VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalesOrderBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsSalesOrderBuilder
' Author: JT
' Created: 2026-03-13
' Description: Application-layer builder. Constructs a clsSalesOrder
'              from a clsTRSubmission by delegating line item
'              construction to clsSalesOrderLineItemBuilder, mapping
'              submission properties to clsSalesOrderData, and
'              passing the result to clsSalesOrderService for
'              persistence.
' DependsOn: clsLoggingSystem, clsSalesOrder, clsSalesOrderData,
'            clsSalesOrderLineItemBuilder, clsSalesOrderService,
'            clsCustomerService, clsTRSubmission, modInvoiceSystem
' ChangeLog:
'   - 1.0.0 - 2026-04-28 - Revised from clsInvoiceSalesOrderBuilder.
'                           Dict construction replaced with typed DTO.
'                           CustomerSvc retained for canonical name.
'                           PricingEngine removed — lives on line item
'                           builder. PaymentTerms from
'                           DEFAULT_PAYMENT_TERMS constant pending
'                           billing profile implementation.
'                           SubmissionDate removed — DateReceived is
'                           the only relevant date from lab perspective.
'
' ================================================================
' LAYER: Application — sits above the domain. Calls service,
'        never calls repository directly.
'
' PIPELINE:
'   clsTRSubmission
'         |
'   clsSalesOrderLineItemBuilder.BuildLineItems
'         | Collection of clsSalesOrderLineItem
'   BuildSalesOrderData
'         | clsSalesOrderData
'   clsSalesOrderService.CreateFromData
'         | clsSalesOrder (persisted, ID sealed)
'
' PaymentTerms placeholder:
'   Currently sourced from DEFAULT_PAYMENT_TERMS in modInvoiceSystem.
'   Pending implementation of a customer billing profile entity that
'   would own per-customer payment terms.
'
' CustomerName  (String) — snapshot of customer name at order time.
'                          Denormalised for the writer — avoids a
'                          service call in the output layer.
'                          CustomerID is the canonical relational link.
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_LineItemBuilder As clsSalesOrderLineItemBuilder
Private m_CustomerSvc As clsCustomerService
Private m_SalesOrderSvc As clsSalesOrderService
Private Const DEFAULT_PAYMENT_TERMS As Long = 30
' ================================================================
' Initialization
' ================================================================
Public Sub Initialize(LineItemBuilder As clsSalesOrderLineItemBuilder, _
                      CustomerSvc As clsCustomerService, _
                      SalesOrderSvc As clsSalesOrderService, _
                      Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert Not LineItemBuilder Is Nothing
    Debug.Assert Not CustomerSvc Is Nothing
    Debug.Assert Not SalesOrderSvc Is Nothing
    Debug.Assert m_Logger Is Nothing

    Set m_Logger = Logger
    Set m_LineItemBuilder = LineItemBuilder
    Set m_CustomerSvc = CustomerSvc
    Set m_SalesOrderSvc = SalesOrderSvc
End Sub

' ================================================================
' APIs
' ================================================================
Public Function BuildFromSubmission(Submission As clsTRSubmission) As clsSalesOrder
    Debug.Assert Not Submission Is Nothing
    AssertInitialized "BuildFromSubmission"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    ' Price first — DTOs carry PaymentTerms
    Dim PricedGroups As Collection
    Set PricedGroups = m_LineItemBuilder.PriceSubmission(Submission)

    ' Read PaymentTerms from first group
    Dim PaymentTerms As Long
    If PricedGroups.Count > 0 Then
        Dim FirstGroup As clsSalesOrderLineItemData
        Set FirstGroup = PricedGroups(1)
        PaymentTerms = FirstGroup.PaymentTerms
    Else
        PaymentTerms = DEFAULT_PAYMENT_TERMS
    End If

    ' Build entities from priced DTOs
    Dim LineItems As Collection
    Set LineItems = m_LineItemBuilder.BuildFromPricedGroups(PricedGroups)

    ' Build sales order DTO and delegate to service
    Dim Data As clsSalesOrderData
    Set Data = BuildSalesOrderData(Submission, LineItems, PaymentTerms)

    Set BuildFromSubmission = m_SalesOrderSvc.CreateFromData(Data)
    Exit Function
EH:
    HandleError "BuildFromSubmission", "Failed on: " & Submission.FileName
End Function

' ================================================================
' Private Helpers
' ================================================================
Private Function BuildSalesOrderData(Submission As clsTRSubmission, LineItems As Collection, PaymentTerms As Long) As clsSalesOrderData
    Dim Data As New clsSalesOrderData

    Data.CustomerName = m_CustomerSvc.GetNameByID(Submission.CustomerID)
    Data.CustomerID = Submission.CustomerID
    Data.SubmissionID = Submission.ID
    Data.SubmissionRef = Submission.FileName
    Data.DateReceived = Submission.DateReceived
    Data.PONumber = Submission.POInformation
    Data.PaymentType = IIf(Submission.CreditCardInformation <> "", _
                             PaymentTypeEnum.CreditCard, _
                             PaymentTypeEnum.PONumber)
    Data.PaymentTerms = PaymentTerms
    Data.RequesterName = Submission.CustomerContact
    Data.RequesterEmail = Submission.ResultsEmailMainFirst
    Data.ServiceDate = Submission.EffectiveServiceDate
    Set Data.LineItems = LineItems

    Set BuildSalesOrderData = Data
End Function

Private Sub HandleError(MethodName As String, Optional Context As String = "")
    m_Logger.LogError "clsSalesOrderBuilder." & MethodName, Err.Number, _
        Err.Description & IIf(Context <> "", " | " & Context, "")
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsSalesOrderBuilder." & MethodName, _
            "The builder must be initialized before use."
    End If
End Sub

