VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalesOrderData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsSalesOrderData
' Author: JT
' Created: 2026-04-28
' Description: Value object carrying all data required to construct
'              a new clsSalesOrder via the creation path. Produced
'              by clsSalesOrderBuilder and consumed by
'              clsSalesOrderService.CreateFromData.
' DependsOn: modEnums
' ChangeLog:
'   - 1.0.0 - 2026-04-28 - Initial release
'   - 1.1.0 - 2026-05-19 - Added PaymentTerms (long)
'   - 1.2.0 - 2026-05-19 - Added SalesOrderID for rehydration path.
'                           Removed InvoiceEmail — owned by customer
'                           domain. DTO now serves both creation and
'                           rehydration paths.
'
' ================================================================
' LAYER: Value Object - no behaviour, no lifecycle, no dependencies.
'
' INTENT:
'   Makes the construction contract between clsSalesOrderBuilder and
'   clsSalesOrderService explicit and compiler-visible. Intentionally
'   a public-field struct. This class has no business rules of its
'   own and no lifecycle beyond the builder-to-service handoff.
'
'   Not used in the rehydration path. For loading a persisted sales
'   order from the database, see clsSalesOrder.LoadFromDict.
'
' Fields:
'   CustomerName    (String)
'   CustomerID      (Long)
'   SubmissionID    (Long)
'   SubmissionRef   (String)
'   DateReceived    (Date)
'   PONumber        (String)
'   PaymentType     (PaymentTypeEnum)
'   PaymentTerms    (Long)
'   RequesterName   (String)
'   RequesterEmail  (String)
'   ServiceDate     (Date)
'   LineItems       (Collection)    - of clsSalesOrderLineItem
' ================================================================

Public CustomerName As String
Public CustomerID As Long
Public SubmissionID As Long
Public SubmissionRef As String
Public DateReceived As Date
Public PONumber As String
Public PaymentType As PaymentTypeEnum
Public PaymentTerms As Long
Public RequesterName As String
Public RequesterEmail As String
Public ServiceDate As Date
Public LineItems As Collection
Public SalesOrderID As Long  ' 0 on creation path, populated by repo on rehydration
