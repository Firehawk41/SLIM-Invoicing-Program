VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalesOrderLineItemData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsSalesOrderLineItemData
' Author: JT
' Created: 2026-03-27
' Description: Value object representing one priced group of samples
'              sharing the same analysis set, customer, and processing
'              time. Produced by clsSalesOrderPricingEngine, enriched
'              by clsSalesOrderLineItemBuilder, and consumed by
'              clsSalesOrderLineItem.LoadFromGroup. Maps 1:1 to one
'              line item on the final sales order.
' DependsOn: modEnums
' ChangeLog:
'   - 1.0.0 - 2026-03-27 - Initial release
'   - 1.1.0 - 2026-04-28 - Added Item, Description, QuotePriceID;
'                           removed Analyses (traceability runs
'                           through QuotePriceID); ProcessingTime
'                           typed to ProcessingTimeEnum
'   - 1.2.0 - 2026-04-28 - Added PaymentTerms. Sourced from
'                           clsQuote by pricing engine and carried
'                           through pipeline to clsSalesOrderBuilder.
'                           Avoids injecting clsQuoteService into
'                           the builder.
'
' ================================================================
' LAYER: Value Object - no behaviour, no lifecycle, no dependencies.
'
' INTENT:
'   Makes the contract between clsSalesOrderPricingEngine and
'   clsSalesOrderLineItemBuilder explicit and compiler-visible.
'   Intentionally a public-field struct with no business rules or
'   lifecycle beyond the pipeline handoff.
'
'   Travels through two pipeline stages before consumption:
'     1. clsSalesOrderPricingEngine populates pricing fields
'        (Quantity, Price, QuotePriceID, PaymentTerms,
'         ProcessingTime, CustomerID, Samples, Analyses)
'     2. clsSalesOrderLineItemBuilder resolves analysis IDs to
'        names and populates Item and Description, then discards
'        Analyses
'     3. clsSalesOrderLineItem.LoadFromGroup consumes the completed
'        group and stores the permanent record
'
'   Analysis traceability is carried by QuotePriceID, not by
'   Analyses. Analyses is a working field used only during stages
'   1-2 and has no meaning after the builder handoff.
'
'   PaymentTerms is carried from clsQuote via the pricing engine.
'   The builder reads it from the first line item in the collection
'   — all groups for a submission share the same customer and
'   therefore the same payment terms.
'
' Fields:
'   StoredID        (Long)                 - rehydration path only. Set by
'                                           clsSalesOrderLineItemRepository
'                                           after SELECT. Always 0 on the
'                                           creation path. Consumed by the
'                                           service to seal the entity via
'                                           SetInternalID.
'   Quantity        (Long)                 - sample count in group
'   Samples         (Collection)           - sample names, for audit trail.
'                                           Pipeline working data only —
'                                           not persisted.
'   Analyses        (Scripting.Dictionary) - AnalysisID -> True.
'                                           Pipeline working data only —
'                                           not persisted. Discarded after
'                                           builder resolves Item and
'                                           Description.
'   ProcessingTime  (ProcessingTimeEnum)
'   CustomerID      (Long)
'   Price           (Currency)             - assigned by pricing engine
'   QuotePriceID    (Long)                 - permanent traceability link.
'                                           Survives to the entity and DB.
'   PaymentTerms    (Long)                 - sourced from clsQuote by
'                                           pricing engine. Carried to
'                                           clsSalesOrderBuilder via first
'                                           line item in collection. Not
'                                           persisted on line item — consumed
'                                           by builder only.
'   Item            (String)               - resolved display name.
'                                           Set by clsSalesOrderLineItemBuilder
'                                           after analysis ID resolution.
'                                           Empty on creation path until
'                                           builder stage completes.
'   Description     (String)               - resolved display description.
'                                           Set by clsSalesOrderLineItemBuilder
'                                           after analysis ID resolution.
'                                           Empty on creation path until
'                                           builder stage completes.
' ================================================================

Public StoredID As Long
Public Quantity As Long
Public Samples As Collection
Public Analyses As Scripting.Dictionary
Public ProcessingTime As ProcessingTimeEnum
Public CustomerID As Long
Public Price As Currency
Public QuotePriceID As Long
Public PaymentTerms As Long
Public Item As String
Public Description As String
