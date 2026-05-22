VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IInvoiceLineItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: IInvoiceLineItem
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-11
' Description: Interface for Invoice Line Items
' DependsOn: clsTestingRequestSample, clsLoggingSystem, clsAccessDatabase, clsInvoicePriceRetriever, clsSLIMDictionaryManager, clsInvoiceLineItemCustomerBSummary, clsInvoiceLineItemCustomerASummary, clsInvoiceLineItemRegIndividual
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Public Property Get ColumnDictionary() As Scripting.Dictionary
    
End Property
Public Property Get BoldDictionary() As Scripting.Dictionary

End Property
Public Property Get TextColour() As Long

End Property
Public Property Get TotalCost() As Long

End Property
Public Sub Initialize(GroupedSamples As Collection, TextColour As Long, Logger As clsLoggingSystem, AccessDB As clsAccessDatabase, PriceRetriever As clsInvoicePriceRetriever, SLIMDict As clsSLIMDictionaryManager)

End Sub
Public Sub FillOutMinCharge(MinimumCharge As Long, TotalCost As Long, TextColour As Long, Logger As clsLoggingSystem)

End Sub
