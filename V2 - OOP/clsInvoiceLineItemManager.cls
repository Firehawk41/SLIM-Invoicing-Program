VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceLineItemManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceLineItemManager
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-11
' Description: Manager class for line item objects
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsInvoicePriceRetriever, IInvoiceLineItem, clsInvoiceLineItemWriter
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_PricingRetriever As clsInvoicePriceRetriever
Private m_LineItems As Collection
Private m_LineItemWriter As clsInvoiceLineItemWriter
Public Property Get LineItems() As Collection
    Set LineItems = m_LineItems
End Property
Public Property Get TotalCost() As Long
    Dim LineItem As IInvoiceLineItem
    
    For Each LineItem In m_LineItems
        TotalCost = TotalCost + LineItem.TotalCost
    Next LineItem
End Property
Public Sub Initialize(Submissions As Collection, Logger As clsLoggingSystem, AccessDB As clsAccessDatabase, Factory As clsInvoiceFactory, PricingRetriever As clsInvoicePriceRetriever, SLIMDict As clsSLIMDictionaryManager)
    Const MINIMUM_CHARGE As Long = 150
    Dim MinCostLineItem As IInvoiceLineItem
    
    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    Set m_PricingRetriever = PricingRetriever
    Set m_LineItems = Factory.SetLineItemCollection(Submissions, AccessDB, PricingRetriever, SLIMDict)
    Call m_LineItemWriter.Initialize(Logger)
    
    ' If the total cost is less than then minimum
    If Me.TotalCost < MINIMUM_CHARGE Then
        Set MinCostLineItem = Factory.AddMinimumCostLineItem
        Call MinCostLineItem.FillOutMinCharge(MINIMUM_CHARGE, Me.TotalCost, m_LineItems(m_LineItems.Count).TextColour, m_Logger)
        m_LineItems.Add MinCostLineItem
    End If
    
End Sub

Public Sub FillInLineItemTable(Table As Word.Table)
    Dim LineItem As IInvoiceLineItem
    Dim LineNumber As Long
    
    Call m_Logger.LogMessage("FillInLineItemTable", LogLevelEnum.LogDebug, "Initiating filling in the Line Item Table")
    
    LineNumber = 2
    
    For Each LineItem In m_LineItems
        
        ' Add a table row if one is needed
        If Table.Rows.Count < LineNumber Then
            Table.Rows.Add
        End If
        
        ' Write the line item to the word table
        Call m_LineItemWriter.WriteLineItem(Table, LineNumber, LineItem.TextColour, LineItem.ColumnDictionary)

        ' Update line number to point to the end of the table (to handle split cells)
        LineNumber = Table.Rows.Count + 1
    Next LineItem
    
    Call m_Logger.LogMessage("FillInLineItemTable", LogLevelEnum.LogInfo, "Successfully filled in the Line Item Table")

End Sub

Private Sub Class_Initialize()
    Set m_LineItems = New Collection
    Set m_LineItemWriter = New clsInvoiceLineItemWriter
End Sub

Private Sub Class_Terminate()
    Dim LineItem As IInvoiceLineItem
    
    For Each LineItem In m_LineItems
        Set LineItem = Nothing
    Next LineItem
    Set m_LineItems = Nothing
    Set m_LineItemWriter = Nothing
End Sub
