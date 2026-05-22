VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalesOrderLineItemRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsSalesOrderLineItemRepository
' Author: JT
' Created: 2026-04-28
' Description: Owns ALL SQL for the sales_order_items table.
'              Called internally by clsSalesOrderRepository —
'              never called directly by the service.
'              No business logic.
' DependsOn: clsLoggingSystem, clsAccessDatabase,
'            clsSalesOrderLineItem
' ChangeLog:
'   - 1.0.0 - 2026-04-28 - Initial release
'
' ================================================================
' LAYER: Repository - translates between objects and SQL.
'
' NOTE: This repository does not own transactions. Transaction
'       management is the responsibility of clsSalesOrderRepository
'       which calls this class within its own transaction scope.
'
' SelectBySalesOrderID returns fields in this column order (0-based):
'   0  sales_order_item_id
'   1  sales_order_id
'   2  item
'   3  description
'   4  quantity
'   5  rate
'   6  quote_price_id
'   7  processing_time_id
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase

' ================================================================
' Initialization
' ================================================================
Public Sub Initialize(AccessDB As clsAccessDatabase, Logger As clsLoggingSystem)
    Debug.Assert Not AccessDB Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
End Sub

' ================================================================
' Reads
' ================================================================
Public Function SelectBySalesOrderID(SalesOrderID As Long) As Collection
    AssertInit "SelectBySalesOrderID"
    Debug.Assert SalesOrderID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim Rows As Variant
    Rows = m_AccessDB.ExecuteQuery( _
        "SELECT sales_order_item_id, sales_order_id, item, " & _
        "       description, quantity, rate, quote_price_id, " & _
        "       processing_time_id " & _
        "FROM   sales_order_items " & _
        "WHERE  sales_order_id = " & SalesOrderID & " " & _
        "ORDER BY sales_order_item_id")

    Dim Results As New Collection
    If Not IsEmpty(Rows) Then
        Dim i As Long
        For i = 0 To UBound(Rows, 2)
            Results.Add RowToGroup(Rows, i)
        Next i
    End If

    Set SelectBySalesOrderID = Results
    Exit Function
EH:
    HandleError "SelectBySalesOrderID", "Failed on ID: " & SalesOrderID
End Function

' ================================================================
' Writes
' ================================================================
Public Function Insert(LI As clsSalesOrderLineItem, SalesOrderID As Long) As Long
    AssertInit "Insert"
    Debug.Assert Not LI Is Nothing
    Debug.Assert LI.ID = 0
    Debug.Assert SalesOrderID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim SQL As String
    SQL = "INSERT INTO sales_order_items " & _
          "    (sales_order_id, item, description, " & _
          "     quantity, rate, quote_price_id, processing_time_id) " & _
          "VALUES (" & _
          "    " & SalesOrderID & ", " & _
          "    " & SqlStr(LI.Item) & ", " & _
          "    " & SqlStr(LI.Description) & ", " & _
          "    " & LI.Quantity & ", " & _
          "    " & LI.Rate & ", " & _
          "    " & LI.QuotePriceID & ", " & _
          "    " & LI.ProcessingTime & ")"

    Dim NewID As Long
    NewID = m_AccessDB.ExecuteNonQuery(SQL)
    LI.SetInternalID NewID

    Insert = NewID
    Exit Function
EH:
    HandleError "Insert", "Failed on: " & LI.Item
End Function

Public Sub DeleteBySalesOrderID(SalesOrderID As Long)
    AssertInit "DeleteBySalesOrderID"
    Debug.Assert SalesOrderID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM sales_order_items " & _
        "WHERE sales_order_id = " & SalesOrderID
    Exit Sub
EH:
    HandleError "DeleteBySalesOrderID", "Failed on ID: " & SalesOrderID
End Sub

Public Sub Delete(LineItemID As Long)
    AssertInit "Delete"
    Debug.Assert LineItemID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM sales_order_items " & _
        "WHERE sales_order_item_id = " & LineItemID
    Exit Sub
EH:
    HandleError "Delete", "Failed on ID: " & LineItemID
End Sub

' ================================================================
' Private Helpers
' ================================================================
Private Function RowToGroup(Rows As Variant, RowIndex As Long) As clsSalesOrderLineItemData
    Dim Data As New clsSalesOrderLineItemData
    Data.Item = modUtilities.NullToDefault(Rows(2, RowIndex), "")
    Data.Description = modUtilities.NullToDefault(Rows(3, RowIndex), "")
    Data.Quantity = modUtilities.NullToDefault(Rows(4, RowIndex), 0)
    Data.Price = modUtilities.NullToDefault(Rows(5, RowIndex), 0)
    Data.QuotePriceID = modUtilities.NullToDefault(Rows(6, RowIndex), 0)
    Data.ProcessingTime = modUtilities.NullToDefault(Rows(7, RowIndex), 0)
    Set RowToGroup = Data
End Function
Private Function SqlStr(V As String) As String
    SqlStr = "'" & Replace(V, "'", "''") & "'"
End Function

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsSalesOrderLineItemRepository." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Optional Context As String = "")
    m_Logger.LogError "clsSalesOrderLineItemRepository." & Src, Err.Number, _
        Err.Description & IIf(Context <> "", " | " & Context, "")
End Sub

