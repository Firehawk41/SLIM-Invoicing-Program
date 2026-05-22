VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalesOrderRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsSalesOrderRepository
' Author: JT
' Created: 2026-04-28
' Description: Owns ALL SQL for the sales_orders table.
'              Calls clsSalesOrderLineItemRepository within a
'              single transaction for full graph persistence.
'              No business logic. No object graph traversal.
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsSalesOrder,
'            clsSalesOrderLineItemRepository
' ChangeLog:
'   - 1.0.0 - 2026-04-28 - Initial release
'
' ================================================================
' LAYER: Repository - translates between objects and SQL.
'
' ASSUMED clsAccessDatabase interface:
'   .ExecuteQuery(SQL)    As Variant   SELECT -> 2D array(field,row)
'                                      or Empty if no rows
'   .ExecuteNonQuery(SQL)              INSERT / UPDATE / DELETE
'   .Executenonquery(SQL)   As Long      INSERT, returns autonumber ID
'   .BeginTransaction
'   .CommitTransaction
'   .RollbackTransaction
'
' SelectByID returns fields in this column order (0-based):
'   0  sales_order_id
'   1  customer_name
'   2  customer_id
'   3  submission_id
'   4  submission_ref
'   5  submission_date
'   6  date_received
'   7  po_number
'   8  payment_type
'   9  payment_value
'   10 requester_name
'   11 requester_email
'   12 service_date
'
' Transaction ownership:
'   Insert and Delete own the full transaction including line items.
'   clsSalesOrderLineItemRepository is called internally.
'   Update operates on header fields only — line items updated
'   separately via clsSalesOrderLineItemRepository.
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_LineItemRepo As clsSalesOrderLineItemRepository

' ================================================================
' Initialization
' ================================================================
Public Sub Initialize(AccessDB As clsAccessDatabase, LineItemRepo As clsSalesOrderLineItemRepository, Logger As clsLoggingSystem)
    Debug.Assert Not AccessDB Is Nothing
    Debug.Assert Not LineItemRepo Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    Set m_LineItemRepo = LineItemRepo
End Sub

' ================================================================
' Reads
' ================================================================
Public Function SelectByID(SalesOrderID As Long) As clsSalesOrderData
    AssertInit "SelectByID"
    Debug.Assert SalesOrderID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim Rows As Variant
    Rows = m_AccessDB.ExecuteQuery( _
        "SELECT ID, customer_name, customer_id, " & _
        "       submission_id, submission_ref, " & _
        "       date_received, po_number, payment_type, " & _
        "       payment_terms, requester_name, requester_email, " & _
        "       service_date " & _
        "FROM   sales_orders " & _
        "WHERE  ID = " & SalesOrderID)

    If IsEmpty(Rows) Or UBound(Rows) < 0 Then
        Set SelectByID = Nothing
        Exit Function
    End If

    Set SelectByID = RowToData(Rows, 0)
    Exit Function
EH:
    HandleError "SelectByID", "Failed on ID: " & SalesOrderID
End Function

Public Function SelectByDateRange(StartDate As Date, EndDate As Date) As Collection
    AssertInit "SelectByDateRange"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim Rows As Variant
    Rows = m_AccessDB.ExecuteQuery( _
        "SELECT sales_order_id, customer_name, customer_id, " & _
        "       submission_id, submission_ref, " & _
        "       date_received, po_number, payment_type, " & _
        "       payment_value, requester_name, requester_email, " & _
        "       service_date " & _
        "FROM   sales_orders " & _
        "WHERE  service_date BETWEEN " & SqlDate(StartDate) & _
        "       AND " & SqlDate(EndDate) & _
        "ORDER BY service_date")

    Dim Results As New Collection
    If Not IsEmpty(Rows) Then
        Dim i As Long
        For i = 0 To UBound(Rows, 2)
            Results.Add RowToData(Rows, i)
        Next i
    End If

    Set SelectByDateRange = Results
    Exit Function
EH:
    HandleError "SelectByDateRange", "Failed on range: " & StartDate & " to " & EndDate
End Function

' ================================================================
' Writes
' ================================================================

Public Function Insert(SO As clsSalesOrder) As Long
    AssertInit "Insert"
    Debug.Assert Not SO Is Nothing
    Debug.Assert SO.ID = 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim NewID As Long
    Dim SQL As String

    SQL = "INSERT INTO sales_orders " & _
          "    (customer_name, customer_id, submission_id, " & _
          "     submission_ref, date_received, " & _
          "     po_number, payment_type, payment_terms, " & _
          "     requester_name, requester_email, service_date) " & _
          "VALUES (" & _
          "    " & SqlStr(SO.CustomerName) & ", " & _
          "    " & SO.CustomerID & ", " & _
          "    " & SO.SubmissionID & ", " & _
          "    " & SqlStr(SO.SubmissionRef) & ", " & _
          "    " & SqlDate(SO.DateReceived) & ", " & _
          "    " & SqlStr(SO.PONumber) & ", " & _
          "    " & SO.PaymentType & ", " & _
          "    " & SO.PaymentTerms & ", " & _
          "    " & SqlStr(SO.RequesterName) & ", " & _
          "    " & SqlStr(SO.RequesterEmail) & ", " & _
          "    " & SqlDate(SO.ServiceDate) & ")"

    NewID = m_AccessDB.ExecuteNonQuery(SQL, True)
    SO.SetInternalID NewID

    Dim LI As clsSalesOrderLineItem
    For Each LI In SO.LineItems
        m_LineItemRepo.Insert LI, NewID
    Next LI

    m_AccessDB.CommitTransaction

    Insert = NewID
    Exit Function
EH:
    m_AccessDB.RollbackTransaction
    HandleError "Insert", "Failed on: " & SO.CustomerName
End Function

Public Sub Update(SO As clsSalesOrder)
    AssertInit "Update"
    Debug.Assert Not SO Is Nothing
    Debug.Assert SO.ID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim SQL As String
    SQL = "UPDATE sales_orders SET " & _
          "    customer_name   = " & SqlStr(SO.CustomerName) & ", " & _
          "    customer_id     = " & SO.CustomerID & ", " & _
          "    submission_id   = " & SO.SubmissionID & ", " & _
          "    submission_ref  = " & SqlStr(SO.SubmissionRef) & ", " & _
          "    date_received   = " & SqlDate(SO.DateReceived) & ", " & _
          "    po_number       = " & SqlStr(SO.PONumber) & ", " & _
          "    payment_type    = " & SO.PaymentType & ", " & _
            "    payment_terms   = " & SO.PaymentTerms & ", " & _
          "    requester_name  = " & SqlStr(SO.RequesterName) & ", " & _
          "    requester_email = " & SqlStr(SO.RequesterEmail) & ", " & _
          "    service_date    = " & SqlDate(SO.ServiceDate) & " " & _
          "WHERE sales_order_id = " & SO.ID

    m_AccessDB.ExecuteNonQuery SQL
    Exit Sub
EH:
    HandleError "Update", "Failed on ID: " & SO.ID
End Sub

Public Sub Delete(SalesOrderID As Long)
    AssertInit "Delete"
    Debug.Assert SalesOrderID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH


    m_LineItemRepo.DeleteBySalesOrderID SalesOrderID

    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM sales_orders WHERE sales_order_id = " & SalesOrderID, True

    m_AccessDB.CommitTransaction
    Exit Sub
EH:
    m_AccessDB.RollbackTransaction
    HandleError "Delete", "Failed on ID: " & SalesOrderID
End Sub

' ================================================================
' Private Helpers
' ================================================================
Private Function RowToData(Rows As Variant, RowIndex As Long) As clsSalesOrderData
    Dim D As New clsSalesOrderData
    D.SalesOrderID = modUtilities.NullToDefault(Rows(0, RowIndex), 0)
    D.CustomerName = modUtilities.NullToDefault(Rows(1, RowIndex), "")
    D.CustomerID = modUtilities.NullToDefault(Rows(2, RowIndex), 0)
    D.SubmissionID = modUtilities.NullToDefault(Rows(3, RowIndex), 0)
    D.SubmissionRef = modUtilities.NullToDefault(Rows(4, RowIndex), "")
    D.DateReceived = modUtilities.NullToDefault(Rows(5, RowIndex), 0)
    D.PONumber = modUtilities.NullToDefault(Rows(6, RowIndex), "")
    D.PaymentType = modUtilities.NullToDefault(Rows(7, RowIndex), 0)
    D.PaymentTerms = modUtilities.NullToDefault(Rows(8, RowIndex), 0)
    D.RequesterName = modUtilities.NullToDefault(Rows(9, RowIndex), "")
    D.RequesterEmail = modUtilities.NullToDefault(Rows(10, RowIndex), "")
    D.ServiceDate = modUtilities.NullToDefault(Rows(11, RowIndex), 0)
    Set RowToData = D
End Function

Private Function SqlStr(V As String) As String
    SqlStr = "'" & Replace(V, "'", "''") & "'"
End Function

Private Function SqlDate(V As Date) As String
    SqlDate = "#" & Format(V, "MM/DD/YYYY") & "#"
End Function

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsSalesOrderRepository." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Optional Context As String = "")
    m_Logger.LogError "clsSalesOrderRepository." & Src, Err.Number, _
        Err.Description & IIf(Context <> "", " | " & Context, "")
End Sub

