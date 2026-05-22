VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomerRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsCustomerRepository
' Author: JT
' Created: 2026-04-10
' Description: Owns ALL SQL for the customers domain.
'              No business logic. No object graph traversal.
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsCustomer
' ChangeLog:
'   - 1.0.0 - 2026-04-10 - Initial release (replaces v1.0.0 placeholder)
'
' ================================================================
' LAYER: Repository - translates between objects and SQL.
'
' SelectAll returns fields in this column order (0-based):
'   0 ID  |  1 name  |  2 street_address  |  3 city  |  4 state
'   5 postal_code  |  6 country
' ================================================================

Private m_Logger   As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase

' ----------------------------------------------------------------
' Initialization
' ----------------------------------------------------------------
Public Sub Initialize(AccessDB As clsAccessDatabase, Logger As clsLoggingSystem)
    Debug.Assert Not AccessDB Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing        ' enforce single-init

    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
End Sub

' ----------------------------------------------------------------
' Bulk read  (called ONLY by clsCustomerCache.Build)
' ----------------------------------------------------------------
Public Function SelectAll() As Variant
    AssertInit "SelectAll"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAll = m_AccessDB.ExecuteQuery( _
        "SELECT ID, customer_name, street_address, city, state, postal_code, country " & _
        "FROM customers " & _
        "ORDER BY customer_name")
    Exit Function
EH:                                                             HandleError "SelectAll", "Failed"
End Function
' ----------------------------------------------------------------
' Form customer lookup  (called ONLY by clsCustomerCache)
' Returns 2D array: col 0 = form_customer, col 1 = form_address, col 2 = form_address2, col 3 = customer_id
' Returns Empty if table is empty
' ----------------------------------------------------------------
Public Function SelectAllFormCustomers() As Variant
    AssertInit "SelectAllFormCustomers"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllFormCustomers = m_AccessDB.ExecuteQuery( _
        "SELECT form_customer, form_address, form_address_2, customer_id " & _
        "FROM form_customers " & _
        "ORDER BY form_customer")
    Exit Function
EH:                                     HandleError "SelectAllFormCustomers", "Failed"
End Function
' ----------------------------------------------------------------
' Writes
' ----------------------------------------------------------------
' Inserts a new customer row. Returns the new autonumber ID.
'
Public Function Insert(C As clsCustomer) As Long
    AssertInit "Insert"
    Debug.Assert Not C Is Nothing
    Debug.Assert C.ID = 0               ' must be unsaved
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim SQL As String
    SQL = "INSERT INTO customers (customer_name, street_address, city, state, postal_code, country) " & _
          "VALUES (" & _
              SqlStr(C.Name) & ", " & _
              SqlStr(C.StreetAddress) & ", " & _
              SqlStr(C.City) & ", " & _
              SqlStr(C.State) & ", " & _
              SqlStr(C.PostalCode) & ", " & _
              SqlStr(C.Country) & ")"

    Insert = m_AccessDB.ExecuteNonQuery(SQL)
    Exit Function
EH:                                                             HandleError "Insert", "Failed on: " & C.Name
End Function
Public Sub InsertFormCustomer(Name As String, Address As String, Address2 As String, CustomerID As Long)
    AssertInit "InsertFormCustomer"
    Debug.Assert Len(Trim(Name)) > 0
    Debug.Assert CustomerID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_AccessDB.ExecuteNonQuery _
        "INSERT INTO form_customers (form_customer_name, form_address, form_address2, customer_id) " & _
        "VALUES (" & SqlStr(Name) & ", " & SqlStr(Address) & ", " & SqlStr(Address2) & ", " & CustomerID & ")"
    Exit Sub
EH:
    HandleError "InsertFormCustomer", "Name: " & Name & " | CustomerID: " & CustomerID
End Sub
Public Sub Update(C As clsCustomer)
    AssertInit "Update"
    Debug.Assert Not C Is Nothing
    Debug.Assert C.ID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim SQL As String
    SQL = "UPDATE customers SET " & _
          "    name           = " & SqlStr(C.Name) & ", " & _
          "    street_address = " & SqlStr(C.StreetAddress) & ", " & _
          "    city           = " & SqlStr(C.City) & ", " & _
          "    state          = " & SqlStr(C.State) & ", " & _
          "    postal_code    = " & SqlStr(C.PostalCode) & ", " & _
          "    country        = " & SqlStr(C.Country) & " " & _
          "WHERE ID = " & C.ID

    m_AccessDB.ExecuteNonQuery SQL
    Exit Sub
EH:                                                             HandleError "Update", "Failed on ID: " & C.ID
End Sub

Public Sub Delete(CustomerID As Long)
    AssertInit "Delete"
    Debug.Assert CustomerID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM customers WHERE ID = " & CustomerID
    Exit Sub
EH:                                                             HandleError "Delete", "Failed on ID: " & CustomerID
End Sub

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Function SqlStr(V As String) As String
    SqlStr = "'" & Replace(V, "'", "''") & "'"
End Function

Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsCustomerRepository." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Context As String)
    m_Logger.LogError "clsCustomerRepository." & Src, Err.Number, Err.Description & " | " & Context
End Sub
