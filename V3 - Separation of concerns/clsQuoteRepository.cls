VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuoteRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsQuoteRepository
' Author: JT
' Created: 2026-04-28
' Description: Owns ALL SQL for the quote domain. No business
'              logic. No object graph construction — that is the
'              cache's job. Returns raw 2D arrays from
'              ExecuteQuery for all reads. Reads properties
'              directly from entities for all writes.
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsQuote,
'            clsQuoteLineItem, clsQuotePrice
' ChangeLog:
'   - 1.0.0 - 2026-04-28 - Initial release. Replaces
'                           clsQuoteLoader prototype.
'   - 1.0.1 - 2026-05-19 - Fixed DeleteLineItemsByQuoteID child
'                           FK column for default tables
'
' ================================================================
' LAYER: Repository — translates between objects and SQL.
'        Returns raw 2D arrays for reads — cache builds objects.
'        Reads properties directly for writes — no ToDict.
'        All multi-table writes are transactional.
'
' READ COLUMN ORDERS (0-based) — consumed by clsQuoteCache.Build
' ----------------------------------------------------------------
' SelectAllQuotes / SelectAllDefaultQuotes:
'   0 ID              1 effective_date    2 expiry_date
'   3 payment_type    4 payment_value     5 valid
'   6 created_by      7 created_date      8 modified_date
'   9 comments
'
' SelectAllLineItems (quote_line_items):
'   0 ID              1 quote_id          2 request_type_id
'   3 chemical_id     4 bulk_min          5 valid
'
' SelectAllDefaultLineItems (default_quote_line_items):
'   0 ID              1 default_quote_id  2 request_type_id
'   3 chemical_id     4 bulk_min          5 valid
'
' SelectAllPrices / SelectAllDefaultPrices:
'   0 ID              1 quote_line_item_id   2 processing_time_id
'   3 regular_price   4 bulk_price           5 valid
'
' SelectAllQuoteAnalyses (quote_analysis):
'   0 ID              1 quote_line_item_id   2 analysis_id
'
' SelectAllDefaultQuoteAnalyses (default_quote_analysis):
'   0 ID              1 default_quote_line_item_id
'   2 analysis_id
'
' SelectAllQuoteElements (quote_line_item_element):
'   0 ID              1 quote_line_item_id   2 element_id
'
' SelectAllDefaultQuoteElements (default_quote_line_item_element):
'   0 ID              1 default_quote_line_item_id
'   2 element_id
'
' SelectAllQuoteCustomers (quote_customer):
'   0 ID              1 quote_id            2 customer_id
'
' WRITE HELPERS (private):
'   InsertLineItems   — iterates line items, delegates to
'                       InsertLineItem, InsertPrices,
'                       InsertAnalyses, InsertElements
'   InsertLineItem    — inserts one line item, seals ID
'   InsertPrices      — inserts all prices for a line item
'   InsertAnalyses    — inserts all analysis links
'   InsertElements    — inserts all element links
'   InsertCustomers   — inserts all customer links
'   DeleteLineItemsByQuoteID — cascades child deletes before
'                              removing line items
'   DeleteQuoteCustomers     — removes quote_customer rows
' ================================================================

Private m_Logger   As clsLoggingSystem
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
' Reads — Customer Quotes
' ================================================================
Public Function SelectAllQuotes() As Variant
    AssertInit "SelectAllQuotes"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllQuotes = SelectFromTable( _
        "quotes", _
        "ID, effective_date, expiry_date, payment_type, payment_value, " & _
        "valid, created_by, created_date, modified_date, comments")
    Exit Function
EH:
    HandleError "SelectAllQuotes"
End Function

Public Function SelectAllLineItems() As Variant
    AssertInit "SelectAllLineItems"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllLineItems = SelectFromTable( _
        "quote_line_items", _
        "ID, quote_id, request_type_id, chemical_id, bulk_min, valid")
    Exit Function
EH:
    HandleError "SelectAllLineItems"
End Function

Public Function SelectAllPrices() As Variant
    AssertInit "SelectAllPrices"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllPrices = SelectFromTable( _
        "quote_prices", _
        "ID, quote_line_item_id, processing_time_id, regular_price, bulk_price, valid")
    Exit Function
EH:
    HandleError "SelectAllPrices"
End Function

Public Function SelectAllQuoteAnalyses() As Variant
    AssertInit "SelectAllQuoteAnalyses"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllQuoteAnalyses = SelectFromTable( _
        "quote_analysis", _
        "ID, quote_line_item_id, analysis_id")
    Exit Function
EH:
    HandleError "SelectAllQuoteAnalyses"
End Function

Public Function SelectAllQuoteElements() As Variant
    AssertInit "SelectAllQuoteElements"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllQuoteElements = SelectFromTable( _
        "quote_line_item_element", _
        "ID, quote_line_item_id, element_id")
    Exit Function
EH:
    HandleError "SelectAllQuoteElements"
End Function

Public Function SelectAllQuoteCustomers() As Variant
    AssertInit "SelectAllQuoteCustomers"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllQuoteCustomers = SelectFromTable( _
        "quote_customer", _
        "ID, quote_id, customer_id")
    Exit Function
EH:
    HandleError "SelectAllQuoteCustomers"
End Function

' ================================================================
' Reads — Default Quotes
' ================================================================
Public Function SelectAllDefaultQuotes() As Variant
    AssertInit "SelectAllDefaultQuotes"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllDefaultQuotes = SelectFromTable( _
        "default_quotes", _
        "ID, effective_date, expiry_date, payment_type, payment_value, " & _
        "valid, created_by, created_date, modified_date, comments")
    Exit Function
EH:
    HandleError "SelectAllDefaultQuotes"
End Function

Public Function SelectAllDefaultLineItems() As Variant
    AssertInit "SelectAllDefaultLineItems"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllDefaultLineItems = SelectFromTable( _
        "default_quote_line_items", _
        "ID, default_quote_id, request_type_id, chemical_id, bulk_min, valid")
    Exit Function
EH:
    HandleError "SelectAllDefaultLineItems"
End Function

Public Function SelectAllDefaultPrices() As Variant
    AssertInit "SelectAllDefaultPrices"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllDefaultPrices = SelectFromTable( _
        "default_quote_prices", _
        "ID, default_quote_line_item_id, processing_time_id, " & _
        "regular_price, bulk_price, valid")
    Exit Function
EH:
    HandleError "SelectAllDefaultPrices"
End Function

Public Function SelectAllDefaultQuoteAnalyses() As Variant
    AssertInit "SelectAllDefaultQuoteAnalyses"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllDefaultQuoteAnalyses = SelectFromTable( _
        "default_quote_analysis", _
        "ID, default_quote_line_item_id, analysis_id")
    Exit Function
EH:
    HandleError "SelectAllDefaultQuoteAnalyses"
End Function

Public Function SelectAllDefaultQuoteElements() As Variant
    AssertInit "SelectAllDefaultQuoteElements"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    SelectAllDefaultQuoteElements = SelectFromTable( _
        "default_quote_line_item_element", _
        "ID, default_quote_line_item_id, element_id")
    Exit Function
EH:
    HandleError "SelectAllDefaultQuoteElements"
End Function

' ================================================================
' Writes — Quote
' ================================================================
Public Function InsertQuote(Q As clsQuote) As Long
    AssertInit "InsertQuote"
    Debug.Assert Not Q Is Nothing
    Debug.Assert Q.ID = 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim NewID As Long


    Dim SQL As String
    SQL = "INSERT INTO quotes " & _
          "    (effective_date, expiry_date, payment_type, " & _
          "     payment_value, valid, created_by, " & _
          "     created_date, modified_date, comments) " & _
          "VALUES (" & _
          "    " & SqlDate(Q.EffectiveDate) & ", " & _
          "    " & SqlDate(Q.ExpireDate) & ", " & _
          "    " & Q.PaymentType & ", " & _
          "    " & Q.PaymentTerms & ", " & _
          "    " & SqlBool(Q.IsValid) & ", " & _
          "    " & SqlStr(Q.CreatedBy) & ", " & _
          "    " & SqlDate(Q.CreatedDate) & ", " & _
          "    " & SqlDate(Q.ModifiedDate) & ", " & _
          "    " & SqlStr(Q.Comments) & ")"

        NewID = m_AccessDB.ExecuteNonQuery(SQL, True)
        Q.SetInternalID NewID

        InsertLineItems Q, NewID, False
        InsertCustomers Q, NewID

    m_AccessDB.CommitTransaction

    InsertQuote = NewID
    Exit Function
EH:
    m_AccessDB.RollbackTransaction
    HandleError "InsertQuote", "Failed on: " & Q.CreatedBy
End Function

Public Sub UpdateQuote(Q As clsQuote)
    AssertInit "UpdateQuote"
    Debug.Assert Not Q Is Nothing
    Debug.Assert Q.ID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim SQL As String
    SQL = "UPDATE quotes SET " & _
          "    effective_date  = " & SqlDate(Q.EffectiveDate) & ", " & _
          "    expiry_date     = " & SqlDate(Q.ExpireDate) & ", " & _
          "    payment_type    = " & Q.PaymentType & ", " & _
          "    payment_value   = " & Q.PaymentTerms & ", " & _
          "    valid           = " & SqlBool(Q.IsValid) & ", " & _
          "    created_by      = " & SqlStr(Q.CreatedBy) & ", " & _
          "    created_date    = " & SqlDate(Q.CreatedDate) & ", " & _
          "    modified_date   = " & SqlDate(Q.ModifiedDate) & ", " & _
          "    comments        = " & SqlStr(Q.Comments) & " " & _
          "WHERE ID = " & Q.ID

    m_AccessDB.ExecuteNonQuery SQL
    Exit Sub
EH:
    HandleError "UpdateQuote", "Failed on ID: " & Q.ID
End Sub

Public Sub DeleteQuote(QuoteID As Long)
    AssertInit "DeleteQuote"
    Debug.Assert QuoteID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    DeleteLineItemsByQuoteID QuoteID, False
    DeleteQuoteCustomers QuoteID

    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM quotes WHERE ID = " & QuoteID, True

    m_AccessDB.CommitTransaction
    Exit Sub
EH:
    m_AccessDB.RollbackTransaction
    HandleError "DeleteQuote", "Failed on ID: " & QuoteID
End Sub

' ================================================================
' Writes — Default Quote
' ================================================================
Public Function InsertDefaultQuote(Q As clsQuote) As Long
    AssertInit "InsertDefaultQuote"
    Debug.Assert Not Q Is Nothing
    Debug.Assert Q.ID = 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim NewID As Long

    Dim SQL As String
    SQL = "INSERT INTO default_quotes " & _
          "    (effective_date, expiry_date, payment_type, " & _
          "     payment_value, valid, created_by, " & _
          "     created_date, modified_date, comments) " & _
          "VALUES (" & _
          "    " & SqlDate(Q.EffectiveDate) & ", " & _
          "    " & SqlDate(Q.ExpireDate) & ", " & _
          "    " & Q.PaymentType & ", " & _
          "    " & Q.PaymentTerms & ", " & _
          "    " & SqlBool(Q.IsValid) & ", " & _
          "    " & SqlStr(Q.CreatedBy) & ", " & _
          "    " & SqlDate(Q.CreatedDate) & ", " & _
          "    " & SqlDate(Q.ModifiedDate) & ", " & _
          "    " & SqlStr(Q.Comments) & ")"

    NewID = m_AccessDB.ExecuteNonQuery(SQL, True)
    Q.SetInternalID NewID

    InsertLineItems Q, NewID, True

    m_AccessDB.CommitTransaction

    InsertDefaultQuote = NewID
    Exit Function
EH:
    m_AccessDB.RollbackTransaction
    HandleError "InsertDefaultQuote", "Failed on: " & Q.CreatedBy
End Function

Public Sub DeleteDefaultQuote(QuoteID As Long)
    AssertInit "DeleteDefaultQuote"
    Debug.Assert QuoteID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH
    
    DeleteLineItemsByQuoteID QuoteID, True

    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM default_quotes WHERE ID = " & QuoteID, True

    m_AccessDB.CommitTransaction
    Exit Sub
EH:
    m_AccessDB.RollbackTransaction
    HandleError "DeleteDefaultQuote", "Failed on ID: " & QuoteID
End Sub

' ================================================================
' Private Write Helpers
' ================================================================
Private Sub InsertLineItems(Q As clsQuote, QuoteID As Long, IsDefault As Boolean)
    Dim LI As clsQuoteLineItem

    For Each LI In Q.LineItems
        Dim LineItemID As Long
        LineItemID = InsertLineItem(LI, QuoteID, IsDefault)
        InsertPrices LI, LineItemID, IsDefault
        InsertAnalyses LI, LineItemID, IsDefault
        InsertElements LI, LineItemID, IsDefault
    Next LI
End Sub

Private Function InsertLineItem(LI As clsQuoteLineItem, QuoteID As Long, _
                                IsDefault As Boolean) As Long
    Dim FKColumn As String
    Dim TableName As String

    If IsDefault Then
        TableName = "default_quote_line_items"
        FKColumn = "default_quote_id"
    Else
        TableName = "quote_line_items"
        FKColumn = "quote_id"
    End If

    Dim SQL As String
    SQL = "INSERT INTO " & TableName & " " & _
          "    (" & FKColumn & ", request_type_id, chemical_id, " & _
          "     bulk_min, valid) " & _
          "VALUES (" & _
          "    " & QuoteID & ", " & _
          "    " & LI.RequestType & ", " & _
          "    " & SqlNullable(LI.ChemicalID) & ", " & _
          "    " & LI.BulkPriceMin & ", " & _
          "    " & SqlBool(LI.IsValid) & ")"

    Dim NewID As Long
    NewID = m_AccessDB.ExecuteNonQuery(SQL)
    LI.SetInternalID NewID

    InsertLineItem = NewID
End Function

Private Sub InsertPrices(LI As clsQuoteLineItem, LineItemID As Long, _
                         IsDefault As Boolean)
    Dim TableName As String
    Dim FKColumn As String

    If IsDefault Then
        TableName = "default_quote_prices"
        FKColumn = "default_quote_line_item_id"
    Else
        TableName = "quote_prices"
        FKColumn = "quote_line_item_id"
    End If

    Dim Key As Variant
    For Each Key In LI.AllPrices
        Dim P As clsQuotePrice
        Set P = LI.AllPrices(Key)

        Dim SQL As String
        SQL = "INSERT INTO " & TableName & " " & _
              "    (" & FKColumn & ", processing_time_id, " & _
              "     regular_price, bulk_price, valid) " & _
              "VALUES (" & _
              "    " & LineItemID & ", " & _
              "    " & P.ProcessingTime & ", " & _
              "    " & P.RegularPrice & ", " & _
              "    " & P.BulkPrice & ", " & _
              "    " & SqlBool(P.IsValid) & ")"

        Dim NewID As Long
        NewID = m_AccessDB.ExecuteNonQuery(SQL)
        P.SetInternalID NewID
    Next Key
End Sub

Private Sub InsertAnalyses(LI As clsQuoteLineItem, LineItemID As Long, _
                           IsDefault As Boolean)
    Dim TableName As String
    Dim FKColumn As String

    If IsDefault Then
        TableName = "default_quote_analysis"
        FKColumn = "default_quote_line_item_id"
    Else
        TableName = "quote_analysis"
        FKColumn = "quote_line_item_id"
    End If

    Dim AnalysisID As Variant
    For Each AnalysisID In LI.AnalysisIDs
        m_AccessDB.ExecuteNonQuery _
            "INSERT INTO " & TableName & " (" & FKColumn & ", analysis_id) " & _
            "VALUES (" & LineItemID & ", " & AnalysisID & ")"
    Next AnalysisID
End Sub

Private Sub InsertElements(LI As clsQuoteLineItem, LineItemID As Long, _
                           IsDefault As Boolean)
    Dim TableName As String
    Dim FKColumn As String

    If IsDefault Then
        TableName = "default_quote_line_item_element"
        FKColumn = "default_quote_line_item_id"
    Else
        TableName = "quote_line_item_element"
        FKColumn = "quote_line_item_id"
    End If

    Dim ElementID As Variant
    For Each ElementID In LI.ElementIDs
        m_AccessDB.ExecuteNonQuery _
            "INSERT INTO " & TableName & " (" & FKColumn & ", element_id) " & _
            "VALUES (" & LineItemID & ", " & ElementID & ")"
    Next ElementID
End Sub

Private Sub InsertCustomers(Q As clsQuote, QuoteID As Long)
    Dim CustomerID As Variant
    For Each CustomerID In Q.CustomerIDs
        m_AccessDB.ExecuteNonQuery _
            "INSERT INTO quote_customer (quote_id, customer_id) " & _
            "VALUES (" & QuoteID & ", " & CustomerID & ")"
    Next CustomerID
End Sub

Private Sub DeleteLineItemsByQuoteID(QuoteID As Long, IsDefault As Boolean)
    Dim LineItemTable As String
    Dim PriceTable As String
    Dim AnalysisTable As String
    Dim ElementTable As String
    Dim FKColumn As String
    Dim ChildFKColumn As String

    If IsDefault Then
        LineItemTable = "default_quote_line_items"
        PriceTable = "default_quote_prices"
        AnalysisTable = "default_quote_analysis"
        ElementTable = "default_quote_line_item_element"
        FKColumn = "default_quote_id"
        ChildFKColumn = "default_quote_line_item_id"
    Else
        LineItemTable = "quote_line_items"
        PriceTable = "quote_prices"
        AnalysisTable = "quote_analysis"
        ElementTable = "quote_line_item_element"
        FKColumn = "quote_id"
        ChildFKColumn = "quote_line_item_id"
    End If

    Dim SubQuery As String
    SubQuery = "(SELECT ID FROM " & LineItemTable & _
               " WHERE " & FKColumn & " = " & QuoteID & ")"

    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM " & PriceTable & _
        " WHERE " & ChildFKColumn & " IN " & SubQuery

    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM " & AnalysisTable & _
        " WHERE " & ChildFKColumn & " IN " & SubQuery

    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM " & ElementTable & _
        " WHERE " & ChildFKColumn & " IN " & SubQuery

    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM " & LineItemTable & _
        " WHERE " & FKColumn & " = " & QuoteID
End Sub

Private Sub DeleteQuoteCustomers(QuoteID As Long)
    m_AccessDB.ExecuteNonQuery _
        "DELETE FROM quote_customer WHERE quote_id = " & QuoteID
End Sub

' ================================================================
' Private SQL Helpers
' ================================================================
Private Function SelectFromTable(TableName As String, _
                                 Columns As String) As Variant
    SelectFromTable = m_AccessDB.ExecuteQuery( _
        "SELECT " & Columns & " FROM " & TableName & " ORDER BY ID")
End Function

Private Function SqlStr(V As String) As String
    SqlStr = "'" & Replace(V, "'", "''") & "'"
End Function

Private Function SqlDate(V As Date) As String
    SqlDate = "#" & Format(V, "MM/DD/YYYY") & "#"
End Function

Private Function SqlBool(V As Boolean) As String
    SqlBool = IIf(V, "True", "False")
End Function

Private Function SqlNullable(V As Variant) As String
    If IsNull(V) Then
        SqlNullable = "NULL"
    Else
        SqlNullable = CStr(CLng(V))
    End If
End Function

' ================================================================
' Private Helpers
' ================================================================
Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsQuoteRepository." & Src, _
            "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Optional Context As String = "")
    m_Logger.LogError "clsQuoteRepository." & Src, Err.Number, _
        Err.Description & IIf(Context <> "", " | " & Context, "")
End Sub

