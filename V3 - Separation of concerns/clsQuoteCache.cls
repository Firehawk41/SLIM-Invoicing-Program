VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuoteCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsQuoteCache
' Author: JT
' Created: 2026-04-28
' Description: In-memory store for the full quote graph. Holds
'              customer quotes indexed by quote ID and customer ID,
'              and default quotes as a flat collection. Built once
'              from clsQuoteRepository at startup. No business
'              logic — matching logic lives in
'              clsSalesOrderPricingEngine.
' DependsOn: clsLoggingSystem, clsQuoteRepository, clsQuote,
'            clsQuoteLineItem, clsQuotePrice, modEnums
' ChangeLog:
'   - 1.0.0 - 2026-04-28 - Initial release. Replaces
'                           clsInvoicePricingCache.
'   - 1.0.1 - 2026-05-19 - GetDefaultQuotes returns defensive
'                           copy. Dictionary iteration variable
'                           renamed to Key in BuildQuoteIndex
'                           and BuildDefaultQuoteCollection.
'
' ================================================================
' LAYER: Cache — in-memory store only. No SQL. No business logic.
'
' INDEX STRUCTURE:
'   m_ByQuoteID      QuoteID (Long)     -> clsQuote
'   m_ByCustomerID   CustomerID (Long)  -> Collection of clsQuote
'   m_DefaultQuotes  flat Collection    -> clsQuote
'
' LIFECYCLE:
'   1. New clsQuoteCache
'   2. .Initialize(Repo, ChemicalSvc, Logger)
'   3. .Build()              <- loads full graph from repo
'
'   Build() is safe to call multiple times — each call clears and
'   rebuilds the full graph. Call after any write to refresh cache.
'
' OBJECT GRAPH CONSTRUCTION ORDER (all in Build):
'   1. Prices built and indexed by line item ID
'   2. Line items built, prices attached, indexed by ID
'   3. Analysis IDs attached to line items
'   4. Element IDs attached to line items
'   5. Quotes built, line items attached, indexed by ID
'   6. Customer IDs attached to quotes, m_ByCustomerID populated
'   Same sequence repeated for default quotes -> m_DefaultQuotes
' ================================================================

Private m_Logger         As clsLoggingSystem
Private m_Repo           As clsQuoteRepository
Private m_ChemicalSvc    As clsChemicalService
Private m_ByQuoteID      As Scripting.Dictionary
Private m_ByCustomerID   As Scripting.Dictionary
Private m_DefaultQuotes  As Collection

' ================================================================
' Initialization
' ================================================================
Public Sub Initialize(Repo As clsQuoteRepository, ChemicalSvc As clsChemicalService, Logger As clsLoggingSystem)
    Debug.Assert Not Repo Is Nothing
    Debug.Assert Not ChemicalSvc Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    Set m_Logger = Logger
    Set m_Repo = Repo
    Set m_ChemicalSvc = ChemicalSvc
End Sub

' ================================================================
' Build
' ================================================================
Public Sub Build()
    AssertInitialized "Build"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Set m_ByQuoteID = New Scripting.Dictionary
    Set m_ByCustomerID = New Scripting.Dictionary
    Set m_DefaultQuotes = New Collection

    BuildCustomerQuotes
    BuildDefaultQuotes

    m_Logger.LogMessage "clsQuoteCache.Build", LogLevelEnum.LogInfo, _
        "Built " & m_ByQuoteID.Count & " customer quotes, " & _
        m_DefaultQuotes.Count & " default quotes."

    Exit Sub
EH:
    HandleError "Build"
End Sub

' ================================================================
' Public Interface
' ================================================================
Public Function GetByQuoteID(QuoteID As Long) As clsQuote
    AssertInitialized "GetByQuoteID"

    If m_ByQuoteID.Exists(QuoteID) Then
        Set GetByQuoteID = m_ByQuoteID(QuoteID)
    Else
        Set GetByQuoteID = Nothing
    End If
End Function

Public Function GetByCustomerID(CustomerID As Long) As Collection
    AssertInitialized "GetByCustomerID"
    If m_ByCustomerID.Exists(CustomerID) Then
        Dim Copy As New Collection
        Dim Q As clsQuote
        For Each Q In m_ByCustomerID(CustomerID)
            Copy.Add Q
        Next Q
        Set GetByCustomerID = Copy
    Else
        Set GetByCustomerID = New Collection
    End If
End Function

Public Function GetDefaultQuotes() As Collection
    AssertInitialized "GetDefaultQuotes"
    Dim Copy As New Collection
    Dim Q As clsQuote
    For Each Q In m_DefaultQuotes
        Copy.Add Q
    Next Q
    Set GetDefaultQuotes = Copy
End Function

Public Function ExistsByCustomerID(CustomerID As Long) As Boolean
    AssertInitialized "ExistsByCustomerID"
    ExistsByCustomerID = m_ByCustomerID.Exists(CustomerID)
End Function

Public Function ExistsByQuoteID(QuoteID As Long) As Boolean
    AssertInitialized "ExistsByQuoteID"
    ExistsByQuoteID = m_ByQuoteID.Exists(QuoteID)
End Function
Public Function AllCustomerQuotes() As Collection
    AssertInitialized "AllCustomerQuotes"

    Dim Results As New Collection
    Dim Key As Variant
    For Each Key In m_ByQuoteID
        Results.Add m_ByQuoteID(Key)
    Next Key
    Set AllCustomerQuotes = Results
End Function
' ================================================================
' Private — Customer Quote Graph Construction
' ================================================================
Private Sub BuildCustomerQuotes()
    ' 1. Build prices indexed by line item ID
    Dim PriceIndex As Scripting.Dictionary
    Set PriceIndex = BuildPriceIndex( _
        m_Repo.SelectAllPrices(), False)

    ' 2. Build line items indexed by ID, attach prices
    Dim LineItemIndex As Scripting.Dictionary
    Set LineItemIndex = BuildLineItemIndex( _
        m_Repo.SelectAllLineItems(), PriceIndex, False)

    ' 3. Attach analysis IDs to line items
    AttachAnalysisIDs LineItemIndex, m_Repo.SelectAllQuoteAnalyses(), False

    ' 4. Attach element IDs to line items
    AttachElementIDs LineItemIndex, m_Repo.SelectAllQuoteElements(), False

    ' 5. Build quotes indexed by ID, attach line items
    BuildQuoteIndex m_Repo.SelectAllQuotes(), LineItemIndex

    ' 6. Attach customer IDs, build customer index
    AttachCustomerIDs m_Repo.SelectAllQuoteCustomers()
End Sub

' ================================================================
' Private — Default Quote Graph Construction
' ================================================================
Private Sub BuildDefaultQuotes()
    ' 1. Build prices indexed by line item ID
    Dim PriceIndex As Scripting.Dictionary
    Set PriceIndex = BuildPriceIndex( _
        m_Repo.SelectAllDefaultPrices(), True)

    ' 2. Build line items indexed by ID, attach prices
    Dim LineItemIndex As Scripting.Dictionary
    Set LineItemIndex = BuildLineItemIndex( _
        m_Repo.SelectAllDefaultLineItems(), PriceIndex, True)

    ' 3. Attach analysis IDs
    AttachAnalysisIDs LineItemIndex, _
        m_Repo.SelectAllDefaultQuoteAnalyses(), True

    ' 4. Attach element IDs
    AttachElementIDs LineItemIndex, _
        m_Repo.SelectAllDefaultQuoteElements(), True

    ' 5. Build quotes, attach line items, store in m_DefaultQuotes
    BuildDefaultQuoteCollection m_Repo.SelectAllDefaultQuotes(), LineItemIndex
End Sub

' ================================================================
' Private — Object Construction Helpers
' ================================================================
Private Function BuildPriceIndex(Rows As Variant, _
                                 IsDefault As Boolean) As Scripting.Dictionary
    ' Returns Dictionary: LineItemID -> Collection of clsQuotePrice
    Dim Index As New Scripting.Dictionary
    Dim P As clsQuotePrice

    If IsEmpty(Rows) Or UBound(Rows) < 0 Then
        Set BuildPriceIndex = Index
        Exit Function
    End If

    Dim i As Long
    For i = 0 To UBound(Rows, 2)
        Set P = New clsQuotePrice
        P.Initialize m_Logger

        Dim D As New Scripting.Dictionary
        D("id") = modUtilities.NullToDefault(Rows(0, i), 0)
        D("quote_line_item_id") = modUtilities.NullToDefault(Rows(1, i), 0)
        D("processing_time_id") = modUtilities.NullToDefault(Rows(2, i), 0)
        D("regular_price") = modUtilities.NullToDefault(Rows(3, i), 0)
        D("bulk_price") = modUtilities.NullToDefault(Rows(4, i), -1)
        D("valid") = modUtilities.NullToDefault(Rows(5, i), False)

        P.LoadFromDict D

        Dim LineItemID As Long
        LineItemID = CLng(D("quote_line_item_id"))

        If Not Index.Exists(LineItemID) Then
            Set Index(LineItemID) = New Collection
        End If
        Index(LineItemID).Add P
    Next i

    Set BuildPriceIndex = Index
End Function

Private Function BuildLineItemIndex(Rows As Variant, _
                                    PriceIndex As Scripting.Dictionary, _
                                    IsDefault As Boolean) As Scripting.Dictionary
    ' Returns Dictionary: LineItemID -> clsQuoteLineItem
    Dim Index As New Scripting.Dictionary
    Dim LI As clsQuoteLineItem
    Dim D As Scripting.Dictionary
    
    If IsEmpty(Rows) Or UBound(Rows) < 0 Then
        Set BuildLineItemIndex = Index
        Exit Function
    End If

    ' Column 1 is quote_id for customer, default_quote_id for default
    ' Both are in column position 1 — same offset, different FK name

    Dim i As Long
    For i = 0 To UBound(Rows, 2)
        Set LI = New clsQuoteLineItem
        LI.Initialize m_ChemicalSvc, m_Logger

        Set D = New Scripting.Dictionary
        D("id") = modUtilities.NullToDefault(Rows(0, i), 0)
        D("quote_id") = modUtilities.NullToDefault(Rows(1, i), 0)
        D("request_type_id") = modUtilities.NullToDefault(Rows(2, i), 0)
        D("chemical_id") = Rows(3, i)
        D("bulk_min") = modUtilities.NullToDefault(Rows(4, i), 0)
        D("valid") = modUtilities.NullToDefault(Rows(5, i), False)

        LI.LoadFromDict D

        ' Attach prices
        Dim LID As Long
        LID = CLng(D("id"))

        If PriceIndex.Exists(LID) Then
            Dim P As clsQuotePrice
            For Each P In PriceIndex(LID)
                LI.UpsertPrice P
            Next P
        End If

        Set Index(LID) = LI
    Next i

    Set BuildLineItemIndex = Index
End Function

Private Sub AttachAnalysisIDs(LineItemIndex As Scripting.Dictionary, _
                              Rows As Variant, _
                              IsDefault As Boolean)
    If IsEmpty(Rows) Or UBound(Rows) < 0 Then Exit Sub

    ' Col 0 = ID, Col 1 = line_item_id, Col 2 = analysis_id
    Dim i As Long
    For i = 0 To UBound(Rows, 2)
        Dim LineItemID As Long
        LineItemID = CLng(modUtilities.NullToDefault(Rows(1, i), 0))

        If LineItemIndex.Exists(LineItemID) Then
            Dim LI As clsQuoteLineItem
            Set LI = LineItemIndex(LineItemID)
            LI.AddAnalysisID CLng(modUtilities.NullToDefault(Rows(2, i), 0))
        Else
            m_Logger.LogMessage "clsQuoteCache.AttachAnalysisIDs", _
                LogLevelEnum.LogWarning, _
                "LineItem ID not found: " & LineItemID
        End If
    Next i
End Sub

Private Sub AttachElementIDs(LineItemIndex As Scripting.Dictionary, _
                             Rows As Variant, _
                             IsDefault As Boolean)
    If IsEmpty(Rows) Or UBound(Rows) < 0 Then Exit Sub

    ' Col 0 = ID, Col 1 = line_item_id, Col 2 = element_id
    Dim i As Long
    For i = 0 To UBound(Rows, 2)
        Dim LineItemID As Long
        LineItemID = CLng(modUtilities.NullToDefault(Rows(1, i), 0))

        If LineItemIndex.Exists(LineItemID) Then
            Dim LI As clsQuoteLineItem
            Set LI = LineItemIndex(LineItemID)
            LI.AddElementID CLng(modUtilities.NullToDefault(Rows(2, i), 0))
        Else
            m_Logger.LogMessage "clsQuoteCache.AttachElementIDs", _
                LogLevelEnum.LogWarning, _
                "LineItem ID not found: " & LineItemID
        End If
    Next i
End Sub

Private Sub BuildQuoteIndex(Rows As Variant, _
                            LineItemIndex As Scripting.Dictionary)
    If IsEmpty(Rows) Or UBound(Rows) < 0 Then Exit Sub
    
    Dim Q As clsQuote
    Dim i As Long
    Dim D As Scripting.Dictionary
    
    For i = 0 To UBound(Rows, 2)
        Set Q = New clsQuote
        Q.Initialize m_Logger

        Set D = New Scripting.Dictionary
        D("ID") = modUtilities.NullToDefault(Rows(0, i), 0)
        D("effective_date") = modUtilities.NullToDefault(Rows(1, i), 0)
        D("expiry_date") = modUtilities.NullToDefault(Rows(2, i), 0)
        D("payment_type") = modUtilities.NullToDefault(Rows(3, i), 0)
        D("payment_value") = modUtilities.NullToDefault(Rows(4, i), 0)
        D("valid") = modUtilities.NullToDefault(Rows(5, i), False)
        D("created_by") = modUtilities.NullToDefault(Rows(6, i), "")
        D("created_date") = modUtilities.NullToDefault(Rows(7, i), 0)
        D("modified_date") = modUtilities.NullToDefault(Rows(8, i), 0)
        D("comments") = modUtilities.NullToDefault(Rows(9, i), "")
        D("is_default_quote") = False

        Q.LoadFromDict D

        ' Attach line items
        Dim Key As Variant
        For Each Key In LineItemIndex
            If CLng(LineItemIndex(Key).QuoteID) = Q.ID Then
                Q.UpsertLineItem LineItemIndex(Key)
            End If
        Next Key
        
        Set m_ByQuoteID(Q.ID) = Q
    Next i
End Sub

Private Sub BuildDefaultQuoteCollection(Rows As Variant, _
                                        LineItemIndex As Scripting.Dictionary)
    If IsEmpty(Rows) Or UBound(Rows) < 0 Then Exit Sub
    
    Dim D As New Scripting.Dictionary
    Dim i As Long
    
    For i = 0 To UBound(Rows, 2)
        Dim Q As New clsQuote
        Q.Initialize m_Logger

        Set D = New Scripting.Dictionary
        D("ID") = modUtilities.NullToDefault(Rows(0, i), 0)
        D("effective_date") = modUtilities.NullToDefault(Rows(1, i), 0)
        D("expiry_date") = modUtilities.NullToDefault(Rows(2, i), 0)
        D("payment_type") = modUtilities.NullToDefault(Rows(3, i), 0)
        D("payment_value") = modUtilities.NullToDefault(Rows(4, i), 0)
        D("valid") = modUtilities.NullToDefault(Rows(5, i), False)
        D("created_by") = modUtilities.NullToDefault(Rows(6, i), "")
        D("created_date") = modUtilities.NullToDefault(Rows(7, i), 0)
        D("modified_date") = modUtilities.NullToDefault(Rows(8, i), 0)
        D("comments") = modUtilities.NullToDefault(Rows(9, i), "")
        D("is_default_quote") = True

        Q.LoadFromDict D

        ' Attach line items
        Dim Key As Variant
        For Each Key In LineItemIndex
            If CLng(LineItemIndex(Key).QuoteID) = Q.ID Then
                Q.UpsertLineItem LineItemIndex(Key)
            End If
        Next Key

        m_DefaultQuotes.Add Q
    Next i
End Sub

Private Sub AttachCustomerIDs(Rows As Variant)
    If IsEmpty(Rows) Or UBound(Rows) < 0 Then Exit Sub

    ' Col 0 = ID, Col 1 = quote_id, Col 2 = customer_id
    Dim i As Long
    For i = 0 To UBound(Rows, 2)
        Dim QuoteID As Long
        Dim CustomerID As Long
        QuoteID = CLng(modUtilities.NullToDefault(Rows(1, i), 0))
        CustomerID = CLng(modUtilities.NullToDefault(Rows(2, i), 0))

        If m_ByQuoteID.Exists(QuoteID) Then
            Dim Q As clsQuote
            Set Q = m_ByQuoteID(QuoteID)
            Q.AddCustomerID CustomerID

            ' Build customer index
            If Not m_ByCustomerID.Exists(CustomerID) Then
                Set m_ByCustomerID(CustomerID) = New Collection
            End If
            m_ByCustomerID(CustomerID).Add Q
        Else
            m_Logger.LogMessage "clsQuoteCache.AttachCustomerIDs", _
                LogLevelEnum.LogWarning, _
                "Quote ID not found: " & QuoteID
        End If
    Next i
End Sub

' ================================================================
' Private Helpers
' ================================================================
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    m_Logger.LogError "clsQuoteCache." & MethodName, Err.Number, _
        Err.Description & IIf(Context <> "", " | " & Context, "")
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsQuoteCache." & MethodName, _
            "The cache must be initialized before use."
    End If
End Sub

