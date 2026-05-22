VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuoteLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsQuoteLoader
' Author: JT
' Created: 2025-12-29
' Description: Loader object for quotes
' DependsOn: clsLoggingSystem, modEnums, clsQuote, clsQuoteLineItem, clsQuotePrice
' ChangeLog:
'   - 0.1.0 - 2025-12-29 - Initial prototype
'   - 0.2.0 - 2026-04-03 - Added m_DefaultQuotes collection and logic
'   - 0.3.0 - 2026-04-14 - Fixed wrong types/columns/method names in Build*Dict; updated Attach*/Add* to use current domain-object API (UpsertPrice, UpsertLineItem, AddAnalysisID, AddElementID, AddCustomerID); removed dead GroupChildObjects/CreateIDDict helpers

Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_Quotes As Collection
Private m_DefaultQuotes As Collection
Private Type IDPair
    ParentID As Long
    ChildID  As Long
End Type
' ==============
' Public Getters
' ==============
Public Property Get Quotes() As Collection
    If m_Quotes Is Nothing Then
        Set m_Quotes = BuildQuoteCollection
    End If
    Set Quotes = m_Quotes
End Property
Public Property Get DefaultQuotes() As Collection
    If m_DefaultQuotes Is Nothing Then
        Set m_DefaultQuotes = BuildDefaultQuoteCollection
    End If
    Set DefaultQuotes = m_DefaultQuotes
End Property
' ==============
' Initialization
' ==============
Public Sub Initialize(AccessDatabase As clsAccessDatabase, Logger As clsLoggingSystem)
    If Not m_Logger Is Nothing Then Err.Raise vbObjectError + 1, , "The loader is already initialized."

    If Not Logger.DebugMode Then On Error GoTo ErrorHandler

    Call Logger.LogMessage("clsQuoteLoader.Initialize", LogLevelEnum.LogDebug, "Initializing Quote Loader")

    Set m_Logger = Logger
    Set m_AccessDB = AccessDatabase

    Call Logger.LogMessage("clsQuoteLoader.Initialize", LogLevelEnum.LogInfo, "Completed initializing Quote Loader")

    Exit Sub
ErrorHandler:
    HandleError "Initialize"
End Sub
' ===============
' Public API
' ===============
Public Function LoadQuotes() As Collection

    AssertInitialized "LoadQuotes"

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    ' Refresh quote collection
    Set m_Quotes = BuildQuoteCollection

    ' Return collection to calling sub
    Set LoadQuotes = m_Quotes

    Exit Function
ErrorHandler:
    HandleError "LoadQuotes"
End Function
' ===============
' Private Helpers
' ===============
Private Function BuildQuoteCollection() As Collection
    Dim QuoteDict As Scripting.Dictionary
    Dim LineItemDict As Scripting.Dictionary
    Dim PriceDict As Scripting.Dictionary
    Dim C As New Collection
    Dim AnalysisData As Variant
    Dim ElementData As Variant

    AssertInitialized "BuildQuoteCollection"

    ' Build flat object dictionaries
    Set QuoteDict = BuildQuoteDict
    Set LineItemDict = BuildLineItemDict
    Set PriceDict = BuildPriceDict

    AnalysisData = GetQuoteAnalyses()
    ElementData = GetQuoteElements()

    ' Build the objects into a heirachy inside quote objects
    AttachPricesToLineItems PriceDict, LineItemDict
    AttachAnalysisIDsToLineItems LineItemDict
    AttachElementIDsToLineItems LineItemDict
    AttachLineItemsToQuotes LineItemDict, QuoteDict
    AttachCustomerIDsToQuotes QuoteDict

    ' Convert the dict to a collection
    Set C = ConvertDictToCollection(QuoteDict)

    ' Return the collection to the calling sub
    Set BuildQuoteCollection = C

End Function
Private Function BuildDefaultQuoteCollection() As Collection
    Dim QuoteDict As Scripting.Dictionary
    Dim LineItemDict As Scripting.Dictionary
    Dim PriceDict As Scripting.Dictionary
    Dim C As New Collection
    Dim AnalysisData As Variant

    AssertInitialized "BuildDefaultQuoteCollection"

    ' Build flat object dictionaries from default tables
    Set QuoteDict = BuildDefaultQuoteDict
    Set LineItemDict = BuildDefaultLineItemDict
    Set PriceDict = BuildDefaultPriceDict

    AnalysisData = GetDefaultQuoteAnalyses()
    ' No elements in default quotes

    ' Build the hierarchy
    AttachPricesToLineItems PriceDict, LineItemDict
    AttachAnalysisIDsToLineItems LineItemDict
    AttachElementIDsToLineItems LineItemDict
    AttachLineItemsToQuotes LineItemDict, QuoteDict
    AttachCustomerIDsToQuotes QuoteDict

    ' NOTE: Default quotes do not have customers, so skip AddCustomersToQuotes

    ' Convert dict to collection
    Set C = ConvertDictToCollection(QuoteDict)
    Set BuildDefaultQuoteCollection = C
End Function
Private Function BuildQuoteDict() As Scripting.Dictionary
    Set BuildQuoteDict = FetchQuoteDicts("quotes")
End Function
Private Function BuildDefaultQuoteDict() As Scripting.Dictionary
    Set BuildDefaultQuoteDict = FetchQuoteDicts("default_quotes")
End Function
Private Function FetchQuoteDicts(TableName As String) As Scripting.Dictionary
    Dim Q As clsQuote
    Dim RowDict As Scripting.Dictionary
    Dim Rows As Collection
    Dim Dict As New Scripting.Dictionary

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    ' Fetch rows collection
    Set Rows = FetchRows(TableName, "ID, effective_date, expiry_date, payment_type, created_by, created_date, modified_date, valid")

    ' Loop through each row and fill out objects with the row dictionary (matching keys to properties)
    For Each RowDict In Rows
        If Not Dict.Exists(RowDict("ID")) Then
            Set Q = New clsQuote
            Call Q.Initialize(m_Logger)
            Call Q.FromDict(RowDict)
            Dict.Add Q.ID, Q
        End If
    Next RowDict

    ' Return the object to the calling sub
    Set BuildQuoteDict = Dict
    Exit Function

ErrorHandler:
    Dim RowID As Long
    If RowDict.Exists("ID") Then RowID = CInt(RowDict("ID"))
    HandleError "FetchQuoteDicts", "Error on row " & RowID
End Function
Private Function BuildLineItemDict() As Scripting.Dictionary
    Set BuildLineItemDict = FetchLineItemDict("quote_line_items")
End Function
Private Function BuildDefaultLineItemDict() As Scripting.Dictionary
    Set BuildDefaultLineItemDict = FetchLineItemDict("default_quote_line_items")
End Function
Private Function FetchLineItemDict(TableName As String) As Scripting.Dictionary
    Dim LI As clsQuoteLineItem
    Dim RowDict As Scripting.Dictionary
    Dim Rows As Collection
    Dim Dict As New Scripting.Dictionary

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    ' Fetch rows collection
    Set Rows = FetchRows(TableName, "id, quote_id, request_type_id, bulk_min, chemical_id, valid")

    ' Loop through each row and fill out objects with the row dictionary (matching keys to properties)
    For Each RowDict In Rows
        If Not Dict.Exists(RowDict("id")) Then
            Set LI = New clsQuoteLineItem
            Call LI.Initialize(m_Logger)
            Call LI.FromDict(RowDict)
            Dict.Add LI.ID, LI
        End If
    Next RowDict

    ' Return the object to the calling sub
    Set BuildQuoteLineItemDict = Dict
    Exit Function

ErrorHandler:
    Dim RowID As Long
    If RowDict.Exists("id") Then RowID = CInt(RowDict("id"))
    HandleError "FetchLineItemDict", "Error on row " & RowID
End Function
Private Function BuildPriceDict() As Scripting.Dictionary
    Set BuildPriceDict = FetchPriceDict("quote_prices")
End Function
Private Function BuildDefaultPriceDict() As Scripting.Dictionary
    Set BuildDefaultPriceDict = FetchPriceDict("default_quote_prices")
End Function
Private Function FetchPriceDict(TableName As String) As Scripting.Dictionary
    Dim P As clsQuotePrice
    Dim RowDict As Scripting.Dictionary
    Dim Rows As Collection
    Dim Dict As New Scripting.Dictionary

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    ' Fetch rows collection
    Set Rows = FetchRows(TableName, "id, quote_line_item_id, regular_price, bulk_price, processing_time, valid")

    ' Loop through each row and fill out objects with the row dictionary (matching keys to properties)
    For Each RowDict In Rows
        If Not Dict.Exists(RowDict("id")) Then
            Set P = New clsQuotePrice
            Call P.Initialize(m_Logger)
            Call P.FromDict(RowDict)
            Dict.Add P.ID, P
        End If
    Next RowDict

    ' Return the object to the calling sub
    Set BuildQuotePriceDict = Dict
    Exit Function

ErrorHandler:
    Dim RowID As Long
    If RowDict.Exists("id") Then RowID = CInt(RowDict("id"))
    HandleError "FetchPriceDict", "Error on row " & RowID
End Function

Private Sub AttachPricesToLineItems(PriceDict As Scripting.Dictionary, LineItemDict As Scripting.Dictionary)
    Dim Key As Variant
    Dim PriceObj As clsQuotePrice
    Dim LI As clsQuoteLineItem

    For Each Key In PriceDict
        Set PriceObj = PriceDict(Key)
        Set LI = LineItemDict(PriceObj.QuoteLineItemID)
        LI.UpsertPrice PriceObj
    Next Key
End Sub
Private Sub AttachLineItemsToQuotes(LineItemDict As Scripting.Dictionary, QuoteDict As Scripting.Dictionary)
    Dim Key As Variant
    Dim LI As clsQuoteLineItem
    Dim Q As clsQuote

    For Each Key In LineItemDict
        Set LI = LineItemDict(Key)
        Set Q = QuoteDict(LI.QuoteID)
        Q.UpsertLineItem LI
    Next Key
End Sub
Private Function ConvertDictToCollection(QuoteDict As Scripting.Dictionary) As Collection
    Dim C As New Collection
    Dim Key As Variant

    For Each Key In QuoteDict
        C.Add QuoteDict(Key)
    Next Key

    Set ConvertDictToCollection = C

End Function
Private Sub AttachAnalysisIDsToLineItems(LineItemDict As Scripting.Dictionary)
    Dim Pairs() As IDPair
    Dim i       As Long
    Dim LI      As clsQuoteLineItem

    Pairs = GetQuoteAnalysisPairs()

    For i = LBound(Pairs) To UBound(Pairs)
        If LineItemDict.Exists(Pairs(i).ParentID) Then
            Set LI = LineItemDict(Pairs(i).ParentID)
            LI.AddAnalysisID Pairs(i).ChildID
        Else
            m_Logger.LogError "clsQuoteLoader.AttachAnalysisIDsToLineItems", "LineItem ID not found: " & Pairs(i).ParentID
        End If
    Next i
End Sub

Private Sub AttachElementIDsToLineItems(LineItemDict As Scripting.Dictionary)
    Dim Pairs() As IDPair
    Dim i       As Long
    Dim LI      As clsQuoteLineItem

    Pairs = GetQuoteElementPairs()

    For i = LBound(Pairs) To UBound(Pairs)
        If LineItemDict.Exists(Pairs(i).ParentID) Then
            Set LI = LineItemDict(Pairs(i).ParentID)
            LI.AddElementID Pairs(i).ChildID
        Else
            m_Logger.LogError "clsQuoteLoader.AttachElementIDsToLineItems", "LineItem ID not found: " & Pairs(i).ParentID
        End If
    Next i
End Sub
Private Sub AttachCustomerIDsToQuotes(QuoteDict As Scripting.Dictionary)
    Dim Pairs() As IDPair
    Dim i       As Long
    Dim Q       As clsQuote

    Pairs = GetQuoteCustomerPairs()

    For i = LBound(Pairs) To UBound(Pairs)
        If QuoteDict.Exists(Pairs(i).ParentID) Then
            Set Q = QuoteDict(Pairs(i).ParentID)
            Q.AddCustomerID Pairs(i).ChildID
        Else
            m_Logger.LogError "clsQuoteLoader.AttachCustomerIDsToQuotes", "Quote ID not found: " & Pairs(i).ParentID
        End If
    Next i
End Sub

' ===========
' SQL helpers
' ===========
Private Function FetchRows(TableName As String, ColumnHeaders As String) As Collection
    Dim SQL As String
    Dim DataArray As Variant
    Dim RowDict As Scripting.Dictionary
    Dim Rows As New Collection
    Dim iIndex As Long
    Dim Headers() As String

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler


    SQL = "SELECT " & ColumnHeaders & " FROM " & TableName & " ORDER BY ID;"
    DataArray = m_AccessDB.ExecuteQuery(SQL)

    Headers = Split(ColumnHeaders, ", ")
    For iIndex = LBound(DataArray, 2) To UBound(DataArray, 2)
        Set RowDict = ArrayRowToDict(DataArray, iIndex, Headers)
        Rows.Add RowDict
    Next iIndex

    Set FetchRows = Rows
    Exit Function
ErrorHandler:
    Call m_Logger.LogMessage("clsQuoteLoader.FetchRows", LogLevelEnum.LogError, "Error " & Err.Number & ": " & Err.Description & " loading table " & TableName)
    Err.Raise Err.Number, "clsQuoteLoader.FetchRows", Err.Description
End Function
Private Function GetQuoteAnalysisPairs() As IDPair()
    GetQuoteAnalysisPairs = FetchIDPairs( _
        "SELECT quote_line_item_id, analysis_id FROM quote_analysis")
End Function

Private Function GetDefaultQuoteAnalysisPairs() As IDPair()
    GetDefaultQuoteAnalysisPairs = FetchIDPairs( _
        "SELECT default_quote_line_item_id, analysis_id FROM default_quote_analysis")
End Function

Private Function GetQuoteElementPairs() As IDPair()
    GetQuoteElementPairs = FetchIDPairs( _
        "SELECT quote_line_item_id, element_id FROM quote_line_item_element")
End Function

Private Function GetQuoteCustomerPairs() As IDPair()
    GetQuoteCustomerPairs = FetchIDPairs( _
        "SELECT quote_id, customer_id FROM quote_customer")
End Function
Private Function FetchIDPairs(SQL As String) As IDPair()
    Dim Data    As Variant
    Dim Pairs() As IDPair
    Dim i       As Long
    Dim Count   As Long

    Data = m_AccessDB.ExecuteQuery(SQL)
    Count = UBound(Data, 2) - LBound(Data, 2) + 1
    ReDim Pairs(1 To Count)

    For i = 1 To Count
        Pairs(i).ParentID = CLng(Data(0, i - 1))
        Pairs(i).ChildID = CLng(Data(1, i - 1))
    Next i

    FetchIDPairs = Pairs
End Function

Private Function ArrayRowToDict(DataArray As Variant, RowIndex As Long, Headers() As String) As Scripting.Dictionary
    Dim RowDict As New Scripting.Dictionary
    Dim ColIndex As Long

    For ColIndex = LBound(Headers) To UBound(Headers)
        RowDict.Add Trim(Headers(ColIndex)), DataArray(ColIndex, RowIndex)
    Next ColIndex

    Set ArrayRowToDict = RowDict
End Function
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsQuoteLoader." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsQuoteLoader." & MethodName, "The builder must be initialized before use."
    End If
End Sub
