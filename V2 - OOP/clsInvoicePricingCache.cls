VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoicePricingCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoicePricingCache
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-14
' Description: Caches invoice pricing details for lazy lookup
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsInvoiceQuoteBundle, modEnums
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Private m_Logger As clsLoggingSystem
Private m_AccessDatabase As clsAccessDatabase
Private m_PricingDict As Scripting.Dictionary
Private m_DefaultPriceDict As Scripting.Dictionary
Private Const INVOICE_DB_PATH As String = "\\PRECILAB-SERVER\LabPlusServer\Documents_In_Works\Thomson\Sample Login.accdb"
Private Const LINE_ITEM_ID_COLUMN As Long = 0
Private Const CUSTOMER_NAME_COLUMN As Long = 1
Private Const REQUEST_TYPE_COLUMN As Long = 2
Private Const BULK_MIN_COLUMN As Long = 3
Private Const CHEMICAL_NAME_COLUMN As Long = 4
Private Const PROCESSING_TIME_COLUMN As Long = 5
Private Const REGULAR_PRICE_COLUMN As Long = 6
Private Const BULK_PRICE_COLUMN As Long = 7
Public Property Get DefaultPriceDict() As Scripting.Dictionary
    ' Use lazy loading to generate the dictionary only when it's needed
    If m_DefaultPriceDict Is Nothing Then
        Set m_DefaultPriceDict = GenerateDefaultPricingDictionary()
    End If
    Set DefaultPriceDict = m_DefaultPriceDict
End Property

Public Property Get PricingDict() As Scripting.Dictionary
    ' Use lazy loading to generate the dictionary only when it's needed
    If m_PricingDict Is Nothing Then
        Set m_PricingDict = GeneratePricingDictionary()
    End If
    Set PricingDict = m_PricingDict
End Property
Public Sub Initialize(Logger As clsLoggingSystem, Optional AccessDatabase As clsAccessDatabase = Nothing)
    Call Logger.LogMessage("Initialize", LogLevelEnum.LogDebug, "Initializing the pricing cache")
    
    Set m_Logger = Logger

    If AccessDatabase Is Nothing Then
        ' Set a new database object
        Set m_AccessDatabase = New clsAccessDatabase
        Call m_AccessDatabase.Initialize(INVOICE_DB_PATH, Logger)
    Else
        ' Set the passed database object
        Set m_AccessDatabase = AccessDatabase
    End If
    
    Call Logger.LogMessage("Initialize", LogLevelEnum.LogDebug, "Successfully initializing the pricing cache")
End Sub
Private Function GeneratePricingDictionary() As Scripting.Dictionary
    Dim LineItemArray As Variant
    Dim AnalysisArray As Variant
    Dim ElementArray As Variant
    
    LineItemArray = GenerateLineItemArray(m_AccessDatabase)
    AnalysisArray = GenerateAnalysisArray(m_AccessDatabase)
    ElementArray = GenerateElementArray(m_AccessDatabase)
    
    Set GeneratePricingDictionary = CompilePricingDictionary(LineItemArray, AnalysisArray, ElementArray)
    
End Function
Private Function GenerateLineItemArray(AccessDatabase As clsAccessDatabase) As Variant
    Dim SQLQuery As String
    
    SQLQuery = "SELECT quote_line_items.ID, customers.customer_name, quote_line_items.request_type_id, quote_line_items.bulk_min, chemicals.chemical_name, quote_prices.processing_time_id, quote_prices.regular_price, quote_prices.bulk_price " & _
                "FROM chemicals RIGHT JOIN ((quotes INNER JOIN (customers INNER JOIN quote_customer ON customers.ID = quote_customer.customer_id) ON quotes.ID = quote_customer.quote_id) INNER JOIN (quote_line_items INNER JOIN quote_prices ON quote_line_items.ID = quote_prices.quote_line_item_id) ON quotes.ID = quote_line_items.quote_id) ON chemicals.ID = quote_line_items.chemical_id;"

    GenerateLineItemArray = AccessDatabase.ExecuteQuery(SQLQuery)
    
End Function
Private Function GenerateAnalysisArray(AccessDatabase As clsAccessDatabase) As Variant
    Dim SQLQuery As String

    SQLQuery = "SELECT quote_analysis.quote_line_item_id, analyses.analysis_name " & _
                "FROM quote_analysis " & _
                "LEFT JOIN analyses ON quote_analysis.analysis_id = analyses.ID;"

    GenerateAnalysisArray = AccessDatabase.ExecuteQuery(SQLQuery)
End Function
Private Function GenerateElementArray(AccessDatabase As clsAccessDatabase) As Variant
    Dim SQLQuery As String
    
    SQLQuery = "SELECT quote_line_item_element.quote_line_item_id, elements.element_symbol " & _
                "FROM quote_line_item_element " & _
                "LEFT JOIN elements ON quote_line_item_element.element_id = elements.ID;"

    GenerateElementArray = AccessDatabase.ExecuteQuery(SQLQuery)
End Function
Private Function CompilePricingDictionary(LineItemArray As Variant, AnalysisArray As Variant, ElementArray As Variant) As Scripting.Dictionary
    Dim LineItemIndex As Long
    Dim PricingDict As New Scripting.Dictionary
    Dim StaticKey As String
    Dim BundleCollection As Collection
    Dim QuoteBundleObj As clsInvoiceQuoteBundle
    
    
    For LineItemIndex = LBound(LineItemArray, 2) To UBound(LineItemArray, 2)
        
        ' Set the bundle object - this is a bundle of analyses (1 to many) that has been quoted to the customer for one price
        Set QuoteBundleObj = GenerateQuoteBundle(LineItemArray, LineItemIndex, AnalysisArray, ElementArray)
        
        ' Set the key for the collection of bundles
        StaticKey = GenerateStaticKey(LineItemArray, LineItemIndex)
        
        ' Retrieve the bundle from the dictionary or create a new one
        If PricingDict.Exists(StaticKey) Then
            Set BundleCollection = PricingDict(StaticKey)
        Else
            Set BundleCollection = New Collection
            PricingDict.Add StaticKey, BundleCollection
        End If
                
        ' Add the bundle object to the collection, sorted by analyses.count
        Call InsertBundleSorted(BundleCollection, QuoteBundleObj)

    Next LineItemIndex
    
    ' Return the dictionary to the calling sub
    Set CompilePricingDictionary = PricingDict
End Function
Private Function GenerateQuoteBundle(LineItemArray As Variant, LineItemIndex As Long, AnalysisArray As Variant, ElementArray As Variant) As clsInvoiceQuoteBundle
    Dim QuoteBundleObj As New clsInvoiceQuoteBundle
    Dim AnalysisCollection As Collection
    Dim ElementCollection As Collection
    Dim LineItemID As Long
    Dim RegularPrice As Long
    Dim BulkPrice As Long
    Dim BulkMin As Long
    
    
    ' Set the line item variables
    LineItemID = LineItemArray(LINE_ITEM_ID_COLUMN, LineItemIndex)
    RegularPrice = LineItemArray(REGULAR_PRICE_COLUMN, LineItemIndex)
    BulkPrice = LineItemArray(BULK_PRICE_COLUMN, LineItemIndex)
    BulkMin = LineItemArray(BULK_MIN_COLUMN, LineItemIndex)
    
    ' Generate the collections
    Set AnalysisCollection = GenerateObjectCollection(LineItemID, AnalysisArray)
    Set ElementCollection = GenerateObjectCollection(LineItemID, ElementArray)
        
    ' Initialize and fill the object
    Call QuoteBundleObj.Initialize(RegularPrice, BulkPrice, BulkMin, AnalysisCollection, ElementCollection, m_Logger)
    
    ' Return the object to the calling sub
    Set GenerateQuoteBundle = QuoteBundleObj
    
End Function
Private Function GenerateObjectCollection(LineItemID As Long, arr As Variant) As Collection
    Dim ObjectCollection As New Collection
    Dim iIndex As Long
    
    
    For iIndex = LBound(arr, 2) To UBound(arr, 2)
        If arr(0, iIndex) = LineItemID Then
            ObjectCollection.Add arr(1, iIndex)
        End If
    Next iIndex
    
    ' Return the collection to the calling sub
    Set GenerateObjectCollection = ObjectCollection
    
End Function
Private Function GenerateStaticKey(LineItemArray As Variant, LineItemIndex As Long) As String
    Dim CustomerName As String
    Dim RequestType As TestingRequestTypeEnum
    Dim ChemicalName As String
    Dim ProcessingTime As ProcessingTimeEnum
    
    CustomerName = LineItemArray(CUSTOMER_NAME_COLUMN, LineItemIndex)
    RequestType = LineItemArray(REQUEST_TYPE_COLUMN, LineItemIndex)
    If IsNull(LineItemArray(CHEMICAL_NAME_COLUMN, LineItemIndex)) Then
        ChemicalName = "Null"
    Else
        ChemicalName = LineItemArray(CHEMICAL_NAME_COLUMN, LineItemIndex)
    End If
    ProcessingTime = LineItemArray(PROCESSING_TIME_COLUMN, LineItemIndex)
    
    ' Return the string to the calling sub
    GenerateStaticKey = CustomerName & "|" & RequestType & "|" & ChemicalName & "|" & ProcessingTime
    
End Function
Private Sub InsertBundleSorted(ByRef BundleCollection As Collection, NewBundle As clsInvoiceQuoteBundle)
    Dim iIndex As Long
    Dim TempBundle As clsInvoiceQuoteBundle
    Dim NewBundleSize As Long
    Dim TempBundleSize As Long

    ' Calculate the size of the new bundle (analyses + additional elements)
    NewBundleSize = NewBundle.Analyses.Count + NewBundle.Elements.Count
    
    ' If the collection is empty, just add the bundle
    If BundleCollection.Count = 0 Then
        BundleCollection.Add NewBundle
        Exit Sub
    End If
    
    ' Loop through the collection to find the right place to insert
    For iIndex = 1 To BundleCollection.Count
        Set TempBundle = BundleCollection(iIndex)
        TempBundleSize = TempBundle.Analyses.Count + TempBundle.Elements.Count
        If NewBundleSize >= TempBundleSize Then
            BundleCollection.Add NewBundle, Before:=iIndex
            Exit Sub
        End If
    Next iIndex
    
    ' If not inserted, add to the end
    BundleCollection.Add NewBundle
End Sub
Private Function GenerateDefaultPricingDictionary() As Scripting.Dictionary
    Dim DefaultPriceArray As Variant
    
    DefaultPriceArray = GetDefaultPriceArray()
    
    Set GenerateDefaultPricingDictionary = GetDefaultPriceDict(DefaultPriceArray)
    
End Function
Private Function GetDefaultPriceArray() As Variant
    Dim SQLQuery As String
    
    SQLQuery = "SELECT analyses.analysis_name, default_price.processing_time_id, default_price.price " & _
                "FROM analyses " & _
                "LEFT JOIN default_price ON analyses.id = default_price.analysis_id " & _
                "WHERE default_price.start_date <= #" & Now & "# AND default_price.end_date >= #" & Now & "# AND default_price.is_active = TRUE;"

    GetDefaultPriceArray = m_AccessDatabase.ExecuteQuery(SQLQuery)
End Function
Private Function GetDefaultPriceDict(DefaultPriceArray As Variant) As Scripting.Dictionary
    Dim DefaultPriceDict As New Scripting.Dictionary
    Dim ArrayIndex As Long
    Dim StaticKey As String
    Const DEFAULT_PRICE_COLUMN As Long = 2
    
    ' Loop through the array
    For ArrayIndex = 0 To UBound(DefaultPriceArray, 2)
        ' Generate the static key
        StaticKey = GenerateDefaultPriceKey(ArrayIndex, DefaultPriceArray)
        ' check that the key doesn't already exist in the dictionary
        If Not DefaultPriceDict.Exists(StaticKey) Then
            ' add the price to the dictionary with the analysis name as the key
            DefaultPriceDict.Add StaticKey, DefaultPriceArray(DEFAULT_PRICE_COLUMN, ArrayIndex)
        End If
    Next ArrayIndex
    
    ' Log the success
    Call m_Logger.LogMessage("GetDefaultPriceDict", LogLevelEnum.LogInfo, "Compiled the default pricing dictionary with " & DefaultPriceDict.Count & " entries.")

    ' Return the dictionary to the calling sub
    Set GetDefaultPriceDict = DefaultPriceDict

End Function
Private Function GenerateDefaultPriceKey(ArrayIndex As Long, DefaultPriceArray As Variant) As String
    Dim Analysis As String
    Dim ProcessingTime As ProcessingTimeEnum
    Const DEFAULT_ANALYSIS_COLUMN As Long = 0
    Const DEFAULT_PROCESSING_TIME_COLUMN As Long = 1
    
    Analysis = DefaultPriceArray(DEFAULT_ANALYSIS_COLUMN, ArrayIndex)
    ProcessingTime = DefaultPriceArray(DEFAULT_PROCESSING_TIME_COLUMN, ArrayIndex)
    
    ' Return the static key to the calling sub
    GenerateDefaultPriceKey = Analysis & "|" & ProcessingTime
End Function
