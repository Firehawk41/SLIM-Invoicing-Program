VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoicePriceRetriever"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoicePriceRetriever
' Version: 1.2.1
' Author: JT
' Created: ?
' LastModified: 2025-10-14
' Description: Retrieves prices for invoices
' DependsOn: clsLoggingSystem, clsInvoicePricingCache, clsTestingRequestSubmission, clsTestingRequestSample, modEnums, clsTestingRequestAnalysis, clsTestingRequestElement, clsInvoiceQuoteBundle, clsInvoicePriceDetail
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
'   - 1.1.0 - 2025-07-07 - Updated to store new class - clsInvoicePriceDetail in the pricing dictionary instead of just price
'   - 1.2.0 - 2025-07-08 - Updated default price to work with clsInvoicePriceDetail and corrected bug in the log that wouldn't display the price or analysis name
'   - 1.2.1 - 2025-10-14 - Fixed bug where translated analysis names weren't matched in the default pricing dictionary by adding reverse translation lookup.
Private m_Logger As clsLoggingSystem
Private m_PricingCache As clsInvoicePricingCache
Private m_TranslationDictionary As Scripting.Dictionary
Private m_ReverseDictionary As Scripting.Dictionary
Public Property Get PricingDict() As Scripting.Dictionary
    If m_PricingCache Is Nothing Then
        Set m_PricingCache = New clsInvoicePricingCache
        Call m_PricingCache.Initialize(m_Logger)
    End If
    Set PricingDict = m_PricingCache.PricingDict
End Property
Public Property Get DefaultPriceDict() As Scripting.Dictionary
    If m_PricingCache Is Nothing Then
        Set m_PricingCache = New clsInvoicePricingCache
        Call m_PricingCache.Initialize(m_Logger)
    End If
    Set DefaultPriceDict = m_PricingCache.DefaultPriceDict
End Property
Public Property Get TranslationDictionary() As Scripting.Dictionary
    If m_TranslationDictionary Is Nothing Then
        Set m_TranslationDictionary = InitializeTranslationDictionary()
    End If
    Set TranslationDictionary = m_TranslationDictionary
End Property
Public Property Get ReverseDictionary() As Scripting.Dictionary
    ' TranslationDictionary sets the reverse dictionary as well
    If m_ReverseDictionary Is Nothing Then
        Set m_ReverseDictionary = InitializeReverseDictionary(Me.TranslationDictionary)
    End If
    Set ReverseDictionary = m_ReverseDictionary
End Property
Public Sub Initialize(Logger As clsLoggingSystem)
    Call Logger.LogMessage("Initialize", LogLevelEnum.LogDebug, "Initializing Price Retriever")
    
    Set m_Logger = Logger
    
End Sub
Public Function RetrievePrices(Submission As clsTestingRequestSubmission, Sample As clsTestingRequestSample, Quantity As Long) As Scripting.Dictionary
    Dim StaticKeyWithoutChemical As String
    Dim StaticKeyWithChemical As String
    Dim AnalysesDict As Scripting.Dictionary
    Dim PriceDict As New Scripting.Dictionary
    
    Call m_Logger.LogMessage("RetrievePrices", LogLevelEnum.LogDebug, "Retrieving prices for " & Sample.SampleName)
    
    ' Generate a static key for the pricing dictionary both with and without the chemical name
    StaticKeyWithoutChemical = GenerateStaticKey(Submission, Sample, False)
    StaticKeyWithChemical = GenerateStaticKey(Submission, Sample, True)
    
    ' Clone the analyses collection so that the original is kept intact
    Set AnalysesDict = ConvertCollectionToDictionary(Sample)
    
    
    ' Check whether there are any prices for the specified chemical
    If Me.PricingDict.Exists(StaticKeyWithChemical) Then
        Call UpdatePriceDict(PriceDict, AnalysesDict, StaticKeyWithChemical, Quantity)
    End If

    ' Check whether there are analyses not yet matched above and whether there is a price without the specific chemical
    If Not AreAllSetToTrue(AnalysesDict) And Me.PricingDict.Exists(StaticKeyWithoutChemical) Then
        Call UpdatePriceDict(PriceDict, AnalysesDict, StaticKeyWithoutChemical, Quantity)
    End If
    
    ' Add any remaining analyses with 0 prices
    If Not AreAllSetToTrue(AnalysesDict) Then
        Call AddUnmatchedAnalysesToPriceDict(PriceDict, AnalysesDict, Sample.ProcessingTime)
    End If
    
    ' Return the dictionary to the calling sub
    Set RetrievePrices = PriceDict
    
    Call m_Logger.LogMessage("RetrievePrices", LogLevelEnum.LogInfo, "Completed retrieving prices for " & Sample.SampleName)
    
End Function
Private Function GenerateStaticKey(Submission As clsTestingRequestSubmission, Sample As clsTestingRequestSample, IsWithChemical As Boolean) As String
    Dim CustomerName As String
    Dim RequestType As TestingRequestTypeEnum
    Dim ChemicalName As String
    Dim ProcessingTime As ProcessingTimeEnum
    
    CustomerName = Submission.Customer.CustomerName
    RequestType = Submission.RequestType
    ChemicalName = IIf(IsWithChemical, Sample.Chemical.ChemicalName, "Null")
    If RequestType = TestingRequestTypeEnum.Water And Sample.ProcessingTime = ProcessingTimeEnum.ExtendedTime Then
        ' Since extended time is priced the same as next day for water samples, convert the processing time to next day
        ProcessingTime = ProcessingTimeEnum.NextDay
    Else
        ProcessingTime = Sample.ProcessingTime
    End If
    
    ' Return the string to the calling sub
    GenerateStaticKey = CustomerName & "|" & RequestType & "|" & ChemicalName & "|" & ProcessingTime

End Function
Private Function InitializeTranslationDictionary() As Scripting.Dictionary
    Dim TranslationMap As New Scripting.Dictionary
    Dim ReverseMap As New Scripting.Dictionary
    Dim Key As Variant
    
    ' Populate translation dictionary
    TranslationMap.Add "4 Anions", "Anions"
    TranslationMap.Add "5 Anions", "Anions"
    TranslationMap.Add "7 Anions", "Anions"
    TranslationMap.Add "4 Anions + MS Confirmation", "Anions + MS Confirmation"
    TranslationMap.Add "5 Anions + MS Confirmation", "Anions + MS Confirmation"
    TranslationMap.Add "7 Anions + MS Confirmation", "Anions + MS Confirmation"
    TranslationMap.Add "NH4", "Cations"
    TranslationMap.Add "6 Cations", "Cations"
    
    ' Return the dict to the calling sub
    Set InitializeTranslationDictionary = TranslationMap
    
End Function
Private Function InitializeReverseDictionary(TranslationDictionary As Scripting.Dictionary) As Scripting.Dictionary
    Dim ReverseMap As New Scripting.Dictionary
    Dim Key As Variant
    Dim CanonicalName As String
    Dim Coll As Collection
    
    ' Loop through the translation dictionary's keys
    For Each Key In TranslationDictionary
        CanonicalName = TranslationDictionary(Key)
        
        ' Ensure that the canonical name exists as a key in the reverse dict
        If Not ReverseMap.Exists(CanonicalName) Then
            Set Coll = New Collection
            ReverseMap.Add CanonicalName, Coll
        End If
        
        ' Add the canonical name to the collection
        ReverseMap(CanonicalName).Add Key
    Next Key
    
    ' Return the dict to the calling sub
    Set InitializeReverseDictionary = ReverseMap
    
End Function
Private Function ConvertCollectionToDictionary(Sample As clsTestingRequestSample, Optional TranslateNames As Boolean = True) As Scripting.Dictionary
    Dim AnalysesCollection As Collection
    Dim ElementCollection As Collection
    Dim AnalysisDict As New Scripting.Dictionary
    Dim ElementDict As New Scripting.Dictionary
    Dim Analysis As clsTestingRequestAnalysis
    Dim Element As clsTestingRequestElement
    Dim TranslatedKey As String
    
    ' Set objects
    Set AnalysesCollection = Sample.Analyses
    Set ElementCollection = Sample.AdditionalElements
    
    ' Loop through the analyses collection and add items to the analysis dictionary
    For Each Analysis In AnalysesCollection
        
        
        ' Translate if exists in lookup table, else keep original name
        If Me.TranslationDictionary.Exists(Analysis.AnalysisName) And TranslateNames Then
            TranslatedKey = Me.TranslationDictionary(Analysis.AnalysisName)
        Else
            TranslatedKey = Analysis.AnalysisName
        End If
        
        ' Avoid duplicate keys in case multiple variations exist
        If Not AnalysisDict.Exists(TranslatedKey) Then
            AnalysisDict.Add TranslatedKey, False
        End If
        
    Next Analysis
    
    ' Loop through the additional elements and add them to a dictionary
    For Each Element In ElementCollection
        If Not ElementDict.Exists(Element.ElementSymbol) Then
            ElementDict.Add Element.ElementSymbol, False
        End If
    Next Element
    
    ' Add the element dict to the analysis dict
    AnalysisDict.Add "__ElementDict", ElementDict
        
    ' Return the new collection to the calling sub
    Set ConvertCollectionToDictionary = AnalysisDict
End Function
Private Sub UpdatePriceDict(PriceDict As Scripting.Dictionary, AnalysesDict As Scripting.Dictionary, StaticKey As String, Quantity As Long)
    Dim QuoteBundleObj As clsInvoiceQuoteBundle
    Dim AnalysisStr As String
    Dim Price As Long
    Dim PriceDetail As clsInvoicePriceDetail

    ' Check whether there are entries in the database for this static key
    If Me.PricingDict.Exists(StaticKey) Then
        
        ' Loop through each bundle for this static key
        For Each QuoteBundleObj In Me.PricingDict(StaticKey)
            
            ' Check whether this bundle matches the analyses ordered by the customer
            If QuoteBundleObj.MatchBundle(AnalysesDict) Then
                
                ' Set the variables to be added to the dictionary
                Set PriceDetail = SetPriceDetail(QuoteBundleObj.Price(Quantity), QuoteBundleObj.IsPriceBundle)
                
                ' Add the analysis string and price to the dictionary
                PriceDict.Add QuoteBundleObj.AnalysisStr, PriceDetail
                
                ' Log the success
                Call m_Logger.LogMessage("UpdatePriceDict", LogLevelEnum.LogDebug, "Added price of $" & PriceDetail.Price & " for " & QuoteBundleObj.AnalysisStr)
            
            End If
            
            ' Exit early if all analyses have been matched
            If AreAllSetToTrue(AnalysesDict) Then
                Call m_Logger.LogMessage("UpdatePriceDict", LogLevelEnum.LogInfo, "All analyses have had prices assigned")
                Exit Sub
            End If
            
        Next QuoteBundleObj
        
    End If
    
End Sub
Private Function SetPriceDetail(Price As Long, Optional IsPriceBundle As Boolean = False) As clsInvoicePriceDetail
    Dim PriceDetail As clsInvoicePriceDetail
    
    Set PriceDetail = New clsInvoicePriceDetail
    PriceDetail.Price = Price
    PriceDetail.IsFlatRate = IsPriceBundle
    
    ' Return object to calling sub
    Set SetPriceDetail = PriceDetail
End Function
Private Function AreAllSetToTrue(AnalysesDict As Scripting.Dictionary) As Boolean
    Dim Key As Variant
    
    ' Iterate across all entries in the dictionary to see if all are set to true
    For Each Key In AnalysesDict.Keys
        
        If Not Key = "__ElementDict" Then ' skip the element dict
            If AnalysesDict(Key) = False Then
                AreAllSetToTrue = False
                Exit Function
            End If
        End If
    Next Key
    
    ' Set to true if no 'false' entries are found
    AreAllSetToTrue = True
End Function
Private Sub AddUnmatchedAnalysesToPriceDict(PriceDict As Scripting.Dictionary, AnalysesDict As Scripting.Dictionary, ProcessingTime As ProcessingTimeEnum)
    Dim Key As Variant
    Dim StaticKey As String
    Dim PriceDetail As clsInvoicePriceDetail
    Dim UntranslatedKey As Variant
    Dim UntranslatedStaticKey As String
    
    ' Loop through each analysis in the analyses collection
    For Each Key In AnalysesDict
        If Not Key = "__ElementDict" Then ' Skip the element dict

            ' Check whether the analysis has been matched with a quote
            If AnalysesDict(Key) = False Then
                
                ' Generate static key for this analysis/processing time
                StaticKey = Key & "|" & ProcessingTime
                ' Check whether a default price exists for the analysis
                If Me.DefaultPriceDict.Exists(StaticKey) Then
                    
                    Call AddDefaultPriceDetail(StaticKey, Key, PriceDict)
                ' Check whether the analysis name needs to be un-translated
                ElseIf Me.ReverseDictionary.Exists(Key) Then
                    
                    ' Loop through each analysis in the collection to find a default price
                    For Each UntranslatedKey In Me.ReverseDictionary(Key)
                        UntranslatedStaticKey = UntranslatedKey & "|" & ProcessingTime
                        If Me.DefaultPriceDict.Exists(UntranslatedStaticKey) Then
                            Call AddDefaultPriceDetail(UntranslatedStaticKey, Key, PriceDict)
                            ' Early exit if match is found
                            Exit For
                        End If
                    Next UntranslatedKey
                
                Else
                    ' If not matched, add it to the pricing dictionary with -1 pricetag to indicate that a quote doesn't exist
                    Set PriceDetail = SetPriceDetail(-1)
                    PriceDict.Add Key, PriceDetail
                    Call m_Logger.LogMessage("AddUnmatchedAnalysesToPriceDict", LogLevelEnum.LogWarning, "Unable to find price match for " & Key)
                End If
            End If
        End If
        
    Next Key
End Sub
Private Sub AddDefaultPriceDetail(StaticKey As String, Key As Variant, PriceDict As Scripting.Dictionary)
    Dim PriceDetail As clsInvoicePriceDetail
    
    ' Add the default price and log it
    Set PriceDetail = SetPriceDetail(Me.DefaultPriceDict(StaticKey))
    If Not PriceDict.Exists(Key) Then
        PriceDict.Add Key, PriceDetail
        Call m_Logger.LogMessage("AddUnmatchedAnalysesToPriceDict", LogLevelEnum.LogInfo, "Used default price for " & Key)
    Else
        Call m_Logger.LogMessage("AddUnmatchedAnalysesToPriceDict", LogLevelEnum.LogWarning, "Duplicate key in price dictionary: " & Key)
    End If
End Sub
