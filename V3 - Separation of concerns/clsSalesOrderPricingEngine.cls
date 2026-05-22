VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalesOrderPricingEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsSalesOrderPricingEngine
' Author: JT
' Created: 2026-03-27
' Description: Pricing orchestrator for the SO pipeline. Takes a
'              clsTRSubmission, groups samples into
'              clsSalesOrderLineItemData objects, matches each
'              group against valid quotes from clsQuoteService,
'              assigns prices, and returns a Collection of fully
'              priced clsSalesOrderLineItemData objects.
' DependsOn: clsLoggingSystem, clsQuoteService, clsTRSubmission,
'            clsTRSample, clsSalesOrderLineItemData, modEnums
' ChangeLog:
'   - 0.1.0 - 2026-03-27 - Initial skeleton
'   - 1.0.0 - 2026-04-28 - Full rewrite. Rewired to clsQuoteService.
'                           clsInvoicePricingCache removed.
'                           Stale type references fixed.
'                           GroupSubmissionForPricing made private.
'                           Return type changed to Collection.
'                           Quote/line item/price validity filtering
'                           added. Bulk price logic implemented.
'                           QuotePriceID stored on matched group.
'                           Raises on unmatched groups.
'   - 1.0.1 - 2026-05-19 - Added DefaultPaymentTerms parameter to
'                           Initialize. Used as fallback when quote
'                           has no payment terms set.
'
' ================================================================
' LAYER: Application — sits above both the quote domain and the
'        SO pipeline. Consumes clsQuoteService for quote data.
'        Produces clsSalesOrderLineItemData for the line item
'        builder.
'
' MATCHING LOGIC:
'   1. Group submission samples by CustomerID|ProcessingTime|
'      SortedAnalysisIDs — one group per unique combination
'   2. Retrieve valid, in-date customer quotes from service
'   3. Sort quotes to favour bundles (more analyses) over singles
'   4. Match groups against customer quotes first, then defaults
'   5. For each match, assign price from clsQuotePrice:
'      - BulkPrice if Quantity >= BulkPriceMin and BulkPrice >= 0
'      - RegularPrice otherwise
'   6. Raise if any group remains unmatched after both passes
'
' VALIDITY FILTER (applied before matching):
'   clsQuote:         IsValid = True, EffectiveDate <= Today
'                     <= ExpireDate
'   clsQuoteLineItem: IsValid = True
'   clsQuotePrice:    IsValid = True
' ================================================================
Private m_Logger    As clsLoggingSystem
Private m_QuoteSvc  As clsQuoteService
Private m_DefaultPaymentTerms As Long

' ================================================================
' Initialization
' ================================================================
Public Sub Initialize(QuoteSvc As clsQuoteService, _
                      DefaultPaymentTerms As Long, _
                      Logger As clsLoggingSystem)
    Debug.Assert Not QuoteSvc Is Nothing
    Debug.Assert DefaultPaymentTerms > 0
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    Set m_Logger = Logger
    Set m_QuoteSvc = QuoteSvc
    m_DefaultPaymentTerms = DefaultPaymentTerms

End Sub

' ================================================================
' Core API
' ================================================================
Public Function CalculatePricingInformation(Submission As clsTRSubmission) As Collection
    Debug.Assert Not Submission Is Nothing

    AssertInitialized "CalculatePricingInformation"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    ' 1. Group samples
    Dim RemainingGroups As Scripting.Dictionary
    Set RemainingGroups = GroupSubmissionForPricing(Submission)

    Dim MatchedGroups As New Collection

    ' 2. Retrieve and filter valid customer quotes
    Dim CustomerQuotes As Collection
    Set CustomerQuotes = FilterValidQuotes( _
        m_QuoteSvc.GetByCustomerID(Submission.CustomerID))
    SortQuotesByLineItemAnalysisCount CustomerQuotes

    ' 3. Retrieve and filter valid default quotes
    Dim DefaultQuotes As Collection
    Set DefaultQuotes = FilterValidQuotes( _
        m_QuoteSvc.AllDefaultQuotes())
    SortQuotesByLineItemAnalysisCount DefaultQuotes

    ' 4. Assign prices — customer quotes first, defaults as fallback
    AssignPrices "Customer quotes", CustomerQuotes, _
                 RemainingGroups, MatchedGroups
    AssignPrices "Default quotes", DefaultQuotes, _
                 RemainingGroups, MatchedGroups

    ' 5. Raise if any groups remain unmatched
    If RemainingGroups.Count > 0 Then
        Err.Raise vbObjectError + 1, _
            "clsSalesOrderPricingEngine.CalculatePricingInformation", _
            RemainingGroups.Count & " pricing group(s) could not be " & _
            "matched to any quote. CustomerID=" & Submission.CustomerID & _
            ". Check that all analyses have valid quotes."
    End If
    
    ' 6. Merge identical line items
    m_Logger.LogMessage _
        "clsSalesOrderPricingEngine.CalculatePricingInformation", _
        LogLevelEnum.LogDebug, _
        "Pre-merge line item count: " & MatchedGroups.Count

    Dim MergedGroups As Collection
    Set MergedGroups = MergeMatchedGroups(MatchedGroups)

    m_Logger.LogMessage _
        "clsSalesOrderPricingEngine.CalculatePricingInformation", _
        LogLevelEnum.LogInfo, _
        "Post-merge line item count: " & MergedGroups.Count & _
        " (from " & MatchedGroups.Count & " matched groups)"
        
    Set CalculatePricingInformation = MergedGroups
    Exit Function
EH:
    HandleError "CalculatePricingInformation", _
        "Failed on submission: " & Submission.FileName
End Function

' ================================================================
' Private — Grouping
' ================================================================
Private Function GroupSubmissionForPricing(Submission As clsTRSubmission) _
                                          As Scripting.Dictionary
    Dim Groups As New Scripting.Dictionary
    Dim Sample As clsTRSample
    Dim AnalysisKey As String
    Dim GroupKey As String
    Dim ExistingData As clsSalesOrderLineItemData
    Dim NewData As clsSalesOrderLineItemData
    Dim AnalysisID As Variant

    For Each Sample In Submission.Samples
        AnalysisKey = BuildAnalysisKey(Sample)
        GroupKey = Submission.CustomerID & "|" & _
                   Sample.ProcessingTime & "|" & AnalysisKey

        If Groups.Exists(GroupKey) Then
            Set ExistingData = Groups(GroupKey)
            ExistingData.Quantity = ExistingData.Quantity + 1
            ExistingData.Samples.Add Sample.SampleName

        Else
            Set NewData = New clsSalesOrderLineItemData
            Set NewData.Samples = New Collection
            Set NewData.Analyses = New Scripting.Dictionary
            NewData.Samples.Add Sample.SampleName
            NewData.Quantity = 1
            NewData.ProcessingTime = Sample.ProcessingTime
            NewData.CustomerID = Submission.CustomerID
            For Each AnalysisID In Sample.AnalysisIDs
                NewData.Analyses.Add CLng(AnalysisID), True
            Next AnalysisID
            Groups.Add GroupKey, NewData

        End If
    Next Sample

    Set GroupSubmissionForPricing = Groups
End Function

Private Function BuildAnalysisKey(Sample As clsTRSample) As String
    Dim IDs() As String
    Dim i As Long
    Dim AnalysisID As Variant

    ReDim IDs(1 To Sample.AnalysisIDs.Count)

    i = 1
    For Each AnalysisID In Sample.AnalysisIDs
        IDs(i) = Format(CLng(AnalysisID), "0000")
        i = i + 1
    Next AnalysisID

    QuickSortStrings IDs, LBound(IDs), UBound(IDs)
    BuildAnalysisKey = Join(IDs, ",")
End Function
Private Function MergeMatchedGroups(Groups As Collection) As Collection
    Dim Merged As New Scripting.Dictionary
    Dim Data As clsSalesOrderLineItemData
    Dim Key As String
    Dim SampleName As Variant

    For Each Data In Groups
        Key = CStr(Data.QuotePriceID)

        If Merged.Exists(Key) Then
            Dim Existing As clsSalesOrderLineItemData
            Set Existing = Merged(Key)
            Existing.Quantity = Existing.Quantity + Data.Quantity
            For Each SampleName In Data.Samples
                Existing.Samples.Add SampleName
            Next SampleName

            m_Logger.LogMessage _
                "clsSalesOrderPricingEngine.MergeMatchedGroups", _
                LogLevelEnum.LogDebug, _
                "Merged: " & Data.Item & " qty " & Data.Quantity & _
                " into existing qty " & Existing.Quantity - Data.Quantity & _
                " -> new qty " & Existing.Quantity
        Else
            Set Merged(Key) = Data
        End If
    Next Data

    Dim Results As New Collection
    Dim K As Variant
    For Each K In Merged.Keys
        Results.Add Merged(K)
    Next K

    m_Logger.LogMessage _
        "clsSalesOrderPricingEngine.MergeMatchedGroups", _
        LogLevelEnum.LogDebug, _
        "Merge complete: " & Groups.Count & " groups -> " & Results.Count & " line items"

    Set MergeMatchedGroups = Results
End Function
' ================================================================
' Private — Filtering
' ================================================================
Private Function FilterValidQuotes(Quotes As Collection) As Collection
    Dim Results As New Collection
    Dim Q As clsQuote
    Dim Today As Date
    Today = VBA.Date

    For Each Q In Quotes
        If Q.IsValid Then
            If Q.EffectiveDate <= Today And Q.ExpireDate >= Today Then
                Results.Add Q
            End If
        End If
    Next Q

    Set FilterValidQuotes = Results
End Function

' ================================================================
' Private — Sorting
' ================================================================
Private Sub SortQuotesByLineItemAnalysisCount(Quotes As Collection)
    ' Bubble sort — favours quotes whose line items cover more
    ' analyses. Ensures bundles are matched before single analyses.
    Dim i As Long
    Dim j As Long
    Dim MaxA As Long
    Dim MaxB As Long
    Dim Temp As clsQuote

    For i = 1 To Quotes.Count - 1
        For j = i + 1 To Quotes.Count
            MaxA = MaxLineItemAnalysisCount(Quotes(i))
            MaxB = MaxLineItemAnalysisCount(Quotes(j))
            If MaxA < MaxB Then
                Set Temp = Quotes(i)
                Set Quotes(i) = Quotes(j)
                Set Quotes(j) = Temp
            End If
        Next j
    Next i
End Sub

Private Function MaxLineItemAnalysisCount(Q As clsQuote) As Long
    Dim LI As clsQuoteLineItem
    Dim MaxCount As Long
    MaxCount = 0
    For Each LI In Q.LineItems
        If LI.AnalysisIDs.Count > MaxCount Then
            MaxCount = LI.AnalysisIDs.Count
        End If
    Next LI
    MaxLineItemAnalysisCount = MaxCount
End Function

' ================================================================
' Private — Price Assignment
' ================================================================
Private Sub AssignPrices(Source As String, _
                         Quotes As Collection, _
                         RemainingDTOs As Scripting.Dictionary, _
                         MatchedDTOs As Collection)
    m_Logger.LogMessage _
        "clsSalesOrderPricingEngine.AssignPrices", _
        LogLevelEnum.LogDebug, _
        "Assigning prices from " & Source

    Dim KeysToRemove As New Collection
    Dim Key As Variant

    For Each Key In RemainingDTOs.Keys
        Dim Data As clsSalesOrderLineItemData
        Set Data = RemainingDTOs(Key)

        Dim Q As clsQuote
        For Each Q In Quotes
            If TryApplyQuoteToData(Q, Data, MatchedDTOs) Then
                KeysToRemove.Add Key
                Exit For
            End If
        Next Q
    Next Key

    For Each Key In KeysToRemove
        RemainingDTOs.Remove Key
    Next Key
End Sub

Private Function TryApplyQuoteToData(Q As clsQuote, _
                                      Data As clsSalesOrderLineItemData, _
                                      MatchedDTOs As Collection) As Boolean
    TryApplyQuoteToData = False
    Dim P As clsQuotePrice
    Dim MatchedData As clsSalesOrderLineItemData
    Dim AnalysisID As Variant
    
    m_Logger.LogMessage "clsSalesOrderPricingEngine.TryApplyQuoteToData", _
        LogLevelEnum.LogDebug, _
        "Evaluating quote ID=" & Q.ID & " with " & Q.LineItems.Count & " line items"
    
    Dim LI As clsQuoteLineItem
    For Each LI In Q.LineItems

        If Not LI.IsValid Then
            m_Logger.LogMessage "clsSalesOrderPricingEngine.TryApplyQuoteToGroup", _
                LogLevelEnum.LogDebug, "LineItem " & LI.ID & " skipped: not valid"
            GoTo NextLineItem
        End If

        If Not MatchAnalyses(LI, Data) Then
            m_Logger.LogMessage "clsSalesOrderPricingEngine.TryApplyQuoteToGroup", _
                LogLevelEnum.LogDebug, "LineItem " & LI.ID & " skipped: analyses don't match"
            GoTo NextLineItem
        End If

        ' Find the price for this processing time
        Set P = LI.GetPrice(Data.ProcessingTime)
        If P Is Nothing Then
            m_Logger.LogMessage "clsSalesOrderPricingEngine.TryApplyQuoteToGroup", _
                LogLevelEnum.LogDebug, "LineItem " & LI.ID & " skipped: no price for processing time " & Data.ProcessingTime
            GoTo NextLineItem
        ElseIf Not P.IsValid Then
            m_Logger.LogMessage "clsSalesOrderPricingEngine.TryApplyQuoteToGroup", _
                LogLevelEnum.LogDebug, "LineItem " & LI.ID & " skipped: no price for processing time " & Data.ProcessingTime
            GoTo NextLineItem
        End If

        ' Apply price — bulk supersedes regular if quantity qualifies
        Set MatchedData = New clsSalesOrderLineItemData
        CopyGroupData Data, MatchedData

        If LI.BulkPriceMin > 0 And _
           Data.Quantity >= LI.BulkPriceMin And _
           P.BulkPrice >= 0 Then
            MatchedData.Price = P.BulkPrice
        Else
            MatchedData.Price = P.RegularPrice
        End If

        MatchedData.QuotePriceID = P.ID
        MatchedData.PaymentTerms = IIf(Q.PaymentTerms > 0, Q.PaymentTerms, m_DefaultPaymentTerms)

        ' Remove matched analyses from Data's remaining set
        For Each AnalysisID In LI.AnalysisIDs
            If Data.Analyses.Exists(CLng(AnalysisID)) Then
                Data.Analyses.Remove CLng(AnalysisID)
                MatchedData.Analyses.Add (CLng(AnalysisID)), True
            End If
        Next AnalysisID

        MatchedDTOs.Add MatchedData

        m_Logger.LogMessage _
            "clsSalesOrderPricingEngine.TryApplyQuoteToData", _
            LogLevelEnum.LogInfo, _
            "Matched QuotePriceID=" & P.ID & _
            " to Data (CustomerID=" & Data.CustomerID & _
            ", ProcessingTime=" & Data.ProcessingTime & _
            ", Price=" & MatchedData.Price & ")"

        ' Return True only if Data is fully consumed
        If Data.Analyses.Count = 0 Then
            TryApplyQuoteToData = True
            Exit Function
        End If


NextLineItem:
    Next LI
End Function

Private Function MatchAnalyses(LI As clsQuoteLineItem, _
                               Data As clsSalesOrderLineItemData) As Boolean
    ' All analyses on the line item must exist in the Data
    Dim AnalysisID As Variant
    For Each AnalysisID In LI.AnalysisIDs
        If Not Data.Analyses.Exists(CLng(AnalysisID)) Then
            MatchAnalyses = False
            Exit Function
        End If
    Next AnalysisID
    MatchAnalyses = True
End Function

Private Sub CopyGroupData(Source As clsSalesOrderLineItemData, _
                          Target As clsSalesOrderLineItemData)
    Target.Quantity = Source.Quantity
    Target.ProcessingTime = Source.ProcessingTime
    Target.CustomerID = Source.CustomerID
    Set Target.Samples = Source.Samples
    Set Target.Analyses = New Scripting.Dictionary
End Sub

' ================================================================
' Private — Sorting Helper
' ================================================================
Private Sub QuickSortStrings(Arr() As String, _
                             ByVal First As Long, _
                             ByVal Last As Long)
    Dim i As Long
    Dim j As Long
    Dim Pivot As String
    Dim Temp As String

    i = First
    j = Last
    Pivot = Arr((First + Last) \ 2)

    Do While i <= j
        Do While Arr(i) < Pivot
            i = i + 1
        Loop
        Do While Arr(j) > Pivot
            j = j - 1
        Loop
        If i <= j Then
            Temp = Arr(i)
            Arr(i) = Arr(j)
            Arr(j) = Temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If First < j Then QuickSortStrings Arr, First, j
    If i < Last Then QuickSortStrings Arr, i, Last
End Sub

' ================================================================
' Private Helpers
' ================================================================
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    m_Logger.LogError "clsSalesOrderPricingEngine." & MethodName, _
        Err.Number, _
        Err.Description & IIf(Context <> "", " | " & Context, "")
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsSalesOrderPricingEngine." & MethodName, _
            "The engine must be initialized before use."
    End If
End Sub

