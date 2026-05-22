VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoicePricingEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoicePricingEngine
' Author: JT
' Created: 2026-03-27
' Description: Engine to sort analyses and assign price
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsTRSubmission, clsTRSample, clsQuote
' ChangeLog:
'   - 0.1.0 - 2026-03-27 - Initial skeleton
Private m_Logger As clsLoggingSystem
Private m_PricingCache As clsInvoicePricingCache
' ==============
' Initialization
' ==============
Public Sub Initialize(PricingCache As clsInvoicePricingCache, Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing
    Debug.Assert Not PricingCache Is Nothing
    Debug.Assert m_PricingCache Is Nothing

    If Not Logger.DebugMode Then On Error GoTo ErrorHandler

    Set m_Logger = Logger
    Set m_PricingCache = PricingCache


    Exit Sub
ErrorHandler:
    Call HandleError("Initialize", Err.Description)
End Sub

' =====================
' Core API
' =====================
Public Function CalculatePricingInformation(Submission As clsTRSubmission) As Scripting.Dictionary


    Dim QuoteColl As Collection
    Dim Q As clsQuote
    Dim RemainingPricingGroups As Scripting.Dictionary
    Dim MatchingPricingGroups As Scripting.Dictionary
    Dim CustomerQuoteCollection As Collection
    Dim DefaultQuoteCollection As Collection

    Debug.Assert Not Submission Is Nothing

    Call AssertInitialized("GetLineItems")

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    If Submission Is Nothing Then
        Err.Raise 910, "clsPricingEngine.GetLineItems", "Submission cannot be Nothing."
    End If

    ' 1) Group samples
    Set RemainingPricingGroups = GroupSubmissionForPricing(Submission)
    Set MatchingPricingGroups = New Scripting.Dictionary

    ' 2) Retrieve quotes
    Set CustomerQuoteCollection = m_PricingCache.RetrieveQuotesByCustomer(Submission.Customer.CustomerID)
    SortQuotesByAnalysisCount CustomerQuoteCollection ' Sort the collection to favour bundles over quotes for individual analyses
    Set DefaultQuoteCollection = m_PricingCache.RetrieveDefaultQuotes

    ' 3) Assign prices
    AssignPrices "Customer quotes", CustomerQuoteCollection, RemainingPricingGroups, MatchingPricingGroups ' Set prices from bundles and single quotes
    AssignPrices "Default quotes", DefaultQuoteCollection, RemainingPricingGroups, MatchingPricingGroups ' Any remaining analyses should be set prices by fall back default quotes

    If RemainingPricingGroups.Count > 0 Then m_Logger.LogMessage "clsInvoicePricingEngine.CalculatePricingInformation", LogLevelEnum.LogWarning, RemainingPricingGroups.Count & " pricing groups with unassigned prices."

    Set CalculatePricingInformation = MatchingPricingGroups

    Exit Function
ErrorHandler:
    Call HandleError("CalculatePricingInformation", Err.Description)
End Function
' ===============
' Private Helpers
' ===============
Public Function GroupSubmissionForPricing(Submission As clsTRSubmission) As Scripting.Dictionary

    Dim AnalysisKey As String
    Dim PricingGroupDict As New Scripting.Dictionary
    Dim Sample As clsTestingRequestSample
    Dim Analysis As clsTestingRequestAnalysis
    Dim GroupKey As String
    Dim Coll As Collection
    Dim PG As clsInvoicePricingGroup

    For Each Sample In Submission.Samples

        ' Build normalized analysis key
        AnalysisKey = BuildAnalysisKey(Sample)

        ' Group by customer + processing time + analysis set
        GroupKey = Submission.Customer.CustomerID & "|" & Sample.ProcessingTime & "|" & AnalysisKey


        If PricingGroupDict.Exists(GroupKey) Then
            Set PG = PricingGroupDict(GroupKey)
            ' Increment quantity and add sample name
            PG.Quantity = PG.Quantity + 1
            PG.Samples.Add Sample.SampleName
            ' --- Logging for traceability ---
            m_Logger.LogMessage "GroupSubmissionForPricing", 1, "Incremented existing line item: " & "GroupKey=" & GroupKey & ", NewQty=" & PG.Quantity & ", AddedSample=" & Sample.SampleName

        Else
            ' First time seeing this analysis + processing combination

            Set PG = New clsInvoicePricingGroup
            Set PG.Samples = New Scripting.Dictionary
            Set PG.Analyses = New Scripting.Dictionary

            ' Add sample
            PG.Samples.Add Sample.SampleName, True

            ' Add all analyses
            For Each Analysis In Sample.Analyses
                PG.Analyses.Add Analysis.AnalysisID, True
            Next Analysis

            PG.Quantity = 1
            PG.ProcessingTime = Sample.ProcessingTime
            PG.CustomerID = Submission.Customer.CustomerID

            PricingGroupDict.Add GroupKey, PG

            ' --- Logging for traceability ---
            m_Logger.LogMessage "GroupSubmissionForPricing", 1, "Created new line item: " & "GroupKey=" & GroupKey & ", Sample=" & Sample.SampleName
        End If
    Next Sample

    Set GroupSubmissionForPricing = PricingGroupDict

End Function
Private Function BuildAnalysisKey(Sample As clsTRSample) As String
    Dim IDs() As String
    Dim i As Long
    Dim Analysis As clsTestingRequestAnalysis

    ReDim IDs(1 To Sample.Analyses.Count)

    i = 1
    For Each Analysis In Sample.Analyses
        IDs(i) = CStr(Format(Analysis.AnalysisID, "0000"))
        i = i + 1
    Next Analysis

    Call QuickSortStrings(IDs, LBound(IDs), UBound(IDs))

    BuildAnalysisKey = Join(IDs, ",")
End Function
Private Sub SortQuotesByAnalysisCount(QuoteColl As Collection)
    Dim i As Long, j As Long
    Dim temp As clsQuote

    ' Bubble sort for simplicity
    For i = 1 To QuoteColl.Count - 1
        For j = i + 1 To QuoteColl.Count
            If QuoteColl(i).Analyses.Count < QuoteColl(j).Analyses.Count Then
                Set temp = QuoteColl(i)
                Set QuoteColl(i) = QuoteColl(j)
                Set QuoteColl(j) = temp
            End If
        Next j
    Next i
End Sub
Private Sub AssignPrices(Source As String, QuoteColl As Collection, RemainingPricingGroups As Scripting.Dictionary, MatchedPricingGroups As Scripting.Dictionary)
    Dim Q As clsQuote
    Dim Key As Variant
    Dim PG As clsInvoicePricingGroup
    Dim KeysToRemove As New Collection

    m_Logger.LogMessage "clsInvoicePricingEngine.AssignPrices", LogLevelEnum.LogDebug, "Assigning prices from " & Source

    For Each Key In RemainingPricingGroups.Keys

        Set PG = RemainingPricingGroups(Key)

        For Each Q In QuoteColl

            If ApplyQuoteToPG(Q, PG, MatchedPricingGroups) Then

                KeysToRemove.Add Key
                Exit For

            End If
        Next Q

    Next Key

    For Each Key In KeysToRemove
        RemainingPricingGroups.Remove Key
    Next Key

End Sub
Private Function ApplyQuoteToPG(Q As clsQuote, PG As clsInvoicePricingGroup, MatchedPricingGroups As Scripting.Dictionary) As Boolean
    Dim AnalysisID As Variant

    ApplyQuoteToPG = False

    ' Check if the quote matches
    If MatchAnalyses(Q, PG) Then
        ' Create new PG with price assigned
        Dim NewPG As clsInvoicePricingGroup
        Set NewPG = CreateNewPG(Q, PG)
        MatchedPricingGroups.Add NewPG, True

        m_Logger.LogMessage "ApplyQuoteToPG", LogLevelEnum.LogInfo, "Applied QuoteID=" & Q.QuoteID & " to Group (CustomerID=" & PG.CustomerID & ", ProcessingTime=" & PG.ProcessingTime & ")" & ". Analyses assigned: " & Join(CollectionKeys(NewPG.Analyses), ", ") & ", Price=" & NewPG.Price

        ' Return True if PG is fully consumed
        If PG.Analyses.Count = 0 Then ApplyQuoteToPG = True

    End If
End Function
Private Function CreateNewPG(Q As clsQuote, SourcePG As clsInvoicePricingGroup) As clsInvoicePricingGroup
    Dim NewPG As clsInvoicePricingGroup
    Dim AnalysisID As Variant

    Set NewPG = New clsInvoicePricingGroup

    NewPG.CustomerID = SourcePG.CustomerID
    NewPG.ProcessingTime = SourcePG.ProcessingTime
    NewPG.Quantity = SourcePG.Quantity
    NewPG.Price = Q.Price(SourcePG.ProcessingTime)

    Set NewPG.Samples = CloneCollection(SourcePG.Samples)

    Set NewPG.Analyses = New Scripting.Dictionary

    For Each AnalysisID In Q.Analyses

        If SourcePG.Analyses.Exists(AnalysisID) Then
            NewPG.Analyses(AnalysisID) = True
            SourcePG.Analyses.Remove AnalysisID
        End If

    Next AnalysisID

    Set CreateNewPG = NewPG

End Function
Private Function MatchAnalyses(Q As clsQuote, PG As clsInvoicePricingGroup) As Boolean
    Dim ID As Variant
    Dim Found As Boolean

    For Each ID In Q.Analyses
        If Not PG.Analyses.Exists(ID) Then
            MatchAnalyses = False
            Exit Function
        End If
    Next ID

    MatchAnalyses = True
End Function
Private Sub QuickSortStrings(Arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String
    Dim temp As String

    i = first
    j = last
    pivot = Arr((first + last) \ 2)

    Do While i <= j

        Do While Arr(i) < pivot
            i = i + 1
        Loop

        Do While Arr(j) > pivot
            j = j - 1
        Loop

        If i <= j Then
            temp = Arr(i)
            Arr(i) = Arr(j)
            Arr(j) = temp

            i = i + 1
            j = j - 1
        End If

    Loop

    If first < j Then QuickSortStrings Arr, first, j
    If i < last Then QuickSortStrings Arr, i, last
End Sub
Private Function CloneCollection(Source As Collection) As Collection
    Dim NewColl As New Collection
    Dim Item As Variant

    For Each Item In Source
        NewColl.Add Item
    Next Item

    Set CloneCollection = NewColl
End Function
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsInvoicePricingEngine." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsInvoicePricingEngine." & MethodName, "The builder must be initialized before use."
    End If
End Sub
