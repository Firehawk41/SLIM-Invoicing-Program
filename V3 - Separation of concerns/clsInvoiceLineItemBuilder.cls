VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceLineItemBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceLineItemBuilder
' Author: JT
' Created: 2026-03-26
' Description: Builds line item to nest inside sales orders
' DependsOn: clsLoggingSystem, clsInvoiceLineItem, clsInvoicePricingEngine, clsTRSubmission, clsAnalysis, modEnums, clsTRSample
' ChangeLog:
'   - 0.1.0 - 2026-03-26 - Initial skeleton

Private m_Logger As clsLoggingSystem
Private m_PricingEngine As clsInvoicePricingEngine
' ==============
' Initialization
' ==============
Public Sub Initialize(PricingEngine As clsInvoicePricingEngine, Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing
    Debug.Assert Not PricingEngine Is Nothing
    Debug.Assert m_PricingEngine Is Nothing

    If Not Logger.DebugMode Then On Error GoTo ErrorHandler

    Set m_Logger = Logger
    Set m_PricingEngine = PricingEngine

    Exit Sub
ErrorHandler:
    Call HandleError("Initialize", Err.Description)
End Sub

' =====================
' Core API
' =====================
' Input: Submission object
' Output: Collection of fully-populated clsInvoiceLineItem objects with default prices
Public Function GetLineItems(Submission As clsTRSubmission) As Collection
    Dim LineItems As New Collection
    Dim Sample As clsTestingRequestSample
    Dim Analysis As clsTestingRequestAnalysis
    Dim LI As clsInvoiceLineItem
    Dim DefaultPrice As Currency
    Dim LIDict As Scripting.Dictionary

    Debug.Assert Not Submission Is Nothing

    Call AssertInitialized("GetLineItems")

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    If Submission Is Nothing Then
        Err.Raise 910, "clsPricingEngine.GetLineItems", "Submission cannot be Nothing."
    End If

    Set LIDict = m_PricingEngine.GroupSubmissionForPricing(Submission)

    Set LineItems = BuildLineItems(LIDict)



    Set GetLineItems = LineItems
    Exit Function
ErrorHandler:
    Call HandleError("GetLineItems", Err.Description)
End Function

' =====================
' Helpers
' =====================
Private Function GetDefaultPrice(Analysis As clsAnalysis) As Currency
    ' Placeholder: retrieve default price from m_PriceCache or DB
    ' For now, return a dummy value
    GetDefaultPrice = 100
End Function
Private Function BuildLineItemDict(Analysis As clsAnalysis, ProcessingTime As ProcessingTimeEnum, Quantity As Long, Samples As Collection, Price As Currency) As Scripting.Dictionary

    Dim Dict As New Scripting.Dictionary

    Dict("item") = Analysis.ItemName & " - " & ProcessingTime
    Dict("description") = Analysis.Description
    Dict("quantity") = Quantity
    Dict("rate") = Price
    Dict("analysis_id") = Analysis.ID
    Dict("processing_time") = ProcessingTime
    Dict("source_quote_id") = 0
    Set Dict("samples") = Samples

    Set BuildLineItemDict = Dict

End Function
Private Function AggregateLineItemData(Submission As clsTRSubmission) As Scripting.Dictionary


    Dim LIDict As New Scripting.Dictionary
    Dim Sample As clsTRSample
    Dim Analysis As clsAnalysis
    Dim GroupKey As String
    Dim Coll As Collection

    For Each Sample In Submission.Samples
        For Each Analysis In Sample.Analyses
            ' Create a unique key per analysis + processing time + customer
            GroupKey = Submission.Customer.CustomerID & "|" & Analysis.ID & "|" & Sample.ProcessingTime

            If LIDict.Exists(GroupKey) Then
                ' Increment quantity and add sample name
                LIDict(GroupKey)("quantity") = LIDict(GroupKey)("quantity") + 1
                LIDict(GroupKey)("samples").Add Sample.SampleName
                ' --- Logging for traceability ---
                m_Logger.LogMessage "AggregateLineItemData", 1, "Incremented existing line item: " & "GroupKey=" & GroupKey & ", NewQty=" & LIDict(GroupKey)("quantity") & ", AddedSample=" & Sample.SampleName

            Else
                ' First time seeing this analysis + processing combination
                Set Coll = New Collection
                Coll.Add Sample.SampleName

                Dim TempDict As New Scripting.Dictionary
                TempDict("quantity") = 1
                Set TempDict("samples") = Coll
                TempDict("analysis") = Analysis
                TempDict("processing_time") = Sample.ProcessingTime

                LIDict.Add GroupKey, TempDict

                ' --- Logging for traceability ---
                m_Logger.LogMessage "AggregateLineItemData", 1, "Created new line item: " & "GroupKey=" & GroupKey & ", Sample=" & Sample.SampleName
            End If
        Next Analysis
    Next Sample

    Set AggregateLineItemData = LIDict
End Function
Private Function BuildLineItems(LIDict As Scripting.Dictionary) As Scripting.Dictionary

    Dim LI As clsInvoiceLineItem
    Dim Key As Variant

    For Each Key In LIDict.Keys
        Dim Price As Currency
        Price = GetDefaultPrice(LIDict(Key)("analysis")) ' placeholder

        ' Build the dictionary for the line item
        Dim LineItemDict As Scripting.Dictionary
        Set LineItemDict = BuildLineItemDict(LIDict(Key)("analysis"), LIDict(Key)("processing_time"), LIDict(Key)("quantity"), LIDict(Key)("samples"), Price)

        ' Create the line item object
        Set LI = New clsInvoiceLineItem
        LI.Initialize m_Logger
        LI.FromDict LineItemDict

        m_Logger.LogMessage "BuildLineItems", 1, "Built line item: " & "Item=" & LineItemDict("item") & ", Qty=" & LineItemDict("quantity") & ", Rate=" & LineItemDict("rate") & ", Samples=" & JoinCollectionNames(LineItemDict("samples"))

        ' Add to the collection
        LineItems.Add LI
    Next Key

    Set BuildLineItems = LineItems
End Function
Private Function CreateLineItems(LIDict As Scripting.Dictionary) As Collection
    Dim Coll As Collection
    Dim Key As Object

    Set Coll = New Collection

    For Each Key In LIDict.Keys
        Dim LI As clsInvoiceLineItem
        LI.Initialize m_Logger
        LI.FromDict LIDict(Key)
        If Not LI.Validate Then Err.Raise vbObjectError + 1, , "Line Item invalid"
        Coll.Add LI
    Next Key

    Set CreateLineItems = Coll

End Function

Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsInvoiceLineItemBuilder." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsInvoiceLineItemBuilder." & MethodName, "The builder must be initialized before use."
    End If
End Sub
