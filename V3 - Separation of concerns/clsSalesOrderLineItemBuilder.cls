VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalesOrderLineItemBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsSalesOrderLineItemBuilder
' Author: JT
' Created: 2026-04-28
' Description: Application-layer builder. Takes a clsTRSubmission,
'              delegates pricing to clsSalesOrderPricingEngine,
'              resolves analysis IDs to names and descriptions via
'              clsAnalysisService, populates Item and Description
'              on each clsSalesOrderLineItemData, constructs
'              clsSalesOrderLineItem objects via LoadFromGroup.
' DependsOn: clsLoggingSystem, clsSalesOrderPricingEngine,
'            clsAnalysisService, clsTRSubmission,
'            clsSalesOrderLineItem, clsSalesOrderLineItemData,
'            modEnums, modUtilities
' ChangeLog:
'   - 1.0.0 - 2026-04-28 - Initial release. Replaces
'                           clsInvoiceLineItemBuilder.
'
' ================================================================
' LAYER: Application — sits above the SO domain. Calls pricing
'        engine and analysis service. Produces clsSalesOrderLineItem
'        objects for clsSalesOrderBuilder.
'
' ITEM FORMAT:
'   "[Analysis names joined with ' + '] — [ProcessingTime]"
'   e.g. "ICP-MS + Si — Next Day"
'
' DESCRIPTION FORMAT:
'   "[Analysis descriptions joined with ' + '] — [ProcessingTime]"
'   e.g. "36 trace elements by ICP-MS + Silicon by ICP-MS — Next Day"
'
' Processing time string sourced from modUtilities.ProcessingTimeToString.
' Analysis name and description sourced from clsAnalysis via
' clsAnalysisService.LoadAnalysis.
' ================================================================

Private m_Logger            As clsLoggingSystem
Private m_PricingEngine     As clsSalesOrderPricingEngine
Private m_AnalysisSvc       As clsAnalysisService

' ================================================================
' Initialization
' ================================================================
Public Sub Initialize(PricingEngine As clsSalesOrderPricingEngine, _
                      AnalysisSvc As clsAnalysisService, _
                      Logger As clsLoggingSystem)
    Debug.Assert Not PricingEngine Is Nothing
    Debug.Assert Not AnalysisSvc Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    If Not Logger.DebugMode Then On Error GoTo EH

    Set m_Logger = Logger
    Set m_PricingEngine = PricingEngine
    Set m_AnalysisSvc = AnalysisSvc

    Exit Sub
EH:
    HandleError "Initialize"
End Sub

' ================================================================
' ================================================================
' Core API
' ================================================================

Public Function PriceSubmission(Submission As clsTRSubmission) As Collection
    ' Returns Collection of clsSalesOrderLineItemData (priced DTOs)
    Debug.Assert Not Submission Is Nothing
    AssertInitialized "PriceSubmission"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Set PriceSubmission = m_PricingEngine.CalculatePricingInformation(Submission)
    Exit Function
EH:
    HandleError "PriceSubmission", "Failed on submission: " & Submission.FileName
End Function

Public Function BuildFromPricedGroups(PricedGroups As Collection) As Collection
    ' Returns Collection of clsSalesOrderLineItem (entities)
    Debug.Assert Not PricedGroups Is Nothing
    AssertInitialized "BuildFromPricedGroups"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim Results As New Collection
    Dim LID As clsSalesOrderLineItemData
    Dim LI As clsSalesOrderLineItem

    For Each LID In PricedGroups
        BuildItemStrings LID

        Set LI = New clsSalesOrderLineItem
        LI.Initialize m_Logger
        LI.LoadFromData LID

        If Not LI.Validate Then
            Err.Raise vbObjectError + 1, _
                "clsSalesOrderLineItemBuilder.BuildFromPricedGroups", _
                "Line item failed validation. Item=" & LID.Item & _
                ", Qty=" & LID.Quantity & ", Rate=" & LID.Price
        End If

        Results.Add LI

        m_Logger.LogMessage _
            "clsSalesOrderLineItemBuilder.BuildFromPricedGroups", _
            LogLevelEnum.LogDebug, _
            "Built line item: Item=" & LID.Item & _
            ", Qty=" & LID.Quantity & _
            ", Rate=" & LID.Price & _
            ", QuotePriceID=" & LID.QuotePriceID
    Next LID

    Set BuildFromPricedGroups = Results
    Exit Function
EH:
    HandleError "BuildFromPricedGroups", "Failed on priced groups"
End Function

' ================================================================
' Private Helpers
' ================================================================
Private Sub BuildItemStrings(LID As clsSalesOrderLineItemData)
    ' Resolves analysis IDs on LID.Analyses to names and descriptions.
    ' Populates LID.Item and LID.Description.
    ' Format: "[names joined ' + '] — [ProcessingTime]"

    Dim Names As New Collection
    Dim Descriptions As New Collection
    Dim AnalysisID As Variant
    Dim A As clsAnalysis

    For Each AnalysisID In LID.Analyses
        Set A = m_AnalysisSvc.LoadAnalysis(CLng(AnalysisID))

        If A Is Nothing Then
            Err.Raise vbObjectError + 2, _
                "clsSalesOrderLineItemBuilder.BuildItemStrings", _
                "Analysis ID not found: " & AnalysisID
        End If

        Names.Add A.Name
        Descriptions.Add A.Description
    Next AnalysisID

    Dim ProcessingTimeStr As String
    ProcessingTimeStr = ProcessingTimeToString(LID.ProcessingTime)

    LID.Item = JoinCollection(Names, " + ") & " - " & ProcessingTimeStr
    LID.Description = JoinCollection(Descriptions, " + ") & " - " & ProcessingTimeStr
End Sub

Private Function JoinCollection(Coll As Collection, Separator As String) As String
    Dim Result As String
    Dim Item As Variant
    Dim First As Boolean
    First = True

    For Each Item In Coll
        If First Then
            Result = CStr(Item)
            First = False
        Else
            Result = Result & Separator & CStr(Item)
        End If
    Next Item

    JoinCollection = Result
End Function

Private Sub HandleError(MethodName As String, Optional Context As String = "")
    m_Logger.LogError "clsSalesOrderLineItemBuilder." & MethodName, _
        Err.Number, _
        Err.Description & IIf(Context <> "", " | " & Context, "")
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsSalesOrderLineItemBuilder." & MethodName, _
            "The builder must be initialized before use."
    End If
End Sub

