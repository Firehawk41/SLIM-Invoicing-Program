VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuoteLineItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsQuoteLineItem
' Author: JT
' Created: 2025-12-01
' Description: Entity representing one line item within a quote.
'              Holds request type, bulk pricing minimum, optional
'              chemical ID (nullable — line item may apply to all
'              chemicals or one specific chemical), collections of
'              analysis IDs and element IDs, and a Dictionary of
'              clsQuotePrice objects keyed by ProcessingTimeEnum.
' DependsOn: clsLoggingSystem, clsChemicalService, clsQuotePrice,
'            modEnums
' ChangeLog:
'   - 2.2.0 - 2026-04-14 - Removed analysis and chemical managers
'   - 2.3.0 - 2026-04-28 - FromDict renamed LoadFromDict. ToDict
'                           removed — repo reads properties directly.
'                           ChemicalSvc injected for ChemicalName.
'                           Full Property Get surface added.
'                           AnalysisIDs/ElementIDs exposed as
'                           defensive copies. AddAnalysisID/
'                           AddElementID silently deduplicate.
'                           ValidateStateDebug removed — rules moved
'                           to Validate. Validate returns Boolean.
'                           SetInternalID added. RequireKey added.
'                           Header updated to new format.
'
' ================================================================
' LAYER: Entity - no DB, no Worksheet.
'        Depends on clsChemicalService for ChemicalName only.
'
' LIFECYCLE:
'   1. New clsQuoteLineItem
'   2. .Initialize(ChemicalSvc, Logger)
'   3. .LoadFromDict(D)           <- scalar fields only
'   4. .AddAnalysisID(ID) x N     <- loader attaches from junction table
'   5. .AddElementID(ID) x N      <- loader attaches from junction table
'   6. .UpsertPrice(P) x N        <- loader attaches prices
'
'   Steps 3-6 performed by clsQuoteLoader.
'
'   ID starts at 0. Sealed by clsQuoteRepository via
'   Friend SetInternalID after INSERT.
'
'   ChemicalID is a nullable Variant. Null means the line item
'   applies to all chemicals. Non-null means it applies to one
'   specific chemical (e.g. water at a different price).
'
'   Repo reads properties directly — no ToDict.
'
' LoadFromDict keys:
'   id, quote_id, request_type_id, bulk_min, chemical_id, valid
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_ChemicalSvc As clsChemicalService
Private m_ID As Long
Private m_QuoteID As Long
Private m_RequestType As TestingRequestTypeEnum
Private m_BulkPriceMin As Long
Private m_IsValid As Boolean
Private m_ChemicalID As Variant
Private m_AnalysisIDs As Collection
Private m_ElementIDs As Collection
Private m_Prices As Scripting.Dictionary

' ================================================================
' Get Properties
' ================================================================
Public Property Get ID() As Long
    ID = m_ID
End Property
Public Property Get QuoteID() As Long
    QuoteID = m_QuoteID
End Property
Public Property Get RequestType() As TestingRequestTypeEnum
    RequestType = m_RequestType
End Property
Public Property Get BulkPriceMin() As Long
    BulkPriceMin = m_BulkPriceMin
End Property
Public Property Get ChemicalID() As Variant
    ChemicalID = m_ChemicalID
End Property
Public Property Get ChemicalName() As String
    If IsNull(m_ChemicalID) Then
        ChemicalName = ""
    Else
        Dim C As clsChemical
        Set C = m_ChemicalSvc.LoadChemical(CLng(m_ChemicalID))
        If C Is Nothing Then
            ChemicalName = ""
        Else
            ChemicalName = C.Name
        End If
    End If
End Property
Public Property Get IsValid() As Boolean
    IsValid = m_IsValid
End Property
Public Property Get AllPrices() As Scripting.Dictionary
    Set AllPrices = m_Prices
End Property
Public Property Get AnalysisIDs() As Collection
    Dim Copy As New Collection
    Dim ID As Variant
    For Each ID In m_AnalysisIDs
        Copy.Add ID
    Next ID
    Set AnalysisIDs = Copy
End Property
Public Property Get ElementIDs() As Collection
    Dim Copy As New Collection
    Dim ID As Variant
    For Each ID In m_ElementIDs
        Copy.Add ID
    Next ID
    Set ElementIDs = Copy
End Property

' ================================================================
' Initialization
' ================================================================
Public Sub Initialize(ChemicalSvc As clsChemicalService, Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert Not ChemicalSvc Is Nothing
    Debug.Assert m_Logger Is Nothing

    If Not Logger.DebugMode Then On Error GoTo EH

    Set m_Logger = Logger
    Set m_ChemicalSvc = ChemicalSvc
    Set m_AnalysisIDs = New Collection
    Set m_ElementIDs = New Collection
    Set m_Prices = New Scripting.Dictionary

    Exit Sub
EH:
    HandleError "Initialize"
End Sub

' ================================================================
' APIs
' ================================================================
Public Sub LoadFromDict(D As Scripting.Dictionary)
    Debug.Assert Not D Is Nothing

    AssertInitialized "LoadFromDict"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    RequireKey D, "id"
    RequireKey D, "quote_id"
    RequireKey D, "request_type_id"
    RequireKey D, "bulk_min"
    RequireKey D, "valid"

    m_ID = NullToDefault(D("id"), 0)
    m_QuoteID = NullToDefault(D("quote_id"), 0)
    m_RequestType = NullToDefault(D("request_type_id"), 0)
    m_BulkPriceMin = NullToDefault(D("bulk_min"), 0)
    m_IsValid = NullToDefault(D("valid"), False)

    If D.Exists("chemical_id") Then
        If IsNull(D("chemical_id")) Then
            m_ChemicalID = Null
        Else
            m_ChemicalID = CLng(D("chemical_id"))
        End If
    Else
        m_ChemicalID = Null
    End If

    Exit Sub
EH:
    HandleError "LoadFromDict", "Failed on ID: " & D("id")
End Sub

Public Function GetPrice(TimeType As ProcessingTimeEnum) As clsQuotePrice
    Debug.Assert TimeType >= ProcessingTimeEnum.ExtendedTime And _
                 TimeType <= ProcessingTimeEnum.FiveDays

    AssertInitialized "GetPrice"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    If m_Prices.Exists(TimeType) Then
        Set GetPrice = m_Prices(TimeType)
    Else
        Set GetPrice = Nothing
    End If

    Exit Function
EH:
    HandleError "GetPrice", "TimeType=" & TimeType
End Function

Public Sub UpsertPrice(P As clsQuotePrice)
    Debug.Assert Not P Is Nothing

    AssertInitialized "UpsertPrice"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    Set m_Prices(P.ProcessingTime) = P

    Exit Sub
EH:
    HandleError "UpsertPrice", "Failed on ProcessingTime: " & P.ProcessingTime
End Sub

Public Sub AddAnalysisID(AnalysisID As Long)
    AssertInitialized "AddAnalysisID"

    Dim ID As Variant
    For Each ID In m_AnalysisIDs
        If ID = AnalysisID Then Exit Sub
    Next ID
    m_AnalysisIDs.Add AnalysisID
End Sub

Public Sub AddElementID(ElementID As Long)
    AssertInitialized "AddElementID"

    Dim ID As Variant
    For Each ID In m_ElementIDs
        If ID = ElementID Then Exit Sub
    Next ID
    m_ElementIDs.Add ElementID
End Sub

Public Sub RemovePrice(TimeType As ProcessingTimeEnum)
    AssertInitialized "RemovePrice"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    If m_Prices.Exists(TimeType) Then m_Prices.Remove TimeType

    Exit Sub
EH:
    HandleError "RemovePrice", "TimeType=" & TimeType
End Sub

Public Function Validate() As Boolean
    AssertInitialized "Validate"

    Validate = True

    If m_ID < 0 Then Validate = False
    If m_QuoteID < 0 Then Validate = False
    If m_RequestType < TestingRequestTypeEnum.RequestType_Min Or _
       m_RequestType > TestingRequestTypeEnum.RequestType_Max Then Validate = False
    If m_BulkPriceMin < -1 Then Validate = False
    If Not IsNull(m_ChemicalID) Then
        If Not IsNumeric(m_ChemicalID) Then Validate = False
    End If
    If m_Prices.Count = 0 Then Validate = False
End Function

Friend Sub SetInternalID(NewID As Long)
    Debug.Assert m_ID = 0
    m_ID = NewID
End Sub

' ================================================================
' Private Helpers
' ================================================================
Private Sub RequireKey(D As Scripting.Dictionary, Key As String)
    Debug.Assert D.Exists(Key)
    If Not D.Exists(Key) Then
        Err.Raise vbObjectError + 200, "clsQuoteLineItem.LoadFromDict", _
            "Missing required key: " & Key
    End If
End Sub

Private Sub HandleError(MethodName As String, Optional Context As String = "")
    m_Logger.LogError "clsQuoteLineItem." & MethodName, Err.Number, _
        Err.Description & IIf(Context <> "", " | " & Context, "")
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsQuoteLineItem." & MethodName, _
            "The object must be initialized before use."
    End If
End Sub

