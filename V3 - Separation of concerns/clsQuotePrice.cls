VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuotePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsQuotePrice
' Author: JT
' Created: 2025-11-24
' Description: Entity representing one price point within a quote
'              line item. Stores regular and bulk prices for a
'              specific processing time. BulkPrice = -1 indicates
'              no bulk pricing available. RegularPrice = -1 should
'              never occur on a valid persisted record.
' DependsOn: clsLoggingSystem, modEnums
' ChangeLog:
'   - 1.4.0 - 2026-04-08 - Refactored for SRP
'   - 1.5.0 - 2026-04-28 - RegularPrice/BulkPrice upgraded to
'                           Currency. FromDict renamed LoadFromDict.
'                           Header updated to new format.
'                           IsValid exposed via Property Get.
'                           SetInternalID added.
'                           Removed ToDict - repo will use get
'                           properties instead
'
' ================================================================
' LAYER: Entity - no DB, no Worksheet, no service dependencies.
'
' LIFECYCLE:
'   1. New clsQuotePrice
'   2. .Initialize(Logger)
'   3. .LoadFromDict(D)       <- rehydration from DB via loader
'
'   ID starts at 0. Sealed by clsQuoteRepository via
'   Friend SetInternalID after INSERT.
'
' Sentinel values:
'   BulkPrice = -1    no bulk pricing available — valid state
'   RegularPrice = -1 should never occur on a persisted record
'
' LoadFromDict keys:
'   id, quote_line_item_id, regular_price, bulk_price,
'   processing_time_id, valid
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_ID As Long
Private m_QuoteLineItemID As Long
Private m_RegularPrice As Currency
Private m_BulkPrice As Currency
Private m_ProcessingTime As ProcessingTimeEnum
Private m_IsValid As Boolean

' ================================================================
' Get Properties
' ================================================================
Public Property Get ID() As Long
    ID = m_ID
End Property
Public Property Get QuoteLineItemID() As Long
    QuoteLineItemID = m_QuoteLineItemID
End Property
Public Property Get ProcessingTime() As ProcessingTimeEnum
    ProcessingTime = m_ProcessingTime
End Property
Public Property Get RegularPrice() As Currency
    RegularPrice = m_RegularPrice
End Property
Public Property Get BulkPrice() As Currency
    BulkPrice = m_BulkPrice
End Property
Public Property Get IsValid() As Boolean
    IsValid = m_IsValid
End Property

' ================================================================
' Initialization
' ================================================================
Public Sub Initialize(Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    If Not Logger.DebugMode Then On Error GoTo EH

    Set m_Logger = Logger

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
    RequireKey D, "quote_line_item_id"
    RequireKey D, "regular_price"
    RequireKey D, "processing_time_id"
    RequireKey D, "valid"
    
    m_ID = NullToDefault(D("id"), 0)
    m_QuoteLineItemID = NullToDefault(D("quote_line_item_id"), 0)
    m_RegularPrice = NullToDefault(D("regular_price"), 0)
    m_BulkPrice = NullToDefault(D("bulk_price"), -1)
    m_ProcessingTime = NullToDefault(D("processing_time_id"), 0)
    m_IsValid = NullToDefault(D("valid"), False)

    Debug.Assert m_RegularPrice >= 0  ' -1 should never occur on a persisted record

    Exit Sub
EH:
    HandleError "LoadFromDict", "Failed on ID: " & D("id")
End Sub

Public Function Validate() As Boolean
    Validate = True
    If m_ID < 0 Then Validate = False
    If m_RegularPrice < 0 Then Validate = False
    If m_BulkPrice < -1 Then Validate = False
    If m_ProcessingTime < ExtendedTime Or m_ProcessingTime > FiveDays Then Validate = False
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
        Err.Raise vbObjectError + 200, "clsAnalysis.LoadFromDict", _
    "Missing required key: " & Key
    End If
End Sub

Private Sub HandleError(MethodName As String, Optional Context As String = "")
    m_Logger.LogError "clsQuotePrice." & MethodName, Err.Number, _
        Err.Description & IIf(Context <> "", " | " & Context, "")
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsQuotePrice." & MethodName, _
            "The object must be initialized before use."
    End If
End Sub

