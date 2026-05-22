VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalesOrderLineItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsSalesOrderLineItem
' Author: JT
' Created: 2026-03-25
' Description: Entity representing one line item on a sales order.
'              Stores item name, description, quantity, rate,
'              processing time, and a quote price ID for
'              traceability. Computes Total as Quantity * Rate.
' DependsOn: clsLoggingSystem, modEnums
' ChangeLog:
'   - 1.0.0 - 2026-03-25 - Initial release. Replaces
'                           clsInvoiceLineItem. AnalysisID removed
'                           — traceability runs through QuotePriceID.
'                           Hydration via LoadFromGroup(DTO) only.
'   - 1.0.1 - 2026-05-19 - LoadFromGroup parameter renamed from
'                           PG to Data following DTO rename from
'                           clsPricedLineItemGroup to
'                           clsSalesOrderLineItemData.
'
' ================================================================
' LAYER: Entity — no DB, no Worksheet, no service dependencies.
'
' LIFECYCLE:
'   1. New clsSalesOrderLineItem
'   2. .Initialize(Logger)
'   3. .LoadFromGroup(Data)      <- creation path only
'                                   called by
'                                   clsSalesOrderLineItemBuilder
'
'   Rehydration from DB is handled by the repository constructing
'   a clsSalesOrderLineItemData from the stored row and calling
'   LoadFromGroup. There is no LoadFromDict path.
'
'   ID starts at 0. Sealed by clsSalesOrderLineItemRepository
'   via Friend SetInternalID after INSERT.
'
' Fields:
'   m_ID                (Long)
'   m_Item              (String)         - display name
'   m_Description       (String)         - display description
'   m_Quantity          (Long)
'   m_Rate              (Currency)
'   m_QuotePriceID      (Long)           - traceability link to
'                                          quote price
'   m_ProcessingTime    (ProcessingTimeEnum)
'
' Computed:
'   Total = Quantity * Rate
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_ID As Long
Private m_Item As String
Private m_Description As String
Private m_Quantity As Long
Private m_Rate As Currency
Private m_QuotePriceID As Long
Private m_ProcessingTime As ProcessingTimeEnum

' ================================================================
' Get Properties
' ================================================================
Public Property Get ID() As Long
    ID = m_ID
End Property
Public Property Get Item() As String
    Item = m_Item
End Property
Public Property Get Description() As String
    Description = m_Description
End Property
Public Property Get Quantity() As Long
    Quantity = m_Quantity
End Property
Public Property Get Rate() As Currency
    Rate = m_Rate
End Property
Public Property Get QuotePriceID() As Long
    QuotePriceID = m_QuotePriceID
End Property
Public Property Get ProcessingTime() As ProcessingTimeEnum
    ProcessingTime = m_ProcessingTime
End Property
Public Property Get Total() As Currency
    Total = m_Quantity * m_Rate
End Property

' ================================================================
' Initialization
' ================================================================
Public Sub Initialize(Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    Set m_Logger = Logger
End Sub

' ================================================================
' APIs
' ================================================================
Public Sub LoadFromData(Data As clsSalesOrderLineItemData)
    Debug.Assert Not Data Is Nothing

    AssertInitialized "LoadFromData"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_ID = 0
    m_Item = Data.Item
    m_Description = Data.Description
    m_Quantity = Data.Quantity
    m_Rate = Data.Price
    m_QuotePriceID = Data.QuotePriceID
    m_ProcessingTime = Data.ProcessingTime

    Exit Sub
EH:
    HandleError "LoadFromData"
End Sub

Public Function Validate() As Boolean
    AssertInitialized "Validate"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    Validate = True

    If m_Quantity <= 0 Then Validate = False
    If m_Rate < 0 Then Validate = False
    If m_Item = "" Then Validate = False

    Exit Function
EH:
    HandleError "Validate"
End Function

Friend Sub SetInternalID(NewID As Long)
    Debug.Assert m_ID = 0
    m_ID = NewID
End Sub

' ================================================================
' Private Helpers
' ================================================================
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsSalesOrderLineItem." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsSalesOrderLineItem." & MethodName, "The line item must be initialized before use."
    End If
End Sub
