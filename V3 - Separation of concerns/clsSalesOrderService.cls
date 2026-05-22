VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalesOrderService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsSalesOrderService
' Author: JT
' Created: 2026-04-28
' Description: Service and sole public API for the sales order
'              domain. Coordinates clsSalesOrderRepository and
'              clsSalesOrderLineItemRepository. Owns both the
'              creation and rehydration lifecycle paths.
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsSalesOrder,
'            clsSalesOrderData, clsSalesOrderLineItemData,
'            clsSalesOrderRepository,
'            clsSalesOrderLineItemRepository
' ChangeLog:
'   - 1.0.0 - 2026-04-28 - Initial release
'
' ================================================================
' LAYER: Service - sole public API for the sales order domain.
'        No cache — sales orders are transactional records fetched
'        on demand, not reference data loaded at startup.
'
' LIFECYCLE — Creation path:
'   1. Builder produces clsSalesOrderData with populated LineItems
'   2. CreateFromData validates, persists, seals ID, returns entity
'
' LIFECYCLE — Rehydration path:
'   1. LoadSalesOrder fetches header via repo -> clsSalesOrderData
'   2. Fetches line items via line item repo -> Collection of
'      clsSalesOrderLineItemData
'   3. Constructs and returns fully hydrated clsSalesOrder
'
' All writes are transactional — owned by clsSalesOrderRepository.
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_Repo As clsSalesOrderRepository
Private m_LineItemRepo As clsSalesOrderLineItemRepository

' ================================================================
' Initialization
' ================================================================
Public Sub Initialize(AccessDB As clsAccessDatabase, Logger As clsLoggingSystem)
    Debug.Assert Not AccessDB Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    If Not Logger.DebugMode Then On Error GoTo EH

    Set m_Logger = Logger

    Set m_LineItemRepo = New clsSalesOrderLineItemRepository
    m_LineItemRepo.Initialize AccessDB, Logger

    Set m_Repo = New clsSalesOrderRepository
    m_Repo.Initialize AccessDB, m_LineItemRepo, Logger

    Exit Sub
EH:
    HandleError "Initialize"
End Sub

' ================================================================
' APIs
' ================================================================
Public Function CreateFromData(Data As clsSalesOrderData) As clsSalesOrder
    Debug.Assert Not Data Is Nothing
    AssertInitialized "CreateFromData"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim SO As New clsSalesOrder
    SO.Initialize m_Logger
    SO.LoadFromData Data

    ' Add line items to entity before validation and persist
    If Not Data.LineItems Is Nothing Then
        Dim LI As clsSalesOrderLineItem
        For Each LI In Data.LineItems
            SO.AddLineItem LI
        Next LI
    End If

    ' Validate after line items are attached
    If Not SO.Validate Then
        Err.Raise vbObjectError + 1, "clsSalesOrderService.CreateFromData", _
            "Sales order failed validation before persist."
    End If

    ' Note: cross-object transaction coordination requires refactoring
    ' clsAccessDatabase to expose explicit transaction control.
    ' Deferred to V4 — SQLAlchemy session management handles this natively.

    Dim NewID As Long
    NewID = m_Repo.Insert(SO)

    Set CreateFromData = SO
    Exit Function
EH:
    HandleError "CreateFromData"
End Function

Public Function LoadSalesOrder(SalesOrderID As Long) As clsSalesOrder
    Debug.Assert SalesOrderID > 0

    AssertInitialized "LoadSalesOrder"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim Data As clsSalesOrderData
    Set Data = m_Repo.SelectByID(SalesOrderID)

    If Data Is Nothing Then
        Set LoadSalesOrder = Nothing
        Exit Function
    End If

    Dim SO As New clsSalesOrder
    SO.Initialize m_Logger
    SO.LoadFromData Data
    SO.SetInternalID SalesOrderID

    Dim LineItems As Collection
    Set LineItems = m_LineItemRepo.SelectBySalesOrderID(SalesOrderID)

    Dim LID As clsSalesOrderLineItemData
    For Each LID In LineItems
        Dim LI As New clsSalesOrderLineItem
        LI.Initialize m_Logger
        LI.LoadFromData LID
        If LID.StoredID > 0 Then LI.SetInternalID LID.StoredID
        SO.AddLineItem LI
    Next LID

    Set LoadSalesOrder = SO
    Exit Function
EH:
    HandleError "LoadSalesOrder", "Failed on ID: " & SalesOrderID
End Function

Public Function LoadByDateRange(StartDate As Date, EndDate As Date) As Collection
    AssertInitialized "LoadByDateRange"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    Dim Results As New Collection
    Dim DataItems As Collection
    Set DataItems = m_Repo.SelectByDateRange(StartDate, EndDate)

    Dim Data As clsSalesOrderData
    For Each Data In DataItems
        Dim SO As clsSalesOrder
        Set SO = LoadSalesOrder(Data.SalesOrderID)
        If Not SO Is Nothing Then Results.Add SO
    Next Data

    Set LoadByDateRange = Results
    Exit Function
EH:
    HandleError "LoadByDateRange", "Failed on range: " & StartDate & " to " & EndDate
End Function

Public Sub SaveSalesOrder(SO As clsSalesOrder)
    Debug.Assert Not SO Is Nothing
    Debug.Assert SO.ID > 0

    AssertInitialized "SaveSalesOrder"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    If Not SO.Validate Then
        Err.Raise vbObjectError + 1, "clsSalesOrderService.SaveSalesOrder", _
            "Sales order failed validation before save."
    End If

    m_Repo.Update SO
    Exit Sub
EH:
    HandleError "SaveSalesOrder", "Failed on ID: " & SO.ID
End Sub

Public Sub DeleteSalesOrder(SalesOrderID As Long)
    Debug.Assert SalesOrderID > 0

    AssertInitialized "DeleteSalesOrder"

    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_Repo.Delete SalesOrderID
    Exit Sub
EH:
    HandleError "DeleteSalesOrder", "Failed on ID: " & SalesOrderID
End Sub

' ================================================================
' Private Helpers
' ================================================================
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    m_Logger.LogError "clsSalesOrderService." & MethodName, Err.Number, _
        Err.Description & IIf(Context <> "", " | " & Context, "")
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsSalesOrderService." & MethodName, _
            "The service must be initialized before use."
    End If
End Sub

