VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSalesOrderWriterCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ClassName: clsSalesOrderWriterCSV
' Author: JT
' Created: 2025-11-19
' Description: Output adapter. Writes one CSV row per line item per sales order.
' DependsOn: clsLoggingSystem, clsSalesOrder, clsSalesOrderLineItem
' ChangeLog:
'   - 0.1.0 - 2025-11-07 - Initial prototype
'   - 0.2.0 - 2025-11-10 - Switched from textstream to excel workbook saved as csv
'   - 0.3.0 - 2025-11-18 - Refactor to implement IInvoiceWriter
'   - 0.3.1 - 2025-11-19 - Added error handling
'   - 0.3.2 - 2026-01-28 - Removed invoice email from CSV
'   - 1.0.0 - 2026-02-11 - Initial release
'   - 2.0.0 - 2026-xx-xx - V3 port: renamed classes, removed SSACustomerID/ExternalID, standardised error handling

Private m_Logger As clsLoggingSystem
Private m_WBk As Workbook
Private m_WkS As Worksheet
Private m_CurrentRow As Long

' ==============
' Initialization
' ==============

Public Sub Initialize(Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing
    Set m_Logger = Logger
End Sub

' ====
' APIs
' ====

Public Sub BeginOutput()
    Call AssertInitialized("BeginOutput")
    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    If Not m_WBk Is Nothing Then
        Err.Raise vbObjectError + 523, "clsSalesOrderWriterCSV.BeginOutput", "BeginOutput has already been called."
    End If

    Set m_WBk = Workbooks.Add
    Set m_WkS = m_WBk.Worksheets(1)
    m_CurrentRow = 1
    Call WriteHeaders

    Exit Sub
ErrorHandler:
    Call HandleError("BeginOutput", Err.Description)
End Sub

Public Sub WriteInvoice(SO As clsSalesOrder)
    Call AssertInitialized("WriteInvoice")
    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    If SO Is Nothing Then
        Err.Raise vbObjectError + 513, "clsSalesOrderWriterCSV.WriteInvoice", "SO cannot be Nothing."
    End If

    If m_WkS Is Nothing Then
        Err.Raise vbObjectError + 522, "clsSalesOrderWriterCSV.WriteInvoice", "BeginOutput must be called before WriteInvoice."
    End If

    Call WriteLineItems(SO)

    Exit Sub
ErrorHandler:
    Call HandleError("WriteInvoice", Err.Description)
End Sub

Public Sub SaveInvoice(FilePath As String)
    Call AssertInitialized("SaveInvoice")
    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    If m_WkS Is Nothing Then
        Err.Raise vbObjectError + 522, "clsSalesOrderWriterCSV.SaveInvoice", "BeginOutput must be called before SaveInvoice."
    End If

    m_WBk.SaveAs FileName:=FilePath, FileFormat:=xlCSV

    Exit Sub
ErrorHandler:
    Call HandleError("SaveInvoice", Err.Description)
End Sub

Public Sub CloseInvoice()
    Call AssertInitialized("CloseInvoice")
    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    If Not m_WBk Is Nothing Then
        m_WBk.Close SaveChanges:=False
        Set m_WkS = Nothing
        Set m_WBk = Nothing
    End If

    Exit Sub
ErrorHandler:
    Call HandleError("CloseInvoice", Err.Description)
End Sub

' ===============
' Private Helpers
' ===============

Private Sub WriteHeaders()
    With m_WkS
        .Cells(m_CurrentRow, 1).Value = "Customer"
        .Cells(m_CurrentRow, 2).Value = "Date"
        .Cells(m_CurrentRow, 3).Value = "Item"
        .Cells(m_CurrentRow, 4).Value = "Description"
        .Cells(m_CurrentRow, 5).Value = "Quantity"
        .Cells(m_CurrentRow, 6).Value = "Rate"
        .Cells(m_CurrentRow, 7).Value = "PO Number"
        .Cells(m_CurrentRow, 8).Value = "Name of Requester"
        .Cells(m_CurrentRow, 9).Value = "Email of Requester"
        .Cells(m_CurrentRow, 10).Value = "Service Date"
        .Cells(m_CurrentRow, 11).Value = "Submission Ref"
    End With
    m_CurrentRow = m_CurrentRow + 1
End Sub

Private Sub WriteLineItems(SO As clsSalesOrder)
    Dim LineItem As clsSalesOrderLineItem
    For Each LineItem In SO.LineItems
        Call WriteLineItem(SO, LineItem)
    Next LineItem
End Sub

Private Sub WriteLineItem(SO As clsSalesOrder, LineItem As clsSalesOrderLineItem)
    With m_WkS
        .Cells(m_CurrentRow, 1).Value = SO.CustomerName
        .Cells(m_CurrentRow, 2).Value = Format(SO.DateReceived, "yyyy-mm-dd")
        .Cells(m_CurrentRow, 3).Value = LineItem.Item
        .Cells(m_CurrentRow, 4).Value = LineItem.Description
        .Cells(m_CurrentRow, 5).Value = LineItem.Quantity
        .Cells(m_CurrentRow, 6).Value = LineItem.Rate
        .Cells(m_CurrentRow, 7).Value = SO.PONumber
        .Cells(m_CurrentRow, 8).Value = SO.RequesterName
        .Cells(m_CurrentRow, 9).Value = SO.RequesterEmail
        .Cells(m_CurrentRow, 10).Value = SO.ServiceDate
        .Cells(m_CurrentRow, 11).Value = SO.SubmissionRef
    End With
    m_CurrentRow = m_CurrentRow + 1
End Sub

Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsSalesOrderWriterCSV." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsSalesOrderWriterCSV." & MethodName, "The writer must be initialized before use."
    End If
End Sub
