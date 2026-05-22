VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceWriterCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' MacroName: clsInvoiceWriterCSV
' Author: JT
' Created: 2025-11-19
' Description: Object for writing invoices to CSV
' DependsOn: clsLoggingSystem, clsInvoiceSalesOrder, clsInvoiceLineItem
' ChangeLog:
'   - 0.1.0 - 2025-11-07 - Initial prototype
'   - 0.2.0 - 2025-11-10 - Switched from textstream to excel workbook that is saved as csv
'   - 0.3.0 - 2025-11-18 - Refactor to implement IInvoiceWriter
'   - 0.3.1 - 2025-11-19 - Added error handling
'   - 0.3.2 - 2026-01-28 - Removed 'invoice email' from the CSV
'   - 1.0.0 - 2026-02-11 - Initial release
Private m_Logger As clsLoggingSystem
Private m_WBk As Workbook
Private m_WkS As Worksheet
Private m_CurrentRow As Long
Public Sub Initialize(Logger As clsLoggingSystem)
    Set m_Logger = Logger
End Sub
Public Sub BeginOutput()
    On Error GoTo ErrorHandler

    If Not m_WBk Is Nothing Then
        Err.Raise vbObjectError + 523, "BeginOutput", "BeginOutput has already been called."
    End If

    ' Create new workbook to be saved as csv
    Set m_WBk = Workbooks.Add
    Set m_WkS = m_WBk.Worksheets(1)
    m_CurrentRow = 1

    ' Write headers
    Call WriteHeaders
    Exit Sub
ErrorHandler:
    Call m_Logger.LogMessage("clsInvoiceWriterCSV.BeginOutput", LogLevelEnum.LogError, Err.Number & ": " & Err.Description)
    If m_Logger.DebugMode Then Err.Raise Err.Number, "clsInvoiceWriterCSV.BeginOutput", Err.Description
End Sub
Public Sub WriteInvoice(Invoice As clsInvoiceSalesOrder)
    On Error GoTo ErrorHandler

    If Invoice Is Nothing Or Invoice.LineItems.Count = 0 Then
        Err.Raise vbObjectError + 513, "IInvoiceWriter_WriteInvoice", "Invoice cannot be Nothing"
    End If

    If m_WkS Is Nothing Then
        Err.Raise vbObjectError + 522, "WriteInvoice", "BeginOutput must be called before WriteInvoice."
    End If

    ' Write all line items
    Call WriteLineItems(Invoice, Invoice.LineItems)

    Exit Sub
ErrorHandler:
    Call m_Logger.LogMessage("clsInvoiceWriterCSV.WriteInvoice", LogLevelEnum.LogError, Err.Number & ": " & Err.Description)
    If m_Logger.DebugMode Then Err.Raise Err.Number, "clsInvoiceWriterCSV.WriteInvoice", Err.Description
End Sub
Private Sub WriteHeaders()
    With m_WkS
        .Cells(m_CurrentRow, 1).Value = "Customer"
        .Cells(m_CurrentRow, 2).Value = "Internal ID"
        .Cells(m_CurrentRow, 3).Value = "Date"
        .Cells(m_CurrentRow, 4).Value = "Item"
        .Cells(m_CurrentRow, 5).Value = "Description"
        .Cells(m_CurrentRow, 6).Value = "Quantity"
        .Cells(m_CurrentRow, 7).Value = "Rate"
        .Cells(m_CurrentRow, 8).Value = "PO Number"
        .Cells(m_CurrentRow, 9).Value = "Name of Requester"
        .Cells(m_CurrentRow, 10).Value = "Email of Requester"
        .Cells(m_CurrentRow, 11).Value = "Service Date"
        .Cells(m_CurrentRow, 12).Value = "External ID"
    End With
    m_CurrentRow = m_CurrentRow + 1
End Sub
Public Sub SaveInvoice(FilePath As String)
    On Error GoTo ErrorHandler

    If Not m_WkS Is Nothing Then
        m_WBk.SaveAs FileName:=FilePath, FileFormat:=xlCSV
    Else
        Call m_Logger.LogMessage("clsInvoiceCSVExporter.SaveInvoice", LogLevelEnum.LogWarning, "No workbook to save. CreateCSV was not called.")
    End If
    Exit Sub
ErrorHandler:
    Call m_Logger.LogMessage("clsInvoiceWriterCSV.SaveInvoice", LogLevelEnum.LogError, Err.Number & ": " & Err.Description)
    If m_Logger.DebugMode Then Err.Raise Err.Number, "clsInvoiceWriterCSV.SaveInvoice", Err.Description
End Sub
Public Sub CloseInvoice()
    On Error GoTo ErrorHandler

    If Not m_WBk Is Nothing Then
        m_WBk.Close SaveChanges:=False
        Set m_WkS = Nothing
        Set m_WBk = Nothing
    End If

    Exit Sub
ErrorHandler:
    Call m_Logger.LogMessage("clsInvoiceWriterCSV.CloseInvoice", LogLevelEnum.LogError, Err.Number & ": " & Err.Description)
    If m_Logger.DebugMode Then Err.Raise Err.Number, "clsInvoiceWriterCSV.CloseInvoice", Err.Description
End Sub

Private Sub WriteLineItems(Invoice As clsInvoiceSalesOrder, LineItems As Collection)
    Dim LineItem As IInvoiceLineItem
    For Each LineItem In LineItems
        If LineItem Is Nothing Then
            Err.Raise vbObjectError + 513, "IInvoiceWriter_WriteLineItems", "LineItem cannot be Nothing"
        End If
        Call WriteLineItem(Invoice, LineItem)
    Next LineItem
End Sub
Private Sub WriteLineItem(Invoice As clsInvoiceSalesOrder, LineItem As clsInvoiceLineItem)

    With m_WkS
        .Cells(m_CurrentRow, 1).Value = Invoice.CustomerName
        .Cells(m_CurrentRow, 2).Value = Invoice.SSACustomerID
        .Cells(m_CurrentRow, 3).Value = Format(Invoice.DateReceived, "yyyy-mm-dd")
        .Cells(m_CurrentRow, 4).Value = LineItem.Item
        .Cells(m_CurrentRow, 5).Value = LineItem.Description
        .Cells(m_CurrentRow, 6).Value = LineItem.Quantity
        .Cells(m_CurrentRow, 7).Value = LineItem.Rate
        .Cells(m_CurrentRow, 8).Value = Invoice.PONumber
        .Cells(m_CurrentRow, 9).Value = Invoice.RequesterName
        .Cells(m_CurrentRow, 10).Value = Invoice.RequesterEmail
        .Cells(m_CurrentRow, 11).Value = Invoice.ServiceDate
        .Cells(m_CurrentRow, 12).Value = Invoice.ExternalID
    End With

    m_CurrentRow = m_CurrentRow + 1

End Sub
