VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: IInvoice
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-11
' Description: Interface for Invoices
' DependsOn: clsTestingRequestSubmission, clsLoggingSystem, clsAccessDatabase, clsInvoiceFactory, clsInvoicePriceRetriever, clsSLIMDictionaryManager, clsInvoiceCustomerBSummary, clsInvoiceCustomerASummary, clsInvoiceDefaultIndividual, clsInvoiceWriter
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Public Property Get CustomerName() As String

End Property
Public Property Get WordInvoice() As Word.Document

End Property
Public Property Get FileName() As String

End Property
Public Sub Initialize(Submissions As Collection, SequenceNumber As Long, Logger As clsLoggingSystem, AccessDB As clsAccessDatabase, Factory As clsInvoiceFactory, PricingRetriever As clsInvoicePriceRetriever, m_SLIMDict As clsSLIMDictionaryManager)

End Sub
Public Sub CreateWordInvoice(WordApp As Word.Application, InvoiceWriter As clsInvoiceWriter)

End Sub
Public Sub AddInvoiceToDatabase()

End Sub
Public Sub CreateExcelInvoice(InvoiceWriter As clsInvoiceWriter)

End Sub
