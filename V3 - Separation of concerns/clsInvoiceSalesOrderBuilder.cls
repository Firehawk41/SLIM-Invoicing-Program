VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceSalesOrderBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceSalesOrderBuilder
' Author: JT
' Created: 2026-03-13
' Description: Object in charge of building sales order objects
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsInvoiceSalesOrder, clsTRSubmission
' ChangeLog:
'   - 0.1.0 - 2026-03-12 - Initial prototype

Private m_Logger As clsLoggingSystem
Private m_LineItemBuilder As clsInvoiceLineItemBuilder
' ==============
' Initialization
' ==============
Public Sub Initialize(LineItemBuilder As clsInvoiceLineItemBuilder, Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing
    Debug.Assert Not LineItemBuilder Is Nothing
    Debug.Assert m_LineItemBuilder Is Nothing

    If Not Logger.DebugMode Then On Error GoTo ErrorHandler

    If Not m_Logger Is Nothing Then Err.Raise 912, , "The builder is already initialized."

    Set m_Logger = Logger
    Set m_LineItemBuilder = LineItemBuilder

    Exit Sub
ErrorHandler:
    Call HandleError("Initialize", Err.Description)

End Sub

' ====
' APIs
' ====
Public Function BuildFromSubmission(Submission As clsTRSubmission) As clsInvoiceSalesOrder
    Dim SalesOrder As clsInvoiceSalesOrder

    Debug.Assert Not Submission Is Nothing

    Call AssertInitialized("BuildFromSubmission")

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    Set SalesOrder = New clsInvoiceSalesOrder
    Call SalesOrder.Initialize(Logger)

    ' --- Build the dictionary for the SalesOrder ---
    Set SalesOrderDict = BuildSalesOrderDict(Submission)

    ' --- Populate the SalesOrder object from the dict ---
    Call SalesOrder.FromDict(SalesOrderDict)

    ' --- Placeholder: Build line items and add them ---
    ' Call BuildLineItems(Submission, SalesOrder)

    Set BuildFromSubmission = SalesOrder

    Exit Function
ErrorHandler:
    Call HandleError("BuildFromSubmission", Err.Description)
End Function

' ===============
' Private Helpers
' ===============
Private Function BuildSalesOrderDict(Submission As clsTRSubmission) As Scripting.Dictionary
    Dim Dict As Scripting.Dictionary
    Set Dict = New Scripting.Dictionary

    ' --- Map properties from submission to dict keys ---
    Dict("ID") = Submission.ID
    Dict("Customer") = Submission.Customer.CustomerName
    Dict("SubmissionDate") = Submission.DateReceived
    Dict("PONumber") = Submission.POInformation
    Dict("PaymentType") = Submission
    Dict("PaymentTerms") = Submission
    Dict("InvoiceEmail") = Submission.InvoiceEmailMainFirst
    Dict("NameOfRequester") = Submission.CustomerContact
    Dict("EmailOfRequester") = Submission.ResultsEmailMainFirst
    Dict("ServiceDate") = Submission.EffectiveServiceDate
    Dict("CustomerID") = Submission.Customer.CustomerID
    Dict("SubmissionID") = Submission.SubmissionRef

    ' --- Placeholder for line items ---
    ' Dict("LineItems") = New Collection

    Set BuildSalesOrderDict = Dict
End Function
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsInvoiceSalesOrderBuilder." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsInvoiceSalesOrderBuilder." & MethodName, "The builder must be initialized before use."
    End If
End Sub

