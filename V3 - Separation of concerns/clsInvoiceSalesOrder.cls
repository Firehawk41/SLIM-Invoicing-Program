VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceSalesOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceSalesOrder
' Author: JT
' Created: 2026-03-13
' Description: Object in charge of storing sales order properties
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsInvoiceLineItem, modEnums
' ChangeLog:
'   - 0.1.0 - 2026-03-12 - Initial prototype

Private m_Logger As clsLoggingSystem
Private m_LineItems As Collection
Private m_ID As Long
Private m_Customer As String
Private m_SubmissionDate As Date
Private m_PONumber As String
Private m_PaymentType As PaymentTypeEnum
Private m_PaymentTerms As String
Private m_InvoiceEmail As String
Private m_NameOfRequester As String
Private m_EmailOfRequester As String
Private m_ServiceDate As Date
Private m_CustomerID As Long
Private m_SubmissionID As String
' ===============
' Get Properties
' ==============
Public Property Get CustomerID() As Long
    CustomerID = m_CustomerID
End Property
Public Property Get SubmissionID() As String
    SubmissionID = m_SubmissionID
End Property
Public Property Get SubmissionDate() As Date
    SubmissionDate = m_SubmissionDate
End Property
' ==============
' Initialization
' ==============
Public Sub Initialize(Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    If Not Logger.DebugMode Then On Error GoTo ErrorHandler

    If Not m_Logger Is Nothing Then Err.Raise 912, , "The sales order is already initialized."

    Set m_Logger = Logger
    Set m_LineItems = New Collection

    Exit Sub
ErrorHandler:
    Call HandleError("Initialize", Err.Description)

End Sub

' ====
' APIs
' ====
Public Sub FromDict(Dict As Scripting.Dictionary)

    Debug.Assert Not Dict Is Nothing

    Call AssertInitialized("FromDict")

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    If Dict.Exists("ID") Then m_ID = Dict("ID")
    If Dict.Exists("Customer") Then m_Customer = Dict("Customer")
    If Dict.Exists("SubmissionDate") Then m_SubmissionDate = Dict("SubmissionDate")
    If Dict.Exists("PONumber") Then m_PONumber = Dict("PONumber")
    If Dict.Exists("PaymentType") Then m_PaymentType = Dict("PaymentType")
    If Dict.Exists("PaymentTerms") Then m_PaymentTerms = Dict("PaymentTerms")
    If Dict.Exists("InvoiceEmail") Then m_InvoiceEmail = Dict("InvoiceEmail")
    If Dict.Exists("NameOfRequester") Then m_NameOfRequester = Dict("NameOfRequester")
    If Dict.Exists("EmailOfRequester") Then m_EmailOfRequester = Dict("EmailOfRequester")
    If Dict.Exists("ServiceDate") Then m_ServiceDate = Dict("ServiceDate")
    If Dict.Exists("CustomerID") Then m_CustomerID = Dict("CustomerID")




    Exit Sub
ErrorHandler:
    Call HandleError("FromDict", Err.Description)
End Sub
Public Function ToDict() As Scripting.Dictionary
    Dim Dict As Scripting.Dictionary
    Set Dict = New Scripting.Dictionary

    Call AssertInitialized("ToDict")

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    ' Map each property to a dictionary key
    Dict("ID") = m_ID
    Dict("Customer") = m_Customer
    Dict("SubmissionDate") = m_SubmissionDate
    Dict("PONumber") = m_PONumber
    Dict("PaymentType") = m_PaymentType
    Dict("PaymentTerms") = m_PaymentTerms
    Dict("InvoiceEmail") = m_InvoiceEmail
    Dict("NameOfRequester") = m_NameOfRequester
    Dict("EmailOfRequester") = m_EmailOfRequester
    Dict("ServiceDate") = m_ServiceDate
    Dict("CustomerID") = m_CustomerID

    ' Placeholder for line items (add later)
    ' Dict("LineItems") = m_LineItemsCollection.ToDictCollection

    Set ToDict = Dict
    Exit Function
ErrorHandler:
    Call HandleError("ToDict", Err.Description)
End Function
Public Sub AddLineItem(LineItem As clsInvoiceLineItem)
    Debug.Assert Not LineItem Is Nothing

    Call AssertInitialized("AddLineItem")

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    If Not LineItem Is Nothing Then
        m_LineItems.Add LineItem
    End If

    Exit Sub
ErrorHandler:
    Call HandleError("LineItem", Err.Description)
End Sub
Public Function Validate() As Boolean
    Call AssertInitialized("FromDict")

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler







    Exit Function
ErrorHandler:
    Call HandleError("Validate", Err.Description)
End Function

' ===============
' Private Helpers
' ===============

Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsInvoiceSalesOrder." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsInvoiceSalesOrder." & MethodName, "The builder must be initialized before use."
    End If
End Sub


