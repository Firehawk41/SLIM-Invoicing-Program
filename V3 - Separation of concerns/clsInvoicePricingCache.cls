VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoicePricingCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoicePricingCache
' Author: JT
' Created: 2025-06-11
' Description: Caches invoice pricing details for lazy lookup
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsQuoteLoader, modEnums, clsQuote
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
'   - 1.0.2 - 2025-12-31 - Changed 'Now' to 'Date' in GetDefaultPriceArray to fix bug that took time into account when checking if the default price was valid
'   - 2.0.0 - 2026-04-02 - Refactored to use quote loader and quote objects. Instead of building the quotes from the DB, the cache now sorts and caches quotes from the loader
Private m_Logger As clsLoggingSystem
Private m_QuoteLoader As clsQuoteLoader
Private m_CustomerKeyedQuoteDict As Scripting.Dictionary
Private m_DefaultQuoteColl As Collection
' ==============
' Get Properties
' ==============
Public Property Get CustomerKeyedQuoteDict() As Scripting.Dictionary
    If m_CustomerKeyedQuoteDict Is Nothing Then
        Set m_CustomerKeyedQuoteDict = CompileCustomerKeyedQuoteDict
    End If
    Set CustomerKeyedQuoteDict = m_CustomerKeyedQuoteDict
End Property
Public Property Get DefaultQuoteColl() As Collection
    If m_DefaultQuoteColl Is Nothing Then
        Set m_DefaultQuoteColl = CompileDefaultQuoteColl
    End If
    Set DefaultQuoteColl = m_DefaultQuoteColl
End Property
' ==============
' Initialization
' ==============
Public Sub Initialize(QuoteLoader As clsQuoteLoader, Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing
    Debug.Assert Not QuoteLoader Is Nothing
    Debug.Assert m_QuoteLoader Is Nothing

    If Not Logger.DebugMode Then On Error GoTo ErrorHandler

    Set m_Logger = Logger
    Set m_QuoteLoader = QuoteLoader


    Exit Sub
ErrorHandler:
    Call HandleError("Initialize", Err.Description)
End Sub
' =====================
' Core API
' =====================
Public Function RetrieveQuotesByCustomer(CustomerID As Long) As Collection
    Debug.Assert CustomerID > 0

    AssertInitialized "RetrieveQuotesByCustomer"

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    ' Early exit if there are no quotes for this customer
    If Not Me.CustomerKeyedQuoteDict.Exists(CustomerID) Then
        Set RetrieveQuotesByCustomer = New Collection
        Exit Function
    End If

    Set RetrieveQuotesByCustomer = CloneQuoteCollection(Me.CustomerKeyedQuoteDict(CustomerID))

    Exit Function
ErrorHandler:
    HandleError "RetrieveQuotesByCustomer", "CustomerID=" & CustomerID
End Function
Public Function RetrieveDefaultQuotes() As Collection

    AssertInitialized "RetrieveDefaultQuotes"

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    Set RetrieveDefaultQuotes = CloneQuoteCollection(Me.DefaultQuoteColl)

    Exit Function
ErrorHandler:
    HandleError "RetrieveDefaultQuotes"
End Function

' ===============
' Private Helpers
' ===============
Private Function CloneQuoteCollection(Source As Collection) As Collection
    Dim NewColl As New Collection
    Dim Q As clsQuote

    For Each Q In Source
        NewColl.Add Q
    Next

    Set CloneQuoteCollection = NewColl
End Function
Private Function CompileCustomerKeyedQuoteDict() As Scripting.Dictionary
    Dim QuoteColl As Collection
    Dim CustomerColl As Collection
    Dim Q As clsQuote
    Dim CustomerID As Variant
    Dim Dict As Scripting.Dictionary

    Set Dict = New Scripting.Dictionary


    For Each Q In m_QuoteLoader.Quotes

        Set CustomerColl = Q.Customers

        For Each CustomerID In CustomerColl
            If Dict.Exists(CustomerID) Then
                Dict(CustomerID).Add Q
            Else

                Set QuoteColl = New Collection
                QuoteColl.Add Q

                Set Dict(CustomerID) = QuoteColl
            End If
        Next CustomerID
    Next Q

    Set CompileCustomerKeyedQuoteDict = Dict
End Function
Private Function CompileDefaultQuoteColl() As Collection
    Dim QuoteColl As Collection
    Dim Q As clsQuote

    Set QuoteColl = New Collection

    For Each Q In m_QuoteLoader.DefaultQuotes
        QuoteColl.Add Q
    Next Q

    Set CompileDefaultQuoteColl = QuoteColl

End Function
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsInvoicePricingCache." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsInvoicePricingCache." & MethodName, "The builder must be initialized before use."
    End If
End Sub
