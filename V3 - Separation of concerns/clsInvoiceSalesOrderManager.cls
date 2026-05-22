VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceSalesOrderManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceSalesOrderManager
' Author: JT
' Created: 2026-03-12
' Description: Object in charge of managing sales order objects
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsSLIMDictionaryManager, clsInvoiceSalesOrder, clsInvoiceSalesOrderBuilder
' ChangeLog:
'   - 0.1.0 - 2026-03-12 - Initial prototype

Private m_Logger As clsLoggingSystem
Private m_Builder As clsInvoiceSalesOrderBuilder
Private m_SalesOrders As Collection

' ==============
' Get Properties
' ==============
Public Property Get SalesOrders() As Collection
    If m_SalesOrders Is Nothing Then
        Set m_SalesOrders = New Collection
    End If
    Set SalesOrders = m_SalesOrders
End Property

' ==============
' Initialization
' ==============
Public Sub Initialize(Builder As clsInvoiceSalesOrderBuilder, Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    If Not Logger.DebugMode Then On Error GoTo ErrorHandler

    If Not m_Logger Is Nothing Then Err.Raise 912, , "The manager is already initialized."

    Set m_Logger = Logger
    Set m_Builder = Builder

    Exit Sub
ErrorHandler:
    Call HandleError("Initialize", Err.Description)

End Sub

' ====
' APIs
' ====
Public Sub BuildFromSubmissions(Submissions As Collection)
    Dim Submission As clsTRSubmission
    Dim SalesOrder As clsInvoiceSalesOrder

    Debug.Assert Not Submissions Is Nothing
    Debug.Assert Submissions.Count > 0

    Call AssertInitialized("BuildFromSubmissions")

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    Set m_SalesOrders = New Collection

    For Each Submission In Submissions

        Set SalesOrder = m_Builder.BuildFromSubmission(Submission)

        If Not SalesOrder Is Nothing Then
            m_SalesOrders.Add SalesOrder
        End If

    Next Submission

    Exit Sub
ErrorHandler:
    Call HandleError("BuildFromSubmission", Err.Description)
End Sub

' ===============
' Private Helpers
' ===============
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsInvoiceSalesOrderManager." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsInvoiceSalesOrderManager." & MethodName, "The manager must be initialized before use."
    End If
End Sub
