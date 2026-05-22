VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceLineItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceLineItem
' Author: JT
' Created: 2025-11-10
' Description: Object to include all line item properties to nest inside Sales Orders
' DependsOn: clsLoggingSystem, modEnum
' ChangeLog:
'   - 0.1.0 - 2026-03-25 - Initial prototype

Private m_Logger As clsLoggingSystem
Private m_ID As Long
Private m_Item As String
Private m_Description As String
Private m_Quantity As Long
Private m_Rate As Currency
Private m_AnalysisID As Long
Private m_ProcessingTime As ProcessingTimeEnum
Private m_SourceQuoteID As Long

' ==============
' Get Properties
' ==============
Public Property Get Total() As Currency
    Total = m_Quantity * m_Rate
End Property
' ==============
' Initialization
' ==============
Public Sub Initialize(Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    If Not Logger.DebugMode Then On Error GoTo ErrorHandler

    If Not m_Logger Is Nothing Then Err.Raise 912, , "The line item is already initialized."

    Set m_Logger = Logger

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
    If Dict.Exists("item") Then m_Item = Dict("item")
    If Dict.Exists("description") Then m_Description = Dict("description")
    If Dict.Exists("quantity") Then m_Quantity = Dict("quantity")
    If Dict.Exists("rate") Then m_Rate = Dict("rate")
    If Dict.Exists("analysis_id") Then m_AnalysisID = Dict("analysis_id")
    If Dict.Exists("processing_time") Then m_ProcessingTime = Dict("processing_time")
    If Dict.Exists("source_quote_id") Then m_SourceQuoteID = Dict("source_quote_id")

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
    Dict("item") = m_Item
    Dict("description") = m_Description
    Dict("quantity") = m_Quantity
    Dict("rate") = m_Rate
    Dict("analysis_id") = m_AnalysisID
    Dict("processing_time") = m_ProcessingTime
    Dict("source_quote_id") = m_SourceQuoteID

    Set ToDict = Dict
    Exit Function
ErrorHandler:
    Call HandleError("ToDict", Err.Description)
End Function
Public Function Validate() As Boolean
    Call AssertInitialized("FromDict")

    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    If m_Quantity <= 0 Then Validate = False
    If m_Rate < 0 Then Validate = False
    If m_Item = "" Then Validate = False

    Exit Function
ErrorHandler:
    Call HandleError("Validate", Err.Description)
End Function

' ===============
' Private Helpers
' ===============
Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("clsInvoiceLineItem." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsInvoiceLineItem." & MethodName, "The line item must be initialized before use."
    End If
End Sub
