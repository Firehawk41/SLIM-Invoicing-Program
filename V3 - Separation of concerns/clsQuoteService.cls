VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQuoteService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsQuoteService
' Author: JT
' Created: 2026-05-19
' Description: Sole public API for the quote domain. Owns and
'              coordinates clsQuoteRepository and clsQuoteCache.
'              All reads are served from cache. All writes hit the
'              repo and trigger a full cache rebuild — the graph
'              is too complex to update incrementally.
' DependsOn: clsLoggingSystem, clsQuoteRepository, clsQuoteCache,
'            clsQuote
' ChangeLog:
'   - 1.0.0 - 2026-05-19 - Initial release
'
' ================================================================
' LAYER: Service — sole public API for the quote domain.
'        Owns repo and cache internally. Callers never touch
'        repo or cache directly.
'
' LIFECYCLE:
'   1. Initialize(Repo, Cache, Logger)
'      — repo and cache constructed externally and injected
'   2. Read methods served from cache (no DB round trips)
'   3. Write methods hit repo, then trigger full cache rebuild
'
' CONSUMERS:
'   clsSalesOrderPricingEngine — AllQuotes, GetByCustomerID,
'                                AllDefaultQuotes at pipeline
'                                startup
'   clsQuoteFormController     — CreateQuote, SaveQuote,
'                                DeleteQuote (future — quote
'                                editor UI)
'
' NOTE: No LoadQuote(ID) on the public interface — all reads go
'       through cache collections. Add if the quote editor
'       requires single-record lookup.
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_Repo   As clsQuoteRepository
Private m_Cache  As clsQuoteCache

' ================================================================
' Initialization
' ================================================================

Public Sub Initialize(Repo As clsQuoteRepository, _
                      Cache As clsQuoteCache, _
                      Logger As clsLoggingSystem)
    Debug.Assert Not Repo Is Nothing
    Debug.Assert Not Cache Is Nothing
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing

    Set m_Logger = Logger
    Set m_Repo = Repo
    Set m_Cache = Cache
End Sub

' ================================================================
' Reads — all served from cache
' ================================================================

Public Function AllQuotes() As Collection
    AssertInitialized "AllQuotes"
    Set AllQuotes = m_Cache.AllCustomerQuotes()
End Function

Public Function AllDefaultQuotes() As Collection
    AssertInitialized "AllDefaultQuotes"
    Set AllDefaultQuotes = m_Cache.GetDefaultQuotes()
End Function

Public Function GetByCustomerID(CustomerID As Long) As Collection
    AssertInitialized "GetByCustomerID"
    Debug.Assert CustomerID > 0
    Set GetByCustomerID = m_Cache.GetByCustomerID(CustomerID)
End Function

Public Function ExistsByCustomerID(CustomerID As Long) As Boolean
    AssertInitialized "ExistsByCustomerID"
    ExistsByCustomerID = m_Cache.ExistsByCustomerID(CustomerID)
End Function

' ================================================================
' Writes
' ================================================================

Public Function CreateQuote(Q As clsQuote) As Long
    AssertInitialized "CreateQuote"
    Debug.Assert Not Q Is Nothing
    Debug.Assert Q.ID = 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    If Not Q.Validate Then
        Err.Raise vbObjectError + 1, "clsQuoteService.CreateQuote", _
            "Quote failed validation before persist."
    End If

    Dim NewID As Long
    NewID = m_Repo.InsertQuote(Q)
    m_Cache.Build

    CreateQuote = NewID
    Exit Function
EH:
    HandleError "CreateQuote"
End Function

Public Function CreateDefaultQuote(Q As clsQuote) As Long
    AssertInitialized "CreateDefaultQuote"
    Debug.Assert Not Q Is Nothing
    Debug.Assert Q.ID = 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    If Not Q.Validate Then
        Err.Raise vbObjectError + 1, "clsQuoteService.CreateDefaultQuote", _
            "Default quote failed validation before persist."
    End If

    Dim NewID As Long
    NewID = m_Repo.InsertDefaultQuote(Q)
    m_Cache.Build

    CreateDefaultQuote = NewID
    Exit Function
EH:
    HandleError "CreateDefaultQuote"
End Function

Public Sub SaveQuote(Q As clsQuote)
    AssertInitialized "SaveQuote"
    Debug.Assert Not Q Is Nothing
    Debug.Assert Q.ID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    If Not Q.Validate Then
        Err.Raise vbObjectError + 1, "clsQuoteService.SaveQuote", _
            "Quote failed validation before save."
    End If

    m_Repo.UpdateQuote Q
    m_Cache.Build

    Exit Sub
EH:
    HandleError "SaveQuote", "Failed on ID: " & Q.ID
End Sub

Public Sub DeleteQuote(QuoteID As Long)
    AssertInitialized "DeleteQuote"
    Debug.Assert QuoteID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_Repo.DeleteQuote QuoteID
    m_Cache.Build

    Exit Sub
EH:
    HandleError "DeleteQuote", "Failed on ID: " & QuoteID
End Sub

Public Sub DeleteDefaultQuote(QuoteID As Long)
    AssertInitialized "DeleteDefaultQuote"
    Debug.Assert QuoteID > 0
    If Not m_Logger.DebugMode Then On Error GoTo EH

    m_Repo.DeleteDefaultQuote QuoteID
    m_Cache.Build

    Exit Sub
EH:
    HandleError "DeleteDefaultQuote", "Failed on ID: " & QuoteID
End Sub

' ================================================================
' Private Helpers
' ================================================================

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "clsQuoteService." & MethodName, _
            "The service must be initialized before use."
    End If
End Sub

Private Sub HandleError(MethodName As String, Optional Context As String = "")
    m_Logger.LogError "clsQuoteService." & MethodName, Err.Number, _
        Err.Description & IIf(Context <> "", " | " & Context, "")
End Sub

