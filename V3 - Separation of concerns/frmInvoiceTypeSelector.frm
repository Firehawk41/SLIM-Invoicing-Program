VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInvoiceTypeSelector 
   Caption         =   "Invoicing Macro"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5430
   OleObjectBlob   =   "frmInvoiceTypeSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInvoiceTypeSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: frmInvoiceTypeSelector
' Author: JT
' Created: 2025-06-05
' Description: Entry point UI for the invoice generation pipeline.
'              Collects invoice selection mode (Individual or
'              Batch) and date range from the user before handing
'              off to modInvoiceSystem.
' DependsOn: clsLoggingSystem, modEnums
' ChangeLog:
'   - 2.0.1 - 2025-06-05 - Added metadata
'   - 2.1.0 - 2025-11-18 - Added output type selection
'   - 3.0.0 - 2026-01-28 - Refactored for CSV output
'   - 4.0.0 - 2026-05-08 - V3 port: removed Customer A/Customer B invoice
'                           types, Word output option, and
'                           week/month combo logic; replaced with
'                           date range combo boxes
'   - 4.0.1 - 2026-05-19 - Fixed missing folder picker in batch
'                           mode. Replaced dtpStartDate/dtpEndDate
'                           references with cboStartDate/cboEndDate.
'
' ================================================================
' LAYER: UI — UserForm. Collects user input only. No business
'        logic. Hands off via Property Gets to modInvoiceSystem.
'
' CONTROLS:
'   btnIndividualMode  — Radio. Disables date range combos.
'   btnBatchMode       — Radio. Enables date range combos.
'                        Selected by default.
'   cboStartDate       — Combo. Populated with last 31 days.
'                        User may type a custom value.
'   cboEndDate         — Combo. Populated with last 31 days.
'                        User may type a custom value.
'   OKButton           — Triggers file/folder picker and
'                        validates date range before hiding.
'   CancelButton       — Sets Cancelled = True and hides.
'
' PUBLIC INTERFACE:
'   Initialize(Logger) — Must be called before Show.
'   InvoiceSelection   — Returns InvoiceSelectionEnum.
'   StartDate          — Returns Date. Batch mode only.
'   EndDate            — Returns Date. Batch mode only.
'   SelectedFilePath   — Returns String. Set during OKButton.
'   Cancelled          — Returns Boolean.
' ================================================================

Private m_Logger As clsLoggingSystem
Private m_Cancelled As Boolean
Private m_SelectedFilePath As String

' =================
' Public Properties
' =================

Public Property Get InvoiceSelection() As InvoiceSelectionEnum
    If btnIndividualMode.Value = True Then
        InvoiceSelection = InvoiceSelectionEnum.Individual
    ElseIf btnBatchMode.Value = True Then
        InvoiceSelection = InvoiceSelectionEnum.Batch
    End If
End Property

Public Property Get StartDate() As Date
    StartDate = CDate(cboStartDate.Value)
End Property

Public Property Get EndDate() As Date
    EndDate = CDate(cboEndDate.Value)
End Property

Public Property Get SelectedFilePath() As String
    SelectedFilePath = m_SelectedFilePath
End Property

Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property

' ==============
' Initialization
' ==============

Public Sub Initialize(Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing
    Set m_Logger = Logger

    btnIndividualMode.Value = True

    Call PopulateDateCombos
End Sub

' ============
' Form Actions
' ============

Private Sub btnIndividualMode_Click()
    Call AssertInitialized("btnIndividualMode_Click")
    cboStartDate.Enabled = False
    cboEndDate.Enabled = False
End Sub

Private Sub btnBatchMode_Click()
    Call AssertInitialized("btnBatchMode_Click")
    cboStartDate.Enabled = True
    cboEndDate.Enabled = True
End Sub


Private Sub OKButton_Click()
    Call AssertInitialized("OKButton_Click")
    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler

    If btnIndividualMode.Value Then
        m_SelectedFilePath = Application.GetOpenFilename( _
            FileFilter:="Excel files (*.xlsx), *.xlsx", _
            MultiSelect:=False)

        If m_SelectedFilePath = "False" Then
            m_Cancelled = True
            Exit Sub
        End If

    ElseIf btnBatchMode.Value Then
        If Not IsDate(cboStartDate.Value) Then
            MsgBox "Start date is not a valid date.", vbExclamation, "Invalid Date"
            cboStartDate.SetFocus
            Exit Sub
        End If

        If Not IsDate(cboEndDate.Value) Then
            MsgBox "End date is not a valid date.", vbExclamation, "Invalid Date"
            cboEndDate.SetFocus
            Exit Sub
        End If

        If CDate(cboStartDate.Value) > CDate(cboEndDate.Value) Then
            MsgBox "Start date must be before end date.", vbExclamation, "Invalid Date Range"
            cboStartDate.SetFocus
            Exit Sub
        End If

        Dim fd As FileDialog
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
        With fd
            .Title = "Select folder containing reports to invoice"
            .AllowMultiSelect = False
            If .Show = False Then
                m_Cancelled = True
                Exit Sub
            End If
            m_SelectedFilePath = .SelectedItems(1)
            If Right(m_SelectedFilePath, 1) <> "\" Then
                m_SelectedFilePath = m_SelectedFilePath & "\"
            End If
        End With
    End If

    Me.Hide
    Exit Sub
ErrorHandler:
    Call HandleError("OKButton_Click", Err.Description)
End Sub

Private Sub CancelButton_Click()
    m_Cancelled = True
    Me.Hide
End Sub

' ===============
' Private Helpers
' ===============
Private Sub PopulateDateCombos()
    Dim i As Long
    Dim D As Date

    cboStartDate.Clear
    cboEndDate.Clear

    For i = 0 To 30
        D = Date - i
        cboStartDate.AddItem Format(D, "yyyy-mm-dd")
        cboEndDate.AddItem Format(D, "yyyy-mm-dd")
    Next i

    cboStartDate.Value = Format(Date - 30, "yyyy-mm-dd")
    cboEndDate.Value = Format(Date, "yyyy-mm-dd")
End Sub

Private Sub HandleError(MethodName As String, Optional Context As String = "")
    Call m_Logger.LogError("frmInvoiceTypeSelector." & MethodName, Err.Number, Err.Description & IIf(Context <> "", " | " & Context, ""), True)
End Sub

Private Sub AssertInitialized(MethodName As String)
    If m_Logger Is Nothing Then
        Err.Raise 911, "frmInvoiceTypeSelector." & MethodName, "The form must be initialized before use."
    End If
End Sub

