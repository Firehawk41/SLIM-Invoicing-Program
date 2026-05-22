Attribute VB_Name = "modInvoiceSystem"
Option Explicit
Option Compare Text
' MacroName: modInvoiceSystem
' Executable: True
' Version: 5.0.1
' Author: JT
' Created: 2022-09-08
' LastModified: 2025-06-05
' Description: Automatically generates an invoice
' DependsOn: clsLoggingSystem, clsAccessDatabase, frmInvoiceTypeSelector, modEnums, clsStalenessChecker, clsInvoiceFactory, IInvoiceSubmissionManager, clsInvoiceManager
' ChangeLog:
'   - 5.0.1 - 2025-06-05 - Added metadata

Const INVOICE_DB_PATH As String = "\\PRECILAB-SERVER\LabPlusServer\Documents_In_Works\Thomson\Sample Login.accdb"
Public Sub CreateInvoice()
    Call DisplayInvoiceUF(True)
End Sub
Private Sub DisplayInvoiceUF(DebugMode As Boolean)

    Dim Logger As clsLoggingSystem
    Dim AccessDB As clsAccessDatabase
    Dim InvoiceUserForm As frmInvoiceTypeSelector
    Dim StartDate As Date
    Dim SelectedFilePath As String
    Dim ErrorMessage As String
    Dim InvoiceType As InvoiceTypeEnum
    Dim Checker As New clsStalenessChecker
    
    ' Initiate the log file
    Set Logger = New clsLoggingSystem
    Call Logger.Initialize("CreateInvoice", DebugMode)
    
    ' Authenticate the workbook
    Call Checker.Initialize(ThisWorkbook.Name, ThisWorkbook.FullName, FileDateTime(ThisWorkbook.FullName), Logger)
    If Not Checker.IsCurrent Then Err.Raise 1984, "DisplayInvoiceUF", Checker.IsObsoleteMessage
    
    ' Set variables depending
    If Not Logger.DebugMode Then
        On Error GoTo ErrorHandler
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    End If
    
    ' Initiate the access database object
    Set AccessDB = New clsAccessDatabase
    Call AccessDB.Initialize(INVOICE_DB_PATH, Logger)
    
    ' Set and display the userform
    Set InvoiceUserForm = New frmInvoiceTypeSelector
    Call InvoiceUserForm.Initialize(Logger, AccessDB)
    InvoiceUserForm.Show
    
    ' Check whether the user cancelled the userform
    If InvoiceUserForm.Cancelled Then
        GoTo CleanUp
    End If
    
    ' Extract information from the userform
    StartDate = InvoiceUserForm.StartDate
    SelectedFilePath = InvoiceUserForm.SelectedFilePath
    InvoiceType = InvoiceUserForm.InvoiceType
    
    ' Close the userform
    Set InvoiceUserForm = Nothing
    
    Call GenerateInvoiceFromFile(InvoiceType, StartDate, SelectedFilePath, AccessDB, Logger)
    
CleanUp:
    
    Call Logger.LogMessage("DisplayInvoiceUF", LogLevelEnum.LogInfo, "Completed sub and entering clean up.")

    ' Close the userform
    If Not InvoiceUserForm Is Nothing Then
        Set InvoiceUserForm = Nothing
    End If
    
    ' Close the log file
    If Not Logger Is Nothing Then
        Logger.CloseLogFile
        Set Logger = Nothing
    End If
    
    ' Close the AccessDB
    If Not AccessDB Is Nothing Then
        Set AccessDB = Nothing
    End If
    
    ' Reset settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    Exit Sub
ErrorHandler:
    Call Logger.LogError("Create Invoice", Err.Number, Err.Description)
    GoTo CleanUp
    
End Sub
Public Function GenerateInvoiceFromFile(InvoiceType As InvoiceTypeEnum, StartDate As Date, SelectedFilePath As String, Optional AccessDB As clsAccessDatabase = Nothing, Optional Logger As clsLoggingSystem = Nothing) As Boolean

    Dim Factory As clsInvoiceFactory
    Dim SubmissionManager As IInvoiceSubmissionManager
    Dim InvoiceManager As clsInvoiceManager
    Dim StartTime As Double
    Dim EndTime As Double

    ' Flag the process as unsuccessful by default
    GenerateInvoiceFromFile = True
    
    If Logger Is Nothing Then
        Set Logger = New clsLoggingSystem
        Call Logger.Initialize("GenerateInvoiceFromFile", False)
    End If
    
    If AccessDB Is Nothing Then
        Set AccessDB = New clsAccessDatabase
        Call AccessDB.Initialize(INVOICE_DB_PATH, Logger)
    End If
    
    If Not Logger.DebugMode Then On Error GoTo ErrorHandler

    ' Initiate the invoice factory object
    Set Factory = New clsInvoiceFactory
    Call Factory.Initialize(InvoiceType, Logger)
    
    'Start the timer (used for debugging)
    StartTime = Timer
    
    Set SubmissionManager = Factory.GetInvoiceSubmissionManager
    Call SubmissionManager.Initialize(StartDate, SelectedFilePath, Logger)
    
    If SubmissionManager.Submissions.Count = 0 Or SubmissionManager.Submissions Is Nothing Then
        Err.Raise 1001, "GenerateInvoiceFromFile", "No submissions found"
    ElseIf SubmissionManager.Submissions(1) Is Nothing Then
        Err.Raise 1001, "GenerateInvoiceFromFile", "No submissions found"

    End If
    
    Set InvoiceManager = New clsInvoiceManager
    Call InvoiceManager.Initialize(SubmissionManager, Logger, AccessDB, Factory)
    Call InvoiceManager.GenerateWordInvoiceCollection
    Call InvoiceManager.SaveInvoices(Logger.DebugMode)
    Call InvoiceManager.CloseInvoices(Logger.DebugMode)
    Call InvoiceManager.AddInvoicesToDatabase(Logger.DebugMode)
    
    ' Flag the process as success
    GenerateInvoiceFromFile = True
    
CleanUp:
    ' Stop the timer
    EndTime = Timer
    
    ' Log the timer information
    If Not StartTime = EndTime Then
        Call Logger.LogMessage("CreateInvoice", LogLevelEnum.LogInfo, "The program took " & Round(EndTime - StartTime, 1) & " seconds to run.")
    End If
    
    ' Terminate objects
    Set Factory = Nothing
    Set SubmissionManager = Nothing
    Set InvoiceManager = Nothing
    
    Exit Function
ErrorHandler:
    Call Logger.LogError("GenerateInvoiceFromFile", Err.Number, Err.Description, False)
    GoTo CleanUp
End Function

