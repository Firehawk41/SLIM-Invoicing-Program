VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInvoiceTypeSelector 
   Caption         =   "Invoicing Macro"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5520
   OleObjectBlob   =   "frmInvoiceTypeSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInvoiceTypeSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
' MacroName: frmInvoiceTypeSelector
' Version: 2.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-05
' Description: Form for users to choose which type of invoice to generate
' DependsOn: clsLoggingSystem, clsAccessDatabase
' ChangeLog:
'   - 2.0.1 - 2025-06-05 - Added metadata
Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_Cancelled As Boolean
Private m_SelectedFilePath As String
Public Property Get InvoiceType() As InvoiceTypeEnum
    If btnIndividual.Value = True Then
        InvoiceType = InvoiceTypeEnum.DefaultIndividual
    ElseIf cboCustomer.Text = "Customer A" Then
        InvoiceType = InvoiceTypeEnum.CustomerASummary
    ElseIf cboCustomer.Text = "Customer B Chemical" Then
        InvoiceType = InvoiceTypeEnum.CustomerBSummary
    End If
End Property
Public Property Get StartDate() As Date
' Only relevant for summary invoices
    If btnWeekly.Value Then
        StartDate = GetStartOfWeek(CInt(cboDate.Value))
    ElseIf btnMonthly.Value Then
        StartDate = GetFirstDayOfMonth(cboDate.Value)
    End If
End Property
Private Function GetStartOfWeek(WeekNum As Integer) As Date
    Dim YearStart As Date
    YearStart = DateSerial(Year(Now), 1, 1)
    GetStartOfWeek = DateAdd("ww", WeekNum - 1, YearStart)
    If Weekday(GetStartOfWeek, vbMonday) <> 1 Then
        GetStartOfWeek = DateAdd("d", -Weekday(GetStartOfWeek, vbMonday) + 1, GetStartOfWeek)
    End If
End Function
Private Function GetFirstDayOfMonth(monthAbbreviation As String) As Date
    GetFirstDayOfMonth = DateSerial(Year(Now), Month(DateValue(monthAbbreviation & " 1")), 1)
End Function
' Returns the file path selected by the user (only relevant for individual invoices)
Public Property Get SelectedFilePath() As String
    SelectedFilePath = m_SelectedFilePath
End Property
Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property
Public Sub Initialize(Logger As clsLoggingSystem, AccessDB As clsAccessDatabase)
    Call Logger.LogMessage("Initialize", LogLevelEnum.LogDebug, "Initializing frmInvoiceTypeSelector")

    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    
    Me.btnIndividual = True
    Me.cboErrorFile.Clear
    'Me.cboErrorFile.List = SetErrorFileList()
    If cboErrorFile.ListCount > 0 Then
        cboErrorFile.Value = cboErrorFile.List(0)
    End If
    
    Call m_Logger.LogMessage("Initialize", LogLevelEnum.LogDebug, "Userform initialization complete")
End Sub
Private Function SetErrorFileList() As Variant
    Dim SQLQuery As String
    Dim DataArray As Variant
    Dim ResultList() As String
    Dim iIndex As Integer
    Dim ListSize As Integer
    
    If Not m_Logger.DebugMode Then
        On Error GoTo ErrorHandler
    End If
    
    SQLQuery = "SELECT [Report_Invoiced] AS DisplayText FROM Invoice_Database " & _
           "WHERE [Error_Message] <> '' AND [Error_Message] <> 'No Error' " & _
           "AND [Error_Message] NOT LIKE 'Fixed%' AND [Report_Invoiced] <> '' AND [Generated_From] = 'Outlook' " & _
           "ORDER BY [Report_Invoiced] DESC"
    
    DataArray = m_AccessDB.ExecuteQuery(SQLQuery)
    
    ' Initialize the size of the result list
    ListSize = UBound(DataArray, 2) - LBound(DataArray, 2) + 1
    ReDim ResultList(1 To ListSize)  ' Use 1-based indexing for convenience
    
    ' Populate the ResultList array with each item from the DataArray
    For iIndex = LBound(DataArray, 2) To UBound(DataArray, 2)
        ResultList(iIndex - LBound(DataArray, 2) + 1) = DataArray(0, iIndex)  ' Adjust to 1-based indexing
    Next iIndex
    
    ' Return the array
    SetErrorFileList = ResultList
    Exit Function
ErrorHandler:
    Call m_Logger.LogError("SetErrorFileList", Err.Number, Err.Description)
End Function
Private Sub btnIndividual_Click()
    btnInvoiceType_Click Me.btnIndividual
End Sub
Private Sub btnMonthly_Click()
    btnInvoiceType_Click Me.btnMonthly
End Sub
Private Sub btnWeekly_Click()
    btnInvoiceType_Click Me.btnWeekly
End Sub

Private Sub btnInvoiceType_Click(Sender As Object)
    Const WEEKLY_INVOICE_CUSTOMERS As String = "Customer A"
    Const MONTHLY_INVOICE_CUSTOMERS As String = "Customer B Chemical"
    
    Call m_Logger.LogMessage("btnInvoiceType_Click", LogLevelEnum.LogDebug, "Invoice type changed: " & Sender.Caption)
    
    Select Case Sender.Name
        Case "btnIndividual"
            ShowInvoiceOptions True
            btnOpenFile.Value = True
        Case "btnWeekly"
            ShowInvoiceOptions False
            PopulateCustomerAndDateComboBoxes WEEKLY_INVOICE_CUSTOMERS, True
        Case "btnMonthly"
            ShowInvoiceOptions False
            PopulateCustomerAndDateComboBoxes MONTHLY_INVOICE_CUSTOMERS, False
    End Select
End Sub

Private Sub PopulateCustomerAndDateComboBoxes(Customers As String, isWeekly As Boolean)
    Const LIST_OF_MONTHS As String = "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec"
    Const WEEKLY_CAPTION As String = "Which week would you like to invoice for?"
    Const MONTHLY_CAPTION As String = "Which month would you like to invoice for?"
    
    cboDate.Clear
    
    PopulateComboBox cboCustomer, Split(Customers, ",")
    If cboCustomer.ListCount > 0 Then
        cboCustomer.Value = cboCustomer.List(0)
    End If
    
    If isWeekly Then
        lblDate.Caption = WEEKLY_CAPTION
        PopulateComboBox cboDate, SetWeekList()
        If cboDate.ListCount > 0 Then
            cboDate.Value = cboDate.List(0)
        End If
        lblWeeklyDateRange.Visible = True
    Else
        lblDate.Caption = MONTHLY_CAPTION
        PopulateComboBox cboDate, Split(LIST_OF_MONTHS, ",")
        cboDate.Value = format(DateAdd("m", -1, Date), "mmm")
        lblWeeklyDateRange.Visible = False
    End If
End Sub

Private Sub PopulateComboBox(cbo As Object, List As Variant)
    Dim iIndex As Long
    
    cbo.Clear
    For iIndex = LBound(List) To UBound(List)
        cbo.AddItem List(iIndex)
    Next iIndex
End Sub

Private Sub btnOpenFile_Click()
    Call m_Logger.LogMessage("btnOpenFile_Click", LogLevelEnum.LogDebug, "Opening file from user input.")
    cboErrorFile.Enabled = False
End Sub
Private Sub btnSelectErrorFile_Click()
    Call m_Logger.LogMessage("btnSelectErrorFile_Click", LogLevelEnum.LogDebug, "Error file selection enabled.")
    cboErrorFile.Enabled = True
End Sub
Private Sub ShowInvoiceOptions(ShowIndividual As Boolean)
    IndividualInvoiceFrame.Visible = ShowIndividual
    SummaryInvoiceFrame.Visible = Not ShowIndividual
End Sub
Private Function SetWeekList() As Variant
    Dim WeekSelectArr(1 To 10) As Integer
    Dim aIndex As Integer
    Dim CurrentWeek As Integer

    CurrentWeek = GetWeekNumber(Now)
    
    For aIndex = 1 To UBound(WeekSelectArr)
        WeekSelectArr(aIndex) = IIf(CurrentWeek - aIndex < 1, CurrentWeek - aIndex + 52, CurrentWeek - aIndex)
    Next aIndex
    
    ' Return the array to the calling sub
    SetWeekList = WeekSelectArr
End Function
Private Function GetWeekNumber(InputDate As Date) As Integer
    Dim Jan1 As Date
    Dim WeekOneMonday As Date
    Dim WeekNum As Integer

    ' Determine January 1 for the given year.
    Jan1 = DateSerial(Year(InputDate), 1, 1)
    
    ' If January 1 is Monday-Friday, then the week that contains Jan 1 is week one.
    ' That week starts on the Monday before (or on) Jan 1.
    ' Otherwise, if Jan 1 is Saturday or Sunday, the first Monday-Friday week is the one
    ' beginning on the following Monday.
    If Weekday(Jan1, vbMonday) <= 5 Then
        ' For a weekday, go backward to the Monday of that week.
        WeekOneMonday = DateAdd("d", -(Weekday(Jan1, vbMonday) - 1), Jan1)
    Else
        ' For a weekend day, go forward to the next Monday.
        WeekOneMonday = DateAdd("d", 8 - Weekday(Jan1, vbMonday), Jan1)
    End If

    ' Calculate the week number by counting the number of weeks between the starting Monday
    ' and the InputDate. The +1 ensures the starting week is counted as week 1.
    WeekNum = DateDiff("ww", WeekOneMonday, InputDate, vbMonday) + 1

    GetWeekNumber = WeekNum
End Function
Private Sub CancelButton_Click()
    m_Cancelled = True
    Me.Hide
End Sub
Private Sub cboDate_Change()
    If btnWeekly.Value And Not cboDate.Value = "" Then
        Dim selectedWeek As Integer
        selectedWeek = CInt(cboDate.Value)
        lblWeeklyDateRange.Caption = "Week " & selectedWeek & ": " & format(GetStartOfWeek(selectedWeek), "mmm dd") & " - " & format(GetStartOfWeek(selectedWeek) + 6, "mmm dd")
        lblWeeklyDateRange.Visible = True
    End If
End Sub

Private Sub OKButton_Click()
    Call m_Logger.LogMessage("OKButton_Click", LogLevelEnum.LogDebug, "OK button clicked. Processing invoice.")
    
    ' Handle based on the selected invoice type
    Select Case True
        Case btnIndividual.Value
            If btnOpenFile.Value Then
                ' Let the user select a file for individual invoices
                m_SelectedFilePath = Application.GetOpenFilename(FileFilter:="Excel files (*.xlsx), *.xlsx")
                If m_SelectedFilePath = "False.xlsx" Then
                    Call m_Logger.LogMessage("OKButton_Click", LogLevelEnum.LogInfo, "User cancelled file selection")
                    m_Cancelled = True ' Cancel if no file selected
                Else
                    Call m_Logger.LogMessage("OKButton_Click", LogLevelEnum.LogInfo, "File selected: " & m_SelectedFilePath)
                End If
                Me.Hide
            ElseIf btnSelectErrorFile.Value Then
                ' Place holder
                Call m_Logger.LogMessage("OKButton_Click", LogLevelEnum.LogDebug, "Select Error File button chosen. This option is not yet supported.", , True)
            End If
        Case btnWeekly.Value Or btnMonthly.Value
            ' Check if customer or date are empty
            If Len(cboCustomer.Value) = 0 Or Len(cboDate.Value) = 0 Then
                Call m_Logger.LogMessage("OKButton_Click", LogLevelEnum.LogWarning, "Invoice cannot be processed: Missing customer or date.", , True)
            Else
                Call m_Logger.LogMessage("OKButton_Click", LogLevelEnum.LogInfo, "Invoice processed for customer: " & cboCustomer.Value & " with date: " & cboDate.Value)
                Me.Hide
            End If
            
        Case Else
            ' Default case
            Call m_Logger.LogMessage("OKButton_Click", LogLevelEnum.LogWarning, "Invoice processing cancelled or invalid.", , True)
    End Select
End Sub
