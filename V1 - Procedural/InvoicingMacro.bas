Attribute VB_Name = "InvoicingMacro"
Option Explicit
Option Compare Text

Sub Create_Invoice_From_Excel()
    'This macro will open the userform to select the file to invoice and then call the InvoiceMacro. This is separate so that the InvoiceMacro can be called from Outlook.

    'Version 1.0    -   11/30/22    - Initial release
    'Version 1.1    -   5/1/23      - Changed where CustomerX reports would be saved
    'Version 2.0    -   5/3/23      - Added compatibility with invoicing reports that caused errors

    Dim WrkBk As Workbook
    Dim wrdApp As Word.Application
    Dim WrdDoc As Word.Document
    Dim ErrMsg As String
    Dim FolderPath, sSaveFolder, File_Name, Open_This_File As String
    Dim StartDate, EndDate, File_Date As Date
    Dim fdObj As Object
    Dim Day_Index As Integer
    Dim CustomerX_Reports As New Collection
    Dim Process_Date As Date

    Set fdObj = CreateObject("Scripting.FileSystemObject")


    Call MacroOptimization
    'Shows userform to select which file to invoice from
    With OpenFileUserForm
        .Show
        Set wrdApp = CreateObject("Word.Application")
        Process_Date = Now
        'Exits if the cancel button is clicked
        If .Cancelling = True Then
            wrdApp.Quit
            Call ResetSettings
            Exit Sub
            'if a file is selected, open it
        ElseIf .OpenFileButton = True Then Set WrkBk = Workbooks.Open(.FileString)
            'if a summary invoice is selected, a collection of reports is made and the first one opened to pass to the invoicing macro
        ElseIf .SummaryInvoiceButton = True Then
            FolderPath = "\\PRECILAB-SERVER\LabPlusServer\Z-Reports\CustomerX-" & Mid(.CustomerSelectComboBox, 14, 2) & "\"
            StartDate = DateSerial(Mid(.DateRangeLabel, 19, 4), Mid(.DateRangeLabel, 13, 2), Mid(.DateRangeLabel, 16, 2))
            EndDate = DateSerial(Mid(.DateRangeLabel, 32, 4), Mid(.DateRangeLabel, 26, 2), Mid(.DateRangeLabel, 29, 2))
            'If there are no reports for the selected week, the macro will cancel
            Set CustomerX_Reports = Collect_CustomerX_Reports(StartDate, FolderPath)
            If CustomerX_Reports.Count = 0 Then
                MsgBox "There are no " & .CustomerSelectComboBox.Text & " reports for the selected week.", vbInformation, "Invoicing Macro"
                wrdApp.Quit
                Call ResetSettings(WrkBk)
                Exit Sub
            End If
            'otherwise, it will open the first report to pass onto the invoice function
            If Not CustomerX_Reports.Count = 0 Then Set WrkBk = Workbooks.Open(CustomerX_Reports(1))
            'if Invoicing a File That Caused an Error is selected, open that file.
        ElseIf .Invoice_An_Error_Button = True Then
            FolderPath = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Saved Invoices\Reports That Caused Errors\"
            File_Name = .Invoice_An_Error_Array(.Errors_Found_Listbox.ListIndex, 0)
            Open_This_File = FolderPath & File_Name
            'The process date is set to the date the error occurred
            Process_Date = .Invoice_An_Error_Array(.Errors_Found_Listbox.ListIndex, 3)
            'If the report isn't in the "Reports That Caused Errors" folder, prompt the user to look for it
            If Dir(Open_This_File) = "" Then
                MsgBox "Cannot find the specified file in the ""Reports That Caused Errors"" folder." & vbNewLine & vbNewLine & "Please navigate to the correct report.", vbExclamation, "Invoicing Macro"
                Open_This_File = Application.GetOpenFilename(FileFilter:="Excel files (*.xlsx), *.xlsx")
                'Notify the user and exit the sub if the workbook selected doesn't match the name of the workbook that caused the error
                If Not Right(Open_This_File, Len(Open_This_File) - InStrRev(Open_This_File, "\")) = File_Name Then
                    MsgBox "The file you selected is not the same file that caused an error.", vbExclamation, "Invoicing Macro"
                    wrdApp.Quit
                    Call ResetSettings(WrkBk)
                    Exit Sub
                End If
            End If
            'Open the report to be invoiced
            Set WrkBk = Workbooks.Open(Open_This_File)
        End If

        'Call the invoice function
        Set WrdDoc = Create_Word_Invoice(WrkBk, wrdApp, ErrMsg, "Excel", Process_Date)
        'if a word file exists (ie the function didn't error out), make it visible and display it
        If Not WrdDoc Is Nothing Then
            wrdApp.Visible = True
            WrdDoc.Application.Activate
            'Save the file the folder labelled as today's date
            sSaveFolder = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Saved Invoices\" & Format(Process_Date, "mm-dd-yyyy") & "\"
            If Not fdObj.folderexists(sSaveFolder) Then MkDir (sSaveFolder)
            WrdDoc.SaveAs2 sSaveFolder & WrdDoc.ActiveWindow.Caption & ".docx"
            'if invoicing an error, update the error code to begin with "Fixed" and the date and delete the report if it's in the "Reports That Caused Errors" folder
            If .Invoice_An_Error_Button = True And Left(Open_This_File, InStrRev(Open_This_File, "\")) = FolderPath Then fdObj.DeleteFile FolderPath & File_Name
            If .Invoice_An_Error_Button = True Then Call Insert_Value_Into_Database("UPDATE Invoice_Database SET [Error_Message] = 'Fixed " & Format(Now, "mm-dd-yyyy") & ": " & .Invoice_An_Error_Array(.Errors_Found_Listbox.ListIndex, 2) & "' WHERE [ID] = " & .Invoice_An_Error_Array(.Errors_Found_Listbox.ListIndex, 1))
        End If
    End With
    'if the word file doesn't exist, quit the Word instance
    If WrdDoc Is Nothing Then wrdApp.Quit


    'Display a message box if an error occurred
    If ErrMsg <> "" Then MsgBox "An error has occured and the invoice wasn't completed." & vbNewLine & vbNewLine & "This has been logged but please let Jamie know and he will troubleshoot the problem." & vbNewLine & vbNewLine & ErrMsg, , "Invoicing Macro (passed successfully)"
    'Reset settings and close workbooks and userforms
    Call ResetSettings(WrkBk, "Excel")
End Sub
Private Sub MacroOptimization()
    'Turns off Screen Updating, Enable Events and turns calculations to Manual in order to speed up running time of macros
    'Written by JT

    'Version 1.0    - 10/3/22    - Initial release

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
End Sub
Private Sub ResetSettings(Optional WrkBk As Workbook, Optional ByVal Requested_By As String)
    'Turns on Screen Updating, Enable Events and turns calculations to Manual to restore settings after a macro has finished running
    'Written by JT

    'Version 1.0    - 10/3/22   - Initial release
    'Version 1.1    - 11/30/22  - removed references to specific userforms and replaced with a loop to close all userforms
    'Version 1.2    - 1/19/23   - Saves the log book from the backup folder to the regular folder.  This is to prevent errors if the logbook is open
    'Version 1.3    - 2/9/23    - Changed the way the log book works - no longer saves from back up but saves to a master file which itself is copied to a working Logbook
    'Version 1.4    - 5/1/23    - Removed reference to the Excel Invoice Logbook as the database has been moved to Access

    Dim myForm As UserForm

    'closes the workbook that is being invoiced
    On Error Resume Next
    Application.DisplayAlerts = False
    WrkBk.Close savechanges:=False
    Application.DisplayAlerts = True
    'Closes all active userforms
    For Each myForm In UserForms
        Unload myForm
    Next myForm
    On Error GoTo 0
    'Resets settings
    With Application
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With

End Sub

Function Create_Word_Invoice(Report_Wbk As Workbook, wrdApp As Word.Application, ByRef Error_Message As String, ByVal RequestedBy As String, Optional ByVal Process_Date As Date) As Word.Document
    Dim iIndex, aIndex As Integer
    Dim Analyses As Collection
    Dim Testing_Row_Index, Report_Collection_Index, Array_Width, Processing_Time_Column, Payment_Column, Array_Index, Sample_String_Count, Payment_Terms As Integer
    Dim Invoice_Week, First_Week_Correction, Bottom_Boundary, Row_Index As Integer
    Dim Customer, Pricing_Index, Processing_Time_Index, Sample_String, Sample_ID, Sample_Matrix, PO_Number, File_Path, Logo_Location As String
    Dim Payment_Details, Total_Cost, Invoice_Number, Suggested_File_Name, File_Name, Pricing_Database, Processing_Time  As String
    Dim Testing_Request, WkS As Worksheet
    Dim Prices() As Variant
    Dim Real_Names() As Variant
    Dim Invoicing_Matrix() As Variant
    Dim Chemicals_List() As String
    Dim Sample_String_Array() As String
    Dim Sub_Total As Long
    Dim Report_Collection As New Collection
    Dim CustomerX_Invoice, Chemical_Found, Analyses_Processing_Time_Found As Boolean
    Dim Date_Samples_Received As Date
    Dim Wbk As Workbook
    Dim Word_Invoice As Word.Document

    'On Error GoTo Err_Handler:


    'Loops the workbook to find the testing request in order to find the customer
    Set Testing_Request = Find_Testing_Sheet(Report_Wbk)
    'If the testing request isn't found, exit the function
    If Testing_Request Is Nothing Then
        Call ResetSettings(Report_Wbk, RequestedBy)
        Exit Function
    End If

    'searches for the customer using custom function
    Customer = Find_Customer(Testing_Request, RequestedBy, Error_Message)
    'If the customer wasn't found cancel the function
    If Customer = "" Then
        Call Fill_Out_Logbook(Customer, Invoice_Number, Report_Wbk.Name, Total_Cost, RequestedBy, Error_Message, Process_Date)
        Call ResetSettings(Report_Wbk)
        Exit Function
    End If


    'Sets variables for CustomerX and everyone else
    ReDim Invoicing_Matrix(9, 0)
    If Left(Customer, 12) = "CustomerX" Then
        'This is the date of the file selected
        Date_Samples_Received = DateSerial(Mid(Report_Wbk.Name, 5, 2), Left(Report_Wbk.Name, 2), Mid(Report_Wbk.Name, 3, 2))
        'Either CustomerX-TX or CustomerX-CA
        File_Path = "\\PRECILAB-SERVER\LabPlusServer\Z-reports\CustomerX-" & Mid(Customer, 14, 2) & "\"
        'Adds the latest version of each report for the week in a collection
        Set Report_Collection = Collect_CustomerX_Reports(Date_Samples_Received, File_Path)
        CustomerX_Invoice = True
        Invoicing_Matrix(0, 0) = "Testing Request #"
        Invoicing_Matrix(1, 0) = "Quantity"
        Invoicing_Matrix(2, 0) = "Description"
        Invoicing_Matrix(3, 0) = "Unit Price"
        Invoicing_Matrix(4, 0) = "Extension"
        Invoicing_Matrix(5, 0) = "Testing Request #"
        Invoicing_Matrix(6, 0) = "Sample Strings"
        Invoicing_Matrix(9, 0) = "Processing Time Index/Description"
        'Variables for non-CustomerX customers
    Else
        Report_Collection.Add Report_Wbk.Name
        Logo_Location = Get_Logo(Customer, RequestedBy)
        CustomerX_Invoice = False
        Invoicing_Matrix(0, 0) = "Quantity"
        Invoicing_Matrix(1, 0) = "Description"
        Invoicing_Matrix(2, 0) = "Turn-Around Time"
        Invoicing_Matrix(3, 0) = "Unit Price"
        Invoicing_Matrix(4, 0) = "Extension"
        Invoicing_Matrix(5, 0) = "Testing Request #"
        Invoicing_Matrix(6, 0) = "Sample Strings"
        Invoicing_Matrix(7, 0) = "Test/Turn around combo"
        Invoicing_Matrix(8, 0) = "Chemicals"
        Invoicing_Matrix(9, 0) = "Processing Time Index/Description"
    End If

    'Loop through each file in the Report Collection (which is only one file if not a CustomerX summary invoice)
    For Report_Collection_Index = 1 To Report_Collection.Count
        Set Wbk = Open_Workbook(Report_Collection(Report_Collection_Index))
        Set WkS = Find_Testing_Sheet(Wbk)
        File_Name = Wbk.Name
        File_Path = Wbk.Path

        PO_Number = WkS.Range("I10")
        'Sets variables depending upon which testing request form is found
        If WkS.Range("D3") = "Chemical Testing Request Form" Then
            Pricing_Database = "Chemical_Pricing"
            Date_Samples_Received = WkS.Range("P5")
            Processing_Time_Column = 11
            Payment_Column = 9
        ElseIf WkS.Range("D3") = "Water Testing Request Form" Then
            Pricing_Database = "Water_Pricing"
            Date_Samples_Received = WkS.Range("S5")
            Processing_Time_Column = 14
            Payment_Column = 10
        ElseIf WkS.Range("D3") = "Wafer Testing Request Form" Then
            Pricing_Database = "Wafer_Pricing"
            Date_Samples_Received = WkS.Range("L5")
            Processing_Time_Column = 9
            Payment_Column = 7
            'exit the for loop if a testing request isn't found
        Else
            If Not RequestedBy = "Outlook" Then MsgBox "Testing request not selected"
            Exit For
        End If
        Bottom_Boundary = WkS.Cells(Rows.Count, 2).End(xlUp).Row - 11
        For Row_Index = 21 To WkS.Cells(Rows.Count, 2).End(xlUp).Row - 10
            If WkS.Cells(Row_Index, 2) = "Examples" And WkS.Cells(Row_Index, 2).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone Then Bottom_Boundary = Row_Index - 1
        Next Row_Index


        'Finds which row the customer's prices are on
        Pricing_Index = Retrieve_Database_Reference("Select [Pricing_Key] from Customer_Database WHERE [Customer_Name] = '" & Customer & "'", "Pricing_Key")

        'Loops through each row of the testing request to collect the analyses information
        For Testing_Row_Index = 21 To Bottom_Boundary
            Sample_ID = WkS.Cells(Testing_Row_Index, 2)
            If WkS.Range("D3") = "Chemical Testing Request Form" Then Sample_Matrix = Replace(WkS.Cells(Testing_Row_Index, 3), ",", " ")
            If WkS.Range("D3") = "Water Testing Request Form" Then Sample_Matrix = "Water"
            If WkS.Range("D3") = "Wafer Testing Request Form" Then Sample_Matrix = WkS.Cells(Testing_Row_Index, 4) & " Wafer"
            'Skips this row if the sample ID is blank
            If Not Sample_ID = "" Then
                Set Analyses = Collect_Analyses(WkS, Testing_Row_Index, Pricing_Index, Pricing_Database, Processing_Time_Column)
                If Not Analyses.Count = 0 Then
                    Processing_Time = Replace(WkS.Cells(Testing_Row_Index, Processing_Time_Column), " ", "_") & "_Price"
                    If Processing_Time = "_Price" And Not Pricing_Database = "Wafer_Pricing" Then Processing_Time = "Next_Day_Price"
                    If Processing_Time = "_Price" And Pricing_Database = "Wafer_Pricing" Then Processing_Time = "Up_To_3_Working_Days_Price"
                    Prices = Collect_Prices(Analyses, Pricing_Index, Pricing_Database, Processing_Time)
                    Sub_Total = Sum_Prices(Prices)
                    Real_Names = Display_Analyses_Names(Analyses)
                    Sample_String = Format(Date_Samples_Received, "mmddyy") & "-" & Sample_Matrix & "-" & Customer & "-" & Sample_ID
                    If CustomerX_Invoice = True Then
                        'creates CustomerX invoicing_Matrix for this sample
                        For iIndex = 0 To UBound(Real_Names)
                            Array_Width = UBound(Invoicing_Matrix, 2) + 1
                            ReDim Preserve Invoicing_Matrix(9, Array_Width)
                            Invoicing_Matrix(0, Array_Width) = PO_Number
                            Invoicing_Matrix(1, Array_Width) = "1 EA"
                            Invoicing_Matrix(2, Array_Width) = Sample_String & vbNewLine & Real_Names(iIndex)
                            If Processing_Time Like "*Rush*" Or Processing_Time Like "*Time_Limited*" Then Invoicing_Matrix(2, Array_Width) = Invoicing_Matrix(2, Array_Width) & " " & WkS.Cells(Testing_Row_Index, Processing_Time_Column)
                            Invoicing_Matrix(3, Array_Width) = Format(Prices(iIndex), "#,##0.00") & " per 1 EA"
                            Invoicing_Matrix(4, Array_Width) = Format(Prices(iIndex), "#,##0.00")
                            Invoicing_Matrix(5, Array_Width) = PO_Number
                            'Leaves the sample string blank if this sample is already in the matrix
                            If Not Sample_String = Invoicing_Matrix(6, Array_Width - 1) Then Invoicing_Matrix(6, Array_Width) = Sample_String
                            Invoicing_Matrix(9, Array_Width) = WkS.Cells(Testing_Row_Index, Processing_Time_Column)
                        Next iIndex
                        'creates Non-CustomerX Invoicing_Matrix for this sample
                    Else
                        For Array_Index = 0 To UBound(Invoicing_Matrix, 2)
                            Analyses_Processing_Time_Found = False
                            Chemical_Found = False
                            'If the analyses + processing time has already been found, this sample is added to that row
                            If Join(Real_Names, ", ") & Processing_Time = Invoicing_Matrix(7, Array_Index) Then
                                Analyses_Processing_Time_Found = True
                                Invoicing_Matrix(0, Array_Index) = Invoicing_Matrix(0, Array_Index) + 1
                                Invoicing_Matrix(6, Array_Index) = Invoicing_Matrix(6, Array_Index) & "," & Sample_String
                                Chemicals_List = Split(Invoicing_Matrix(8, Array_Index), ", ")
                                For iIndex = 0 To UBound(Chemicals_List)
                                    If Sample_Matrix = Chemicals_List(iIndex) Then Chemical_Found = True
                                Next iIndex
                                If Chemical_Found = False Then Invoicing_Matrix(8, Array_Index) = Invoicing_Matrix(8, Array_Index) & ", " & Sample_Matrix
                                Exit For
                            End If
                        Next Array_Index
                        'If the analyses + processing time has NOT been found, a new row is added
                        If Analyses_Processing_Time_Found = False Then
                            Array_Width = UBound(Invoicing_Matrix, 2) + 1
                            ReDim Preserve Invoicing_Matrix(9, Array_Width)
                            Invoicing_Matrix(0, Array_Width) = 1
                            Invoicing_Matrix(1, Array_Width) = vbNewLine & vbTab & Join(Real_Names, vbNewLine & vbTab)
                            Invoicing_Matrix(2, Array_Width) = WkS.Cells(Testing_Row_Index, Processing_Time_Column)
                            Invoicing_Matrix(3, Array_Width) = vbNewLine & Join(Prices, vbNewLine)
                            Invoicing_Matrix(4, Array_Width) = Format(Sub_Total, "#,##0.00")
                            Invoicing_Matrix(6, Array_Width) = Sample_String
                            Invoicing_Matrix(7, Array_Width) = Join(Real_Names, ", ") & Processing_Time
                            Invoicing_Matrix(8, Array_Width) = Sample_Matrix
                            Invoicing_Matrix(9, Array_Width) = WkS.Cells(Testing_Row_Index, Processing_Time_Column)
                        End If
                    End If
                End If
            End If
        Next Testing_Row_Index
        'Sets payment variables - whether payment was by credit card and also the payment details ie CC number or the PO number
        If Not Customer Like "CustomerX*" Then Payment_Details = Get_Payment_Details(WkS, Customer, Payment_Column)
        'Closes the report/testing request
        Wbk.Close savechanges:=False
    Next Report_Collection_Index


    If CustomerX_Invoice = False Then
        'Formats the analysis description to the way Yves would like
        Invoicing_Matrix = Finish_Description(Invoicing_Matrix)
        Set Word_Invoice = wrdApp.Documents.Add(Template:="\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Templates and Logbooks\Invoicing Template.dotx", NewTemplate:=False, DocumentType:=0)
        'Adds Non-CustomerX variables to the invoice
        Call Add_Non_CustomerX_Variables(Word_Invoice, Customer, Logo_Location)
        Invoice_Number = Get_Invoice_Number(Customer, Date_Samples_Received, RequestedBy)
        'Sets the file name
        If Payment_Details Like "Credit Card*" Then Suggested_File_Name = "PRECILAB invoice receipt " & Invoice_Number & " to " & Customer
        If Not Payment_Details Like "Credit Card*" Then Suggested_File_Name = "PRECILAB invoice " & Invoice_Number & " to " & Customer & " " & Payment_Details
        If Customer Like "Monument*" Then Suggested_File_Name = "PRECILAB invoice " & Invoice_Number & " to " & Customer & " [" & Format(Now, "mm-dd-yyyy") & "]"
    ElseIf CustomerX_Invoice = True Then
        'Creates word document from template
        Set Word_Invoice = wrdApp.Documents.Add(Template:="\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Templates and Logbooks\CustomerX Invoicing Template.dotx", NewTemplate:=False, DocumentType:=0)
        If Report_Collection.Count > 1 Then Word_Invoice.Tables(2).Cell(2, 1).Range.Text = "From " & Invoicing_Matrix(0, 1) & " to " & Invoicing_Matrix(0, UBound(Invoicing_Matrix, 2)) & " inclusive"
        If Report_Collection.Count = 1 Then Word_Invoice.Tables(2).Cell(2, 1).Range.Text = Invoicing_Matrix(0, 1)
        'The date of the first report of the week
        Date_Samples_Received = Format(Date_Samples_Received - Weekday(Date_Samples_Received) + 2, "mm-dd-yyyy")
        If Customer = "CustomerX TX" Then Invoice_Number = Str(Format(Date_Samples_Received, "yyyymmdd") & "1")
        If Customer = "CustomerX CA" Then Invoice_Number = Str(Format(Date_Samples_Received, "yyyymmdd") & "2")
        'For Excel, week 1 starts on January 1st. For PRECILAB, week 1 starts on the 1st Monday of the year
        'This "FirstWeekCorrection" corrects for this difference by reducing the week number by one if
        'January 1st is not Sunday or Monday
        If Weekday(DateSerial(Year(Date_Samples_Received), 1, 1)) > 3 Then First_Week_Correction = 1
        Invoice_Week = WorksheetFunction.WeekNum(Date_Samples_Received) - First_Week_Correction
        Suggested_File_Name = "PRECILAB invoice " & Invoice_Number & " to CustomerX-" & Mid(Customer, 14, 2) & " week " & Invoice_Week & "  " & Format(Date_Samples_Received, "mm-dd-yy")
    End If
    'Fills in Invoicing table (3) and Sample ID table (5)
    If Not UBound(Invoicing_Matrix, 2) = 0 Then
        For iIndex = 1 To UBound(Invoicing_Matrix, 2)
            With Word_Invoice.Tables(3)
                If iIndex = .Rows.Count Then .Rows.Add
                .Cell(iIndex + 1, 1).Range.Text = Invoicing_Matrix(0, iIndex)
                .Cell(iIndex + 1, 2).Range.Text = Invoicing_Matrix(1, iIndex)
                .Cell(iIndex + 1, 3).Range.Text = Invoicing_Matrix(2, iIndex)
                .Cell(iIndex + 1, 4).Range.Text = Invoicing_Matrix(3, iIndex)
                .Cell(iIndex + 1, 5).Range.Text = Invoicing_Matrix(4, iIndex)
                'colors alternating reports (CustomerX) or lines (non-CustomerX) blue to make it easier for the customer to read
                If Customer Like "CustomerX*" And Invoicing_Matrix(0, iIndex) Mod 2 = 0 Then .Rows(iIndex + 1).Range.Font.TextColor = wdColorBlue
                If Customer Like "CustomerX*" And Not Invoicing_Matrix(0, iIndex) Mod 2 = 0 Then .Rows(iIndex + 1).Range.Font.TextColor = wdColorBlack
                If Not Customer Like "CustomerX*" And iIndex Mod 2 = 1 Then .Rows(iIndex + 1).Range.Font.TextColor = wdColorBlue
                If Not Customer Like "CustomerX*" And Not iIndex Mod 2 = 1 Then .Rows(iIndex + 1).Range.Font.TextColor = wdColorBlack
                'Colors non-routine processing times
                If Invoicing_Matrix(9, iIndex) Like "*RUSH*" Or Invoicing_Matrix(9, iIndex) Like "*Time Limited*" Then
                    .Cell(iIndex + 1, 4).Range.Font.TextColor = wdColorRed 'colors unit price
                    If Not Customer Like "CustomerX*" Then .Cell(iIndex + 1, 3).Range.Font.TextColor = wdColorRed 'colors processing times for non-CustomerX
                    If Customer Like "CustomerX*" Then 'for CustomerX samples, loops through each character of the processing time to turn it red
                        For aIndex = Len(Invoicing_Matrix(2, iIndex)) - Len(Invoicing_Matrix(9, iIndex)) To Len(Invoicing_Matrix(2, iIndex))
                            .Cell(iIndex + 1, 3).Range.Characters(aIndex).Font.TextColor = wdColorRed
                        Next aIndex
                    End If
                End If
            End With
            If Not Invoicing_Matrix(6, iIndex) = "" Then
                With Word_Invoice.Tables(5)
                    'if multiple samples have the same analyses/turn around time combination, they are grouped together
                    Sample_String_Array = Split(Invoicing_Matrix(6, iIndex), ",")
                    For aIndex = 0 To UBound(Sample_String_Array)
                        If Sample_String_Count = .Rows.Count Then Word_Invoice.Tables(5).Rows.Add
                        Sample_String_Count = Sample_String_Count + 1
                        .Cell(Sample_String_Count, 1).Range.Text = Invoicing_Matrix(5, iIndex)
                        .Cell(Sample_String_Count, 2).Range.Text = Sample_String_Array(aIndex)
                        'colors alternating reports (CustomerX) or lines (non-CustomerX) blue to make it easier for the customer to read
                        If Customer Like "CustomerX*" And Invoicing_Matrix(5, iIndex) Mod 2 = 0 Then .Rows(Sample_String_Count).Range.Font.TextColor = wdColorBlue
                        If Customer Like "CustomerX*" And Not Invoicing_Matrix(5, iIndex) Mod 2 = 0 Then .Rows(Sample_String_Count).Range.Font.TextColor = wdColorBlack
                        If Not Customer Like "CustomerX*" And iIndex Mod 2 = 1 Then .Rows(Sample_String_Count).Range.Font.TextColor = wdColorBlue
                        If Not Customer Like "CustomerX*" And Not iIndex Mod 2 = 1 Then .Rows(Sample_String_Count).Range.Font.TextColor = wdColorBlack
                    Next aIndex
                End With
            End If
        Next iIndex
    End If
    'Fills out payment informat, cost and last dates
    With Word_Invoice
        If Payment_Details Like "Credit Card*" Then
            .Variables("PONumber") = " "
            .Tables(6).Cell(1, 1).Range.Text = "DO NOT PAY" & vbNewLine & "Already paid by " & Payment_Details & " on " & Format(Now, "mm-dd-yyyy")
            .Tables(2).Cell(2, 4).Range.Text = "N.A."
            .Tables(2).Cell(2, 3).Range.Text = "Credit Card"
            'ie if Payment_Details Like "P.O.*" Then
        Else
            .Tables(6).Rows(1).Delete
            .Variables("PONumber") = Payment_Details
            
            
            Payment_Terms = Retrieve_Database_Reference("Select [Payment_Terms] from Customer_Database WHERE [Customer_Name] = '" & Customer & "'", "Payment_Terms")
            If Payment_Terms = 0 Then
                .Tables(2).Cell(2, 3).Range.Text = "At Reception"
                .Tables(2).Cell(2, 4).Range.Text = "ASAP"
            Else
                .Tables(2).Cell(2, 3).Range.Text = "Net " & CInt(Payment_Terms) & " Days"
                If CustomerX_Invoice = True Then .Tables(2).Cell(2, 4).Range.Text = Format(Date_Samples_Received + Payment_Terms, "mm-dd-yyyy")
                If CustomerX_Invoice = False Then .Tables(2).Cell(2, 4).Range.Text = Format(Now + Payment_Terms, "mm-dd-yyyy")
            End If
        End If
        Total_Cost = Get_Total(Invoicing_Matrix)
        .Tables(4).Cell(1, 2).Range.Text = Total_Cost
        .Tables(4).Cell(3, 2).Range.Text = Total_Cost
        .Tables(2).Cell(2, 2).Range.Text = Format(Now, "mm-dd-yyyy")
        .Variables("CustomerAddress") = Customer_Details(Customer, "Address1", "Address2", "Address3", "Address4")
        .Variables("InvoiceDate") = Format(Date_Samples_Received, "mm-dd-yyyy")
        .Variables("InvoiceNumber") = Invoice_Number
        'activate and close print preview due to a bug that won't refresh changes otherwise
        .PrintPreview
        .ClosePrintPreview
        .Fields.Update
        'get rid of characters that aren't compatible with file names
        Suggested_File_Name = Replace(Replace(Replace(Replace(Replace(Suggested_File_Name, ".", ""), "# :", ""), "-", Chr(45)), "#:", ""), "/", "")
        'Make the title and default Save As name the Suggested File Name
        With wrdApp.Dialogs(wdDialogFileSummaryInfo)
            .Title = Suggested_File_Name
            .Execute
        End With
        .ActiveWindow.Caption = Suggested_File_Name
    End With

Error_Resume:
    'Fills out the Invoice Logbook
    Call Fill_Out_Logbook(Customer, Invoice_Number, File_Name, Total_Cost, RequestedBy, Error_Message, Process_Date, Invoice_Week, Suggested_File_Name)
    If Not RequestedBy = "Excel" Then Call ResetSettings(Report_Wbk, RequestedBy)
    Set Create_Word_Invoice = Word_Invoice
    Exit Function
Err_Handler:
    Error_Message = "Error Number: " & Err.Number & vbNewLine & "Error Description: " & Err.Description
    'If RequestedBy = "Excel" Then Resume next
    If RequestedBy = "Outlook" Then GoTo Error_Resume:
End Function
Private Function Open_Workbook(ByVal Workbook_To_Invoice As String) As Workbook
    'Loops through each open workbook to see the workbook that we want to invoice is already open. If not, the function will open it.

    Dim File_Open As Boolean
    Dim Wb, Wbk As Workbook

    'Loops through each open workbook
    For Each Wb In Workbooks
        File_Open = False
        If Wb.Name = Workbook_To_Invoice Then
            File_Open = True
            Set Wbk = Wb
            Exit For
        End If
    Next Wb
    'If the file is not open, open it and assign it to "Wbk"
    If File_Open = False Then Set Wbk = Workbooks.Open(Workbook_To_Invoice)
    'Set the function = "Wbk"
    Set Open_Workbook = Wbk
End Function
Function Retrieve_Database_Reference(ByVal strQry As String, ByVal Database_Column As String) As String

    Dim strPath, strProv, strConn As String
    Dim Conn As New Connection
    Dim rsQry As New Recordset

    'If the Access connection is open, close it
    If rsQry.State = adStateOpen Then rsQry.Close

    strPath = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Templates and Logbooks\Invoice and Pricing Databases.accdb"
    strProv = "Microsoft.ACE.OLEDB.12.0;"
    strConn = "Provider=" & strProv & "Data Source=" & strPath

    'Connection open
    Conn.Open strConn

    'Looks for a match between the testing request company and address in the database
    rsQry.Open strQry, Conn
    If rsQry.EOF = False And IsNull(rsQry(Database_Column)) = False Then Retrieve_Database_Reference = rsQry(Database_Column)

    'close the connection
    If rsQry.State = adStateOpen Then rsQry.Close
    Conn.Close

End Function
Sub message()
MsgBox Retrieve_Database_Array("SELECT [PRECILAB_name] FROM [Chemical_Database] WHERE [Customer_name] = 'DMSO'")(0, 0)
End Sub

Function Retrieve_Database_Array(ByVal strQry As String) As Variant

    Dim strPath, strProv, strConn As String
    Dim Conn As New Connection
    Dim rsQry As New Recordset

    'If the Access connection is open, close it
    If rsQry.State = adStateOpen Then rsQry.Close

    strPath = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Templates and Logbooks\Invoice and Pricing Databases.accdb"
    strProv = "Microsoft.ACE.OLEDB.12.0;"
    strConn = "Provider=" & strProv & "Data Source=" & strPath

    Conn.Open strConn
    rsQry.Open strQry, Conn
    If rsQry.EOF = False Then Retrieve_Database_Array = rsQry.GetRows(rsQry.RecordCount)

    'close the connection
    If rsQry.State = adStateOpen Then rsQry.Close
    Conn.Close


End Function

Sub Insert_Value_Into_Database(ByVal strQry As String)

    Dim strPath, strProv, strConn As String
    Dim Conn As New Connection
    Dim rsQry As New Recordset

    'If the Access connection is open, close it
    If rsQry.State = adStateOpen Then rsQry.Close

    strPath = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Templates and Logbooks\Invoice and Pricing Databases.accdb"
    strProv = "Microsoft.ACE.OLEDB.12.0;"
    strConn = "Provider=" & strProv & "Data Source=" & strPath

    'Connection open
    Conn.Open strConn

    'Excutes the command
    Conn.Execute strQry

    'close the connection
    If rsQry.State = adStateOpen Then rsQry.Close
    Conn.Close

End Sub
Private Function Find_Testing_Sheet(ByVal Workbook_To_Invoice As Workbook) As Worksheet
    'Loops through each worksheet in a specific workbook to find the testing request.

    Dim WkS As Worksheet
    Dim Tab_Count As Integer

    For Tab_Count = 1 To Workbook_To_Invoice.Worksheets.Count
        If Workbook_To_Invoice.Worksheets(Tab_Count).Range("D3").Value Like "*Testing Request Form" Then
            Set WkS = Workbook_To_Invoice.Sheets(Tab_Count)
            Exit For
        End If
    Next Tab_Count

    Set Find_Testing_Sheet = WkS

End Function

Private Function Find_Customer(ByVal Testing_Request As Worksheet, ByVal Requested_By As String, ByRef Error_Message As String) As String
    'Finds the customer in the Customer Database
    Dim Customer_String, Customer, PONumber, CreditCardNumber As String
    Dim Array_Index As Integer
    Dim Report_Arr As Variant

    'Sets variables
    Customer = ""
    With Testing_Request
        Customer_String = .Range("C10") & " (" & .Range("C11") & " " & .Range("C12") & ")"
    End With

    Customer = Retrieve_Database_Reference("Select [Customer_Name] from Customer_ID WHERE [Name_Address] = '" & Customer_String & "'", "Customer_Name")

    'If the customer is not found, a userform is displayed unless the invoice was requested by Outlook, in which case an error message is formed.
    If Customer = "" Then
        'If the macro was run from Outlook, cancel the macro and log the error so that the next invoice can be processed.
        If Requested_By = "Outlook" Then
            Error_Message = "Customer Not Found"

        Else
            PONumber = Testing_Request.Range("I10")
            CreditCardNumber = Testing_Request.Range("I11")
            Call WaitASecond(0.25)
            With New CustomerNotFoundUserform
                .Label1.Caption = Customer_String & " not found in the database.  Please choose either an existing customer or enter a new one."

                Report_Arr = Retrieve_Database_Array("Select [Customer_Name] from Customer_Database")

                'converts the Report_Array to a 1D array and then sorts it alphabetically
                ReDim New_Arr(UBound(Report_Arr, 2))
                For Array_Index = 0 To UBound(Report_Arr, 2)
                    New_Arr(Array_Index) = Report_Arr(0, Array_Index)
                Next Array_Index
                New_Arr = Alphabetize_Array(New_Arr)

                .CurrentCustomerComboBox.List = New_Arr
                .CurrentCustomerOptionButton = True
                .CustomerName.Text = Testing_Request.Range("C10")
                .PricingIndexComboBox.Text = "Standard Price"
                .PricingIndexComboBox.List = Retrieve_Database_Array("SELECT DISTINCT [Pricing_Key] from Customer_Database")
                .PaymentTerms.Text = "30"
                If Not CreditCardNumber = "" Then .PaymentInformation = "CREDIT CARD #: xxxx-xxxx-xxxx-" & CreditCardNumber
                If Not PONumber = "" Then .PaymentInformation = "P.O. # " & PONumber
                .Address1.Text = Testing_Request.Range("C10")
                .Address2.Text = Testing_Request.Range("C11")
                .Address3.Text = Testing_Request.Range("C12")

                .CC1_1.Text = Testing_Request.Range("C13")
                .CC1_3.Text = Testing_Request.Range("C14")
                .CC1_4.Text = Testing_Request.Range("C15")
                .Show vbModal
                If .IsCancelled = True Then
                    Error_Message = "Customer Not Found"
                    Call ResetSettings
                    Exit Function
                End If
                If .CurrentCustomerOptionButton = True Then
                    Customer = .CurrentCustomerComboBox.Text
                    Call Insert_Value_Into_Database("INSERT INTO Customer_ID ([Customer_Name],[Name_Address]) VALUES ('" & Customer & "','" & Customer_String & "')")
                ElseIf .NewCustomerOptionButton = True Then
                    Customer = .CustomerName.Text
                    Call Insert_Value_Into_Database("INSERT INTO Customer_Database ([Customer_Name],[Pricing_Key],[Payment_Terms],[Payment_Information],[Address1],[Address2],[Address3],[Address4],[Attention_Name],[Attention_Title],[Attention_Email],[Attention_Phone],[CC1_Name],[CC1_Title],[CC1_Email],[CC1_Phone],[CC2_Name],[CC2_Title],[CC2_Email],[CC2_Phone]) VALUES ('" & .CustomerName.Text & "','" & .PricingIndexComboBox.Text & "','" & .PaymentTerms.Text & "','" & .PaymentInformation.Text & "','" & .Address1.Text & "','" & .Address2.Text & "','" & .Address3.Text & "','" & .Address4.Text & "','" & .Attention1.Text & "','" & .Attention2.Text & "','" & .Attention3.Text & "','" & .Attention4.Text & "','" & .CC1_1.Text & "','" & .CC1_2.Text & "','" & .CC1_3.Text & "','" & .CC1_4.Text & "','" & .CC2_1.Text & "','" & .CC2_2.Text & "','" & .CC2_3.Text & "','" & .CC2_4.Text & "')")
                    Call Insert_Value_Into_Database("INSERT INTO Customer_ID ([Customer_Name],[Name_Address]) VALUES ('" & Customer & "','" & Customer_String & "')")
                End If

            End With
        End If
    End If
    'The customer's name is returned to the sub
    Find_Customer = Customer

End Function
Private Function Alphabetize_Array(ByVal Input_Array As Variant) As Variant
    'Alphabetize Sheet Names in Array List https://www.thespreadsheetguru.com/vba/2015/3/24/applying-an-alphabetical-sort-to-your-vba-array-list - accessed 4/24/23

    Dim x_Index, y_Index As Long
    Dim TempTxt1, TempTxt2 As String

    For x_Index = LBound(Input_Array) To UBound(Input_Array)
        For y_Index = x_Index To UBound(Input_Array)
            If UCase(Input_Array(y_Index)) < UCase(Input_Array(x_Index)) Then
                TempTxt1 = Input_Array(x_Index)
                TempTxt2 = Input_Array(y_Index)
                Input_Array(x_Index) = TempTxt2
                Input_Array(y_Index) = TempTxt1
            End If
        Next y_Index
    Next x_Index
    Alphabetize_Array = Input_Array
End Function
Private Function Get_Invoice_Number(ByVal Customer As String, Samples_Received As Date, ByVal RequestedBy As String) As String
    'Returns an invoice number based on how many invoices have been generated for the customer that day
    Dim Invoice_Number As String
    Dim Invoice_Counter As Integer
    Dim Invoice_Counter_Arr As Variant

    'Returns an array (which is then converted to an integer) of the number of invoices with this customer and process date
    Invoice_Counter_Arr = Retrieve_Database_Array("SELECT COUNT(*) FROM Invoice_Database WHERE Customer_Name = '" & Customer & "' AND Invoice_Number like '" & Format(Samples_Received, "yyyymmdd") & "%'")
    Invoice_Counter = CInt(Invoice_Counter_Arr(0, 0))

    Invoice_Number = Format(Samples_Received, "yyyymmdd") & Invoice_Counter + 1
    Get_Invoice_Number = Invoice_Number
End Function

Private Function Get_Logo(ByVal Customer As String, ByVal Requested_By As String) As String
    'Retrieves the file name of the customer's logo. If the logo isn't found, a userform will prompt the user to find one or skip it.

    Dim LogoFolder, LogoFile As String
    Dim Company_Logo_Found As Boolean

    LogoFolder = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Customer Logos\"
    LogoFile = LogoFolder & Customer & " Logo.png"
    'Checks whether the logo file exists
    If Dir(LogoFile) <> "" Then
        Company_Logo_Found = True
        'If the logo file doesn't exist, call a userform to either set the logo file or to skip using it
    Else
        If Requested_By = "Outlook" Then Exit Function
        With FindLogoUserForm
            Call WaitASecond(0.25)
            .Show
            If .Cancelling = True Then
                Exit Function
            ElseIf .BrowseOptionButton = True Or .ChooseExistingOptionButton = True Then
                FileCopy .CopyFile, LogoFolder & Customer & " Logo" & Right(.CopyFile, 4)
                Company_Logo_Found = True
            End If
        End With
    End If

    If Company_Logo_Found = True Then Get_Logo = LogoFile

End Function
Private Function Collect_Analyses(ByVal WkS As Worksheet, ByVal Row_Index As Integer, ByVal Pricing_Index As String, ByVal Pricing_Database As String, ByVal Processing_Column As Integer) As Collection
    'Version 0.1
    'WkS is the testing request used to generate invoice
    'Row_Index is the row on the testing request the function looks at this loop
    'Pricing_index is the string that links the customer database to the pricing database (eg "Customer B Chemical  (various sites)")


    Dim Column_Index, Array_Index, Collection_Index, a_Index, b_Index As Integer
    Dim Translated_Analyses_Collection As New Collection
    Dim Matching_Collection As New Collection
    Dim Database_Array As Variant
    Dim DB_Array_Index As Integer
    Dim Analyses_Array As Variant

    'For Wafers, only the wafer size matters for pricing and there are no pricing packages
    If WkS.Range("D3") = "Wafer Testing Request Form" Then
        Set Translated_Analyses_Collection = Collect_Translated_Names(WkS, Row_Index, 4, 4, 4)
    Else
        Set Translated_Analyses_Collection = Collect_Translated_Names(WkS, Row_Index, 4, Processing_Column - 1, Processing_Column + 3)
        'If additional elements are selected, they are added to the end of the Translated_Analyses_Collection
        If Not WkS.Cells(Row_Index, 5) = "" Then Set Translated_Analyses_Collection = Collect_Additional_Elements(WkS, Row_Index, Translated_Analyses_Collection)
        'Returns an array of the pricing packages for the specified customer
        Database_Array = Retrieve_Database_Array("Select [Analysis] from " & Pricing_Database & " WHERE [Price_Index] = '" & Pricing_Index & "' And [Analysis] like '%/%'")
        'If pricing packages for that customer exist, then test to see if any can be applied.
        If IsEmpty(Database_Array) = False Then
            'loop through every each pricing package to see if the customer has requested all of the analyses
            For DB_Array_Index = 0 To UBound(Database_Array, 2)
                Analyses_Array = Split(Database_Array(0, DB_Array_Index), "/")
                Set Matching_Collection = New Collection
                For a_Index = 0 To UBound(Analyses_Array)
                    For b_Index = 1 To Translated_Analyses_Collection.Count
                        'Adds a "True" to a new collection for each match
                        If Translated_Analyses_Collection(b_Index) = Analyses_Array(a_Index) Then Matching_Collection.Add True
                    Next b_Index
                Next a_Index
                'If the number of items in the matching collection equals the number of items in the pricing key, all items match (remember that arrays start at 0, not 1)
                If Matching_Collection.Count = UBound(Analyses_Array) + 1 Then
                    'add the pricing package as a separate item
                    Translated_Analyses_Collection.Add Database_Array(0, DB_Array_Index)
                    'removes each individual item in the pricing package to avoid duplicates since they're grouped together
                    For Collection_Index = Translated_Analyses_Collection.Count To 1 Step -1
                        For Array_Index = 0 To UBound(Analyses_Array)
                            If Analyses_Array(Array_Index) = Translated_Analyses_Collection(Collection_Index) Then Translated_Analyses_Collection.Remove (Collection_Index)
                        Next Array_Index
                    Next Collection_Index
                End If
            Next DB_Array_Index
        End If
        'Any additional elements that are not included in a pricing "package" are recombined.
        Set Translated_Analyses_Collection = Recombine_Additional_Elements(WkS, Row_Index, Translated_Analyses_Collection)
    End If
    Set Collect_Analyses = Translated_Analyses_Collection
End Function
Private Function Collect_Translated_Names(ByVal WkS As Worksheet, ByVal Row_Index As Integer, ByVal Start_Column As Integer, ByVal End_Column, ByVal Extra_Column As Integer) As Collection
    'Loops through the analyses of a sample row to create an collection of "translated" analyses names
    Dim Column_Index, Array_Index As Integer
    Dim Database_Name_Array() As String
    Dim Translated_Analyses_Collection As New Collection

    For Column_Index = Start_Column To Extra_Column
        'This line makes sure that the processing time columns are missed but the additional requests or notes are checked
        If Column_Index <= End_Column Or Column_Index = Extra_Column Then
            'References the Access database
            Database_Name_Array = Split(Retrieve_Database_Reference("Select [Database_Name] from Analysis_Database WHERE [Testing_Request_Name] = '" & WkS.Cells(Row_Index, Column_Index) & "'", "Database_Name"), "/")
            'if a single cell in the testing request contains mulitple tests (ie 6 Cations + Methylamines) then they are individually added
            For Array_Index = 0 To UBound(Database_Name_Array)
                Translated_Analyses_Collection.Add Database_Name_Array(Array_Index)
            Next Array_Index

        End If
    Next Column_Index
    Set Collect_Translated_Names = Translated_Analyses_Collection
End Function
Private Function Collect_Additional_Elements(ByVal WkS As Worksheet, ByVal Row_Index As Integer, Translated_Analyses_Collection As Collection) As Collection
    'Adds any additional elements to the analysis collection
    Dim Additional_Elements_Array As Variant
    Dim Additional_Elements As String
    Dim Array_Index As Integer
    'removes any spaces in the string
    Additional_Elements = Replace(WkS.Cells(Row_Index, 5), " ", "")
    'Splits the list of additional elements into an array
    Additional_Elements_Array = Split(Additional_Elements, ",")
    ReDim Labelled_Array(UBound(Additional_Elements_Array))
    'Adds each additional element to the analysis collection with a label
    For Array_Index = 0 To UBound(Additional_Elements_Array)
        Translated_Analyses_Collection.Add "Additional Elements " & Additional_Elements_Array(Array_Index)
    Next Array_Index
    'Passes the new collection back
    Set Collect_Additional_Elements = Translated_Analyses_Collection
End Function
Private Function Recombine_Additional_Elements(ByVal WkS As Worksheet, ByVal Row_Index As Integer, Translated_Analyses_Collection As Collection) As Collection
    'Recombines Additional Elements that aren't included in special pricing "packages"
    Dim Collection_Index, Additional_Elements_Count As Integer
    Dim Additional_Elements_Array() As Variant
    'Loops through the collection, recording any additional elements that aren't part of a pricing package and removing them from the collection
    ReDim Additional_Elements_Array(0)
    For Collection_Index = Translated_Analyses_Collection.Count To 1 Step -1
        If Translated_Analyses_Collection(Collection_Index) Like "Additional Elements*" And Len(Translated_Analyses_Collection(Collection_Index)) = Len(Replace(Translated_Analyses_Collection(Collection_Index), "/", "")) Then
            ReDim Preserve Additional_Elements_Array(Additional_Elements_Count)
            Additional_Elements_Array(Additional_Elements_Count) = Right(Translated_Analyses_Collection(Collection_Index), Len(Translated_Analyses_Collection(Collection_Index)) - 20)
            Additional_Elements_Count = Additional_Elements_Count + 1
            Translated_Analyses_Collection.Remove (Collection_Index)
        End If
    Next Collection_Index
    'Any leftover additional elements are added to the end of the collection as one group
    If Not Additional_Elements_Array(0) = "" Then Translated_Analyses_Collection.Add "Additional Elements " & Join(Additional_Elements_Array, ", ")
    Set Recombine_Additional_Elements = Translated_Analyses_Collection
End Function
Private Function Collect_Prices(ByVal Analyses As Collection, ByVal Pricing_Index As String, ByVal Pricing_Database As String, ByVal Processing_Time As String) As Variant
    'Collects the prices for analyses of a sample
    'Analyses is the collection of analyses the function is fetching the price for
    'Pricing_index is the string that links the customer database to the pricing database (eg "Customer B Chemical  (various sites)")
    'Pricing_Database is the sheet in the Invoicing Database use to fetch the price
    'Processing_Time is the column in the database to look in

    Dim Collection_Index, Additional_Element_Count As Integer
    Dim Price_Array() As Variant
    Dim Analyses_in_Prices As New Collection

    ReDim Price_Array(Analyses.Count - 1)

    For Collection_Index = 1 To Analyses.Count
        'calculates how many additional elements have been requested and renames the item so it's price can be retrieved
        If Left(Analyses(Collection_Index), 20) = "Additional Elements " And Len(Analyses(Collection_Index)) - Len(Replace(Analyses(Collection_Index), "/", "")) = 0 Then
            Additional_Element_Count = Len(Analyses(Collection_Index)) - Len(Replace(Analyses(Collection_Index), ",", "")) + 1
            Analyses_in_Prices.Add "Additional Element"
        Else
            'Additional_element_count is assigned "1" if this analysis isn't for additional element so that the price won't change when it is multipled by this count
            Additional_Element_Count = 1
            Analyses_in_Prices.Add Analyses(Collection_Index)
        End If
        'Checks the database for pricing specific to the customer
        Price_Array(Collection_Index - 1) = Retrieve_Database_Reference("Select [" & Processing_Time & "] from " & Pricing_Database & " WHERE [Price_Index] = '" & Pricing_Index & "' And [Analysis] = '" & Analyses_in_Prices(Collection_Index) & "'", Processing_Time)
        'If the customer doesn't have pricing specific to it, the standard price is used
        If Price_Array(Collection_Index - 1) = "" Then Price_Array(Collection_Index - 1) = Retrieve_Database_Reference("Select " & Processing_Time & " from " & Pricing_Database & " WHERE [Price_Index] = 'Standard Price' And [Analysis] = '" & Analyses_in_Prices(Collection_Index) & "'", Processing_Time)
        'If a standard price isn't found, set the price to "0" to prevent errors
        If Price_Array(Collection_Index - 1) = "" Then Price_Array(Collection_Index - 1) = 0
        'the price is formatted (this is done here to avoid errors if the macro can't find a price specific to the customer)
        Price_Array(Collection_Index - 1) = Format(Price_Array(Collection_Index - 1) * Additional_Element_Count, "#,##0.00")
    Next Collection_Index
    Collect_Prices = Price_Array

End Function
Private Function Sum_Prices(ByVal Prices As Variant) As Long
    'Sums the individual prices
    Dim Array_Index As Integer
    Dim Total_Price As Long
    'Loops through each price to get the total
    For Array_Index = 0 To UBound(Prices)
        Total_Price = Total_Price + Prices(Array_Index)
    Next Array_Index

    Sum_Prices = Total_Price
End Function
Private Function Display_Analyses_Names(ByVal Analyses As Collection) As Variant
    'Converts the analysis names to the description to be used in the invoice
    Dim Collection_Index, Array_Index As Integer
    Dim Proper_Names() As Variant
    Dim Analysis_Package As Variant
    Dim Grouped_Analyses() As Variant
    Dim Additional_Elements_Description_Added As Boolean

    ReDim Proper_Names(Analyses.Count - 1)
    'Loops through each analysis and searches for its match in the invoicing database
    For Collection_Index = 1 To Analyses.Count
        Analysis_Package = Split(Analyses(Collection_Index), "/")
        ReDim Grouped_Analyses(UBound(Analysis_Package))
        For Array_Index = 0 To UBound(Analysis_Package)
            If Left(Analysis_Package(Array_Index), 20) = "Additional Elements " And Additional_Elements_Description_Added = False Then Grouped_Analyses(Array_Index) = "Additional Trace Elements (by ICP-MS): " & Right(Analysis_Package(Array_Index), Len(Analysis_Package(Array_Index)) - 20)
            If Left(Analysis_Package(Array_Index), 20) = "Additional Elements " And Additional_Elements_Description_Added = True Then Grouped_Analyses(Array_Index) = Right(Analysis_Package(Array_Index), Len(Analysis_Package(Array_Index)) - 20)
            If Left(Analysis_Package(Array_Index), 20) = "Additional Elements " And Additional_Elements_Description_Added = False Then Additional_Elements_Description_Added = True
            If Not Left(Analysis_Package(Array_Index), 20) = "Additional Elements " Then Grouped_Analyses(Array_Index) = Retrieve_Database_Reference("Select [Invoice_Name] from Analysis_Database WHERE [Database_Name] = '" & Analysis_Package(Array_Index) & "'", "Invoice_Name")
        Next Array_Index
        Proper_Names(Collection_Index - 1) = Join(Grouped_Analyses, " / ")
    Next Collection_Index

    Display_Analyses_Names = Proper_Names
End Function

Public Function Collect_CustomerX_Reports(ByVal Testing_Request_Date As Date, ByVal Folder_Name As String) As Collection
    'Purpose: Creates a collection of all the files that are labelled as between the start and end dates. These files can then be manipulated as desired in the sub.
    'The advantage is that the files don't have to be opened, saving a lot of time and memory.

    'Written by JT, based on solution found at https://stackoverflow.com/questions/67155327/open-excel-files-one-by-one-based-on-filename-with-specific-range-date-time, retreived 9/20/22

    'Version 1.0    - 9/22/22   - Initial release
    'Version 2.0    - 12/28/22  - now generates a collection for the week of the "Testing_Request_Date" with only the latest files
    'Version 2.1    - 4/14/23   - changed to public function so that it can be access from Outlook

    Dim Start_Date, End_Date, Case_Date As Date
    Dim FileName As String
    Dim GetFileNames As New Collection
    Dim aIndex As Integer
    Dim Found_In_Collection As Boolean

    'Sets date variables by finding the Monday and Friday of the week specified to the funcion
    Start_Date = Testing_Request_Date - Weekday(Testing_Request_Date) + 2
    End_Date = Testing_Request_Date - Weekday(Testing_Request_Date) + 6

    'Sets the file name as every xlsx file in the specified folder
    FileName = Dir(Folder_Name & "*.xlsx")

    'Loops through each file in the specified folder
    Do While FileName <> ""
        'sets the date by the file's name in order to compare it to the start and end dates
        Case_Date = DateSerial(Mid(FileName, 5, 2), Left(FileName, 2), Mid(FileName, 3, 2))
        If Case_Date >= Start_Date And Case_Date <= End_Date Then
            Found_In_Collection = False
            'Loops through the collection to see if the Case_Date is already in the collection
            For aIndex = 1 To GetFileNames.Count
                If Left(Folder_Name & FileName, 60) = Left(GetFileNames(aIndex), 60) Then
                    Found_In_Collection = True
                    'if the Case_Date is already in the collection but this file was editted more recently than then other then this is added and the other removed (so only the latest of each date is in the collection)
                    If FileDateTime(Folder_Name & FileName) > FileDateTime(GetFileNames(aIndex)) Then
                        GetFileNames.Add Item:=Folder_Name & FileName, Before:=aIndex
                        GetFileNames.Remove (aIndex + 1)
                        Exit For
                    End If
                End If
            Next aIndex
            'If the Case_Date was not in the collection, this file is added
            If Found_In_Collection = False Then GetFileNames.Add Folder_Name & FileName
        End If
        'Proceed to the next file name
        FileName = Dir()
    Loop

    Set Collect_CustomerX_Reports = GetFileNames
End Function
Private Function Finish_Description(ByVal Invoice_Array As Variant) As Variant
    'Adds the chemical(s) to the beginning of the description row of the array
    Dim Array_Index, Comma_Count, Comma_Index, Comma_Location As Integer

    Dim Short_Chemical_List As String


    'Loops through each instance of the array to add the chemical(s) (ie row 8) to the description (ie row 1)
    For Array_Index = 1 To UBound(Invoice_Array, 2)
        'Sets the list of chemicals for this line
        Short_Chemical_List = Invoice_Array(8, Array_Index)
        'The number of separate chemicals is the number of commas + 1
        Comma_Count = Len(Short_Chemical_List) - Len(Replace(Short_Chemical_List, ",", ""))
        'If there's more than one chemical and the number of characters in the list of chemicals is > 40, the list is cut down and etc is added
        If Comma_Count > 0 And Len(Short_Chemical_List) > 40 Then
            Comma_Location = 1
            'Cuts the description off with the end of the first chemical that takes the description over 40 characters
            For Comma_Index = 1 To Comma_Count
                Comma_Location = InStr(Comma_Location + 1, Short_Chemical_List, ",")
                If Comma_Location > 40 Then Exit For
            Next Comma_Index
            Short_Chemical_List = Left(Short_Chemical_List, Comma_Location - 1) & ", etc"
        End If
        'If there's more than once chemical for this line, add "Various Chemicals" the beginning of the description
        If Comma_Count > 0 Then Short_Chemical_List = "Various Chemicals (" & Short_Chemical_List & ")"
        Invoice_Array(1, Array_Index) = Short_Chemical_List & Invoice_Array(1, Array_Index)
        'Multiplies total cost of each row by the total number of samples
        Invoice_Array(4, Array_Index) = Format(Invoice_Array(0, Array_Index) * Invoice_Array(4, Array_Index), "#,###.00")
    Next Array_Index

    Finish_Description = Invoice_Array
End Function
Private Sub WaitASecond(ByVal Sec As Single)
    'Retrieved from https://stackoverflow.com/questions/54142988/excel-vba-automation-error-the-object-invoked-has-disconnected-from-its-clients 10/11/22 JT
    'Pauses the program to prevent crashing
    Dim WaitTill As Single

    WaitTill = Timer + Sec
    Do
        DoEvents
    Loop While Timer < WaitTill
End Sub
Private Sub Add_Non_CustomerX_Variables(ByVal Word_Document As Word.Document, ByVal Customer As String, Optional ByVal Logo_Location As String)
    'Adds variables to a word invoice that are in a non-CustomerX invoice but not a CustomerX summary invoice

    Dim Company_Logo As Object
    Dim Logo_Width, Logo_Height, CC_1, CC_2 As String

    'Sets the company logo if one has been selected
    If Not Logo_Location = "" Then
        Set Company_Logo = Word_Document.Bookmarks("CompanyLogo").Range.InlineShapes.AddPicture(FileName:=Logo_Location, Linktofile:=False, savewithdocument:=True)
        Logo_Width = GetFileProperty(Logo_Location, "width")
        Logo_Height = GetFileProperty(Logo_Location, "height")
        Company_Logo.Height = 0.75 * (135 / Left(Logo_Width, (Len(Logo_Width) - 7)) * Left(Logo_Height, (Len(Logo_Height) - 7)))
        Company_Logo.Width = 135
    End If
    CC_1 = Customer_Details(Customer, "CC1_Name", "CC1_Title", "CC1_Email", "CC1_Phone")
    CC_2 = Customer_Details(Customer, "CC2_Name", "CC2_Title", "CC2_Email", "CC2_Phone")
    With Word_Document.Tables(1)
        'sets the "Attention" field
        .Cell(1, 2).Range.Text = Customer_Details(Customer, "Attention_Name", "Attention_Title", "Attention_Email", "Attention_Phone")  '"Attention" column
        .Cell(1, 2).Range.Paragraphs(1).Range.Bold = True
        'sets the "CC1" and "CC2" fields (unless neither are required, in which case the labels are hidden).
        If CC_1 = "" And CC_2 = "" Then .Cell(3, 1).Range.Font.TextColor = wdColorWhite
        .Cell(3, 2).Range.Text = CC_1 '"CC1" column
        .Cell(3, 2).Range.Paragraphs(1).Range.Bold = True
        .Cell(3, 3).Range.Text = CC_2 '"CC2" column
        .Cell(3, 3).Range.Paragraphs(1).Range.Bold = True
    End With
End Sub

Private Function Customer_Details(ByVal Customer As String, ByVal Field1 As String, ByVal Field2 As String, ByVal Field3 As String, ByVal Field4 As String) As String
    'Retrieves the details for either Address, Attention, CC1 or CC2
    Dim Customer_Arr As Variant
    Dim Array_Index As Integer
    'References database
    Customer_Arr = Retrieve_Database_Array("Select [" & Field1 & "], [" & Field2 & "], [" & Field3 & "], [" & Field4 & "] FROM Customer_Database WHERE [Customer_Name] = '" & Customer & "'")
    'Removes Null values
    For Array_Index = 0 To UBound(Customer_Arr, 1)
        If IsNull(Customer_Arr(Array_Index, 0)) = False Then Customer_Details = Customer_Details & Customer_Arr(Array_Index, 0) & "///"
    Next Array_Index
    'Removes the last "///" (with the left function) and replaces the remainers with a new line
    If Not Customer_Details = "" Then Customer_Details = Replace(Left(Customer_Details, Len(Customer_Details) - 3), "///", vbNewLine)
End Function

Function Get_Payment_Details(ByVal Testing_Request_Sheet As Worksheet, ByVal Customer As String, ByVal Payment_Column As Integer) As String
    'outputs the payment details - from the testing request if it is there or the database if not

    Dim Invoice_Sheet As Worksheet
    Dim Row_Index, Customer_Row As Integer
    Dim PO_Number, CC_Number, Payment_Terms, Payment_Information As String

    'Sets payment terms and information
    Payment_Terms = Retrieve_Database_Reference("Select [Payment_Terms] from Customer_Database WHERE [Customer_Name] = '" & Customer & "'", "Payment_Terms")
    Payment_Information = Retrieve_Database_Reference("Select [Payment_Information] from Customer_Database WHERE [Customer_Name] = '" & Customer & "'", "Payment_Information")

    'Sets credit card number and PO number from the database
    If Payment_Terms = "Credit Card" Then CC_Number = Payment_Information
    If Not Payment_Terms = "Credit Card" Then PO_Number = Payment_Information

    'If there is no payment information in the database, set the payment information from the testing request
    If Payment_Information = "" Then PO_Number = Testing_Request_Sheet.Cells(10, Payment_Column)
    If Payment_Information = "" Then CC_Number = Testing_Request_Sheet.Cells(11, Payment_Column) & " " & Testing_Request_Sheet.Cells(12, Payment_Column)

    'If the PO and CC numbers don't start with "PO and "Credit Card" respectively, this is added
    If Not PO_Number = "" And Not Left(PO_Number, 1) = "P" Then PO_Number = "P.O. #: " & PO_Number
    If Not CC_Number = "" And IsNumeric(Left(CC_Number, 1)) = True Then CC_Number = "Credit Card: " & CC_Number

    'The payment information is send back to the sub, prioritizing the PO number if both exist
    If Not PO_Number = "" Then Get_Payment_Details = PO_Number
    If PO_Number = "" Then Get_Payment_Details = CC_Number

End Function

Private Function Get_Total(ByVal Invoicing_Matrix As Variant) As String
    'Totals the price for the invoice

    Dim Array_Index, Running_Total As Integer
    'Loops through each sub-total to add up to the total cost
    For Array_Index = 1 To UBound(Invoicing_Matrix, 2)
        If Not Invoicing_Matrix(4, Array_Index) = "" Then Running_Total = Running_Total + Invoicing_Matrix(4, Array_Index)
    Next Array_Index
    'Formats the total cost to insert into the invoice
    Get_Total = Format(Running_Total, "#,###.00") & " USD"
End Function

Private Function GetFileProperty(ByVal sFile As String, _
                      ByVal sPropertyName As String) As String
    ' retrieved from https://www.devhut.net/how-to-retrieve-a-files-properties-with-vba/ on 7/28/22 by Daniel Pineault, JT

    'On Error GoTo error_handler
    Dim oShell                As Object    'Shell
    Dim oFolder               As Object    'Folder
    Dim oFolderItem           As Object    'FolderItem
    Dim sFilePath             As String
    Dim sFileName             As String
    Dim i                     As Long
    Dim lPropertyNumber       As Long
    Dim vPropValue            As Variant

    sFilePath = Left(sFile, InStrRev(sFile, "\") - 1)
    sFileName = Right(sFile, Len(sFile) - InStrRev(sFile, "\"))

    Set oShell = CreateObject("Shell.Application")
    Set oFolder = oShell.Namespace(CStr(sFilePath))

    If (Not oFolder Is Nothing) Then
        Set oFolderItem = oFolder.ParseName(sFileName)
        For i = 0 To 320 'This could be bumped up in case MS increase the number again
            If oFolder.GetDetailsOf(oFolder.Items, i) = sPropertyName Then
                lPropertyNumber = i
                Exit For
            End If
        Next

        If lPropertyNumber = 0 Then
            'Property not found
            GoTo Error_Handler_Exit
        Else
            vPropValue = oFolder.GetDetailsOf(oFolderItem, lPropertyNumber)
            If Trim(vPropValue & vbNullString) <> "" Then
                vPropValue = Replace(Replace(Replace(Replace(vPropValue, ChrW(8236), ""), ChrW(8234), ""), ChrW(8207), ""), ChrW(8206), "")
            End If
            GetFileProperty = vPropValue
        End If
    End If

Error_Handler_Exit:
    On Error Resume Next
    If Not oFolderItem Is Nothing Then Set oFolderItem = Nothing
    If Not oFolder Is Nothing Then Set oFolder = Nothing
    If Not oShell Is Nothing Then Set oShell = Nothing
    Exit Function
End Function

Private Sub Fill_Out_Logbook(ByVal Customer As String, ByVal Invoice_Number As String, ByVal Report_Name As String, ByVal Total_Cost As String, ByVal Generated_From As String, Optional ByVal Error_Message As String, Optional ByVal Process_Date As Date, Optional ByVal Invoice_Week As Integer, Optional ByVal Suggested_Name As String)
    'Fills out the Invoice Logbook troubleshoots errors

    'Version 2.0    - 4/18/23   - Updated to use Access instead of Excel-based database

    Dim Save_Folder As String

    Save_Folder = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Saved Invoices\" & Format(Process_Date, "mm-dd-yyyy") & "\"
    If Total_Cost = "" Then Total_Cost = "0.00 USD"
    If Invoice_Number = "" Then Invoice_Number = 0
    If Customer Like "CustomerX*" Then Report_Name = Customer & " Summary Invoice Week " & Invoice_Week
    Call Insert_Value_Into_Database("INSERT INTO Invoice_Database (Invoice_Date, Invoice_Time, Invoice_Number, Customer_Name, Report_Invoiced, Amount_Charged, Generated_From, Error_Message, Invoice_File_Path,Invoice_File_Name) VALUES ('" & Format(Now, "m/d/yyyy") & "', '" & Format(Now, "HH:MM Am/Pm") & "', '" & Invoice_Number & "', '" & Customer & "', '" & Report_Name & "', '" & Left(Total_Cost, Len(Total_Cost) - 4) & "', '" & Generated_From & "', '" & Error_Message & "', '" & Save_Folder & "', '" & Suggested_Name & "')")

End Sub
Sub Email_Invoices()
    'Emails invoices to the customer

    Dim a_Index As Integer
    Dim Short_File_Name, Full_File_Name, Customer, Recepient_Email, CC_Email, Subject_Line, Invoice_Type As String
    Dim Invoice_DB_Array As Variant

    
    
    With Email_Invoices_Userform
        .Show
        If .Cancelling = True Then
            Unload Email_Invoices_Userform
            Exit Sub
        End If
        For a_Index = 0 To .Invoice_List_Box.ListCount - 1
            Short_File_Name = .Invoice_List_Box.List(a_Index)
            Full_File_Name = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Saved Invoices\" & .Date_Combo_Box.Value & "\" & Short_File_Name & ".docx"
            Customer = Retrieve_Database_Array("SELECT [Customer_Name] from Invoice_Database WHERE [Invoice_File_Name] = '" & Short_File_Name & "'")(0, 0)
            Invoice_DB_Array = Retrieve_Database_Array("SELECT [Attention_Name], [Attention_Email], [CC1_Name], [CC1_Email], [CC2_Name], [CC2_Email] from Customer_Database WHERE Customer_Name = '" & Customer & "'")
            Recepient_Email = Invoice_DB_Array(1, 0)
            If IsNull(Invoice_DB_Array(3, 0)) = False Then CC_Email = Invoice_DB_Array(3, 0)
            If IsNull(Invoice_DB_Array(5, 0)) = False Then CC_Email = CC_Email & "; " & Invoice_DB_Array(5, 0)
            If Left(Short_File_Name, 24) = "PRECILAB Invoice Receipt" Then Subject_Line = Format(Now, "mm-dd-yyyy") & " PRECILAB Invoice Receipt to " & Customer
            If Not Left(Short_File_Name, 24) = "PRECILAB Invoice Receipt" Then Subject_Line = Format(Now, "mm-dd-yyyy") & " PRECILAB Invoice to " & Customer
            Call Send_Email("jamie.thomson@precilab.com", Subject_Line, Full_File_Name, Short_File_Name, , Recepient_Email & "; " & CC_Email)
        Next a_Index
    End With

    Unload Email_Invoices_Userform

End Sub
Private Sub Send_Email(ByVal Recepient_Email As String, ByVal Email_Subject As String, ByVal Attachment_FullName As String, ByVal Attachment_ShortName, Optional ByVal CC_Receipient As String, Optional ByVal Extra_Information As String)

    Dim EmailApp As Outlook.Application
    Dim EmailItem As Outlook.MailItem
    Dim wrdApp As Word.Application
    Dim WrdDoc As Word.Document
    Dim Save_Folder, Signature As String

    'Saves the invoice as a pdf (has to open word to do this
    Set wrdApp = CreateObject("Word.Application")
    Set WrdDoc = wrdApp.Documents.Open(Attachment_FullName)
    Save_Folder = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Sent Invoices (test)\"
    WrdDoc.ExportAsFixedFormat OutputFileName:=Save_Folder & Attachment_ShortName & ".pdf", ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, IncludeDocProps:=True, CreateBookmarks:=wdExportCreateWordBookmarks, BitmapMissingFonts:=True
    WrdDoc.Close savechanges:=False
    wrdApp.Quit

    Set EmailApp = New Outlook.Application
    Set EmailItem = EmailApp.CreateItem(olMailItem)

    With EmailItem
        .To = Recepient_Email
        .CC = CC_Receipient
        .Subject = Email_Subject
        'displays the email so that the default signiture can be saved (due to a vba bug)
        .Display
        Signature = .HTMLBody
        'Adds the body of the email
        .HTMLBody = "Dear PRECILAB Customer,<br><br>Please find attached " & Attachment_ShortName & "."
        .HTMLBody = .HTMLBody & "<br><br>Thank you for your business."
        .HTMLBody = .HTMLBody & "<br><br>NOTE: this would have gone to " & Extra_Information & "."
        'Adds the signature back to the bottom of the email text
        .HTMLBody = .HTMLBody & "<br>" & Signature
        'Adds the pdf version of the invoice
        .Attachments.Add Save_Folder & Attachment_ShortName & ".pdf"
        .Display
        '.Send 'uncomment this line to automate sending
    End With

End Sub

Sub Invoice_Customers_Not_Found(ByVal Is_Invoice_WorkBook_Open As Boolean)
    'This macro will open the userform to select the file to invoice and then call the InvoiceMacro. This is separate so that the InvoiceMacro can be called from Outlook.

    'Version 1.0    -   11/30/22    - Initial release
    'Version 1.1    -   1/18/23     - made the invoice save in the folder for the date of the error
    'Version 1.11   -   1/19/23     - Cleaned up formatting and descriptions
    'Version 1.12   -   1/20/23     - Bug fix for logging information
    'Version 2.0    -   1/20/23     - Moved to Invoice Macro Workbook

    Dim wrdApp As Word.Application
    Dim WrdDoc As Word.Document
    Dim ErrMsg As String
    Dim Error_List() As Variant
    Dim Error_Row_List() As Variant
    Dim Row_Index, Array_Index, Error_Count As Integer

    ReDim Error_List(0)
    ReDim Error_Row_List(0)
    Error_Count = 0
    'Counts the errors and stores their location
    With Workbooks("Invoice and Error Logbook.xlsm").Sheets("Invoice Log")
        For Row_Index = .Cells(Rows.Count, 1).End(xlUp).Row To 3 Step -1
            If Not .Cells(Row_Index, 7) = "" And Not .Cells(Row_Index, 7) = "Fixed and logged below" And Not .Cells(Row_Index, 7) = "No Error" And Not .Cells(Row_Index, 4) = "" Then
                ReDim Preserve Error_List(Error_Count)
                ReDim Preserve Error_Row_List(Error_Count)
                Error_List(Error_Count) = .Cells(Row_Index, 4)
                Error_Row_List(Error_Count) = Row_Index
                Error_Count = Error_Count + 1
            End If
        Next Row_Index
    End With
    'if there are no errors, cancel the macro
    If Error_Count = 0 Then
        MsgBox "No errors found.", , "Invoicing Macro"
        If Is_Invoice_WorkBook_Open = False Then Workbooks("Invoicing Macro Workbook.xlsm").Close savechanges:=False
        Exit Sub
    End If
    'if there are errors, show the userform
    With Invoice_Errors_Userform
        .Invoice_One_Combobox.List = Error_List
        .Show
        Set wrdApp = CreateObject("Word.Application")
        'if the cancel button is clicked, cancel the macro
        If .Cancelling = True Then
            wrdApp.Quit
            If Is_Invoice_WorkBook_Open = False Then Workbooks("Invoicing Macro Workbook.xlsm").Close savechanges:=False
            Exit Sub
            'loop through the error list array until the selected error is found and create an invoice from that file
        ElseIf .Invoice_One_Button = True Then
            For Array_Index = UBound(Error_Row_List) To 0 Step -1
                If .Invoice_One_Combobox.Text = Error_List(Array_Index) Then
                    Set WrdDoc = Generate_Invoice(wrdApp, Error_Row_List(Array_Index), ErrMsg)
                    Exit For
                End If
            Next Array_Index
            'loop through each error and create each invoice
        ElseIf .Invoice_All_Button = True Then
            For Array_Index = 0 To UBound(Error_Row_List)
                Set WrdDoc = Generate_Invoice(wrdApp, Error_Row_List(Array_Index), ErrMsg)
            Next Array_Index
        End If
    End With
    Unload Invoice_Errors_Userform
    'open the invoice
    If Not WrdDoc Is Nothing Then
        wrdApp.Visible = True
        WrdDoc.Application.Activate
    Else
        wrdApp.Quit
    End If
End Sub
Private Function Generate_Invoice(wrdApp As Word.Application, ByVal Error_Row As Integer, ByRef ErrMsg As String) As Word.Document
    'opens a workbook from the logfile and creates an invoice
    Dim WrkBk As Workbook
    Dim WrdDoc As Word.Document
    Dim sSaveFolder As String
    Dim Invoice_Sheet As Worksheet
    Dim Last_Row As Integer
    Dim fdObj As Object

    'sets variables
    Set fdObj = CreateObject("Scripting.FileSystemObject")
    Set Invoice_Sheet = Workbooks("Invoice and Error Logbook.xlsm").Sheets("Invoice Log")
    sSaveFolder = "\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Saved Invoices\" & Format(Invoice_Sheet.Cells(Error_Row, 1), "mm-dd-yyyy") & "\"
    'Opens the workbook that caused the error
    Set WrkBk = Workbooks.Open(Invoice_Sheet.Cells(Error_Row, 8) & "\" & Invoice_Sheet.Cells(Error_Row, 4))
    'Creates an invoice through the Invoicing Macro Workbook
    Set WrdDoc = Create_Word_Invoice(WrkBk, wrdApp, ErrMsg, "Log Book", Invoice_Sheet.Cells(Error_Row, 1))
    'Fills in the Logbook and saves
    Workbooks.Open ("\\PRECILAB-SERVER\LabPlusServer\Macros\Invoicing Macro\Templates and Logbooks\Backups\Invoice Log - Master File.xlsx"), notify:=False
    With Workbooks("Invoice Log - Master File.xlsx").Sheets("Invoice Log")
        .Cells(Error_Row, 7) = "Fixed and logged below"
        Last_Row = .Cells(Rows.Count, 1).End(xlUp).Row
        .Cells(Last_Row, 9).Hyperlinks.Add Invoice_Sheet.Cells(Last_Row, 9), sSaveFolder, , "The folder the invoice is stored in.", sSaveFolder
        .Range(.Cells(3, 1), .Cells(Last_Row, 9)).Copy Destination:=Invoice_Sheet.Range("A3")
    End With
    Application.CutCopyMode = False
    Workbooks("Invoice Log - Master File.xlsx").Close savechanges:=True
    'If the folder we want to save the invoice in doesn't exist, it is created
    If Not fdObj.folderexists(sSaveFolder) Then MkDir (sSaveFolder)
    'If no errors occured (ie the invoice exists), the invoice is saved.
    If Not WrdDoc Is Nothing Then WrdDoc.SaveAs2 sSaveFolder & WrdDoc.ActiveWindow.Caption & ".docx"
    'The invoice is set to the function.
    Set Generate_Invoice = WrdDoc
End Function
