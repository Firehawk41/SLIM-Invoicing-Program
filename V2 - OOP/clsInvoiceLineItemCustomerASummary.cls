VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceLineItemCustomerASummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IInvoiceLineItem
' clsInvoiceLineItemCustomerA
' MacroName: clsInvoiceLineItemCustomerA
' Version: 1.2.0
' Author: JT
' Created: ?
' LastModified: 2025-10-27
' Description: Specialized Line Item object for Customer A invoices
' DependsOn: clsLoggingSystem, clsInvoiceLineItemBase, clsInvoiceLineItemFormattedText, clsAccessDatabase, clsInvoicePriceRetriever, clsSLIMDictionaryManager, clsTestingRequestSubmission, clsTestingRequestSample, clsTestingRequestAnalysis
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
'   - 1.1.0 - 2025-07-22 - Adjusted GetUnitPrices to work for the new clsInvoicePriceDetail
'   - 1.2.0 - 2025-10-27 - Fixed bug where additional elements weren't labelled or priced correctly by using helper methods from clsInvoiceLineItemBase
Private m_Logger As clsLoggingSystem
Private m_LineItemBase As clsInvoiceLineItemBase
Private m_ColumnDictionary As Scripting.Dictionary
Private m_TextColour As Long
Private m_RequestNumber() As clsInvoiceLineItemFormattedText
Private m_Quantity() As clsInvoiceLineItemFormattedText
Private m_Description() As clsInvoiceLineItemFormattedText
Private m_UnitPrices() As clsInvoiceLineItemFormattedText
Private m_ExtendedCost() As clsInvoiceLineItemFormattedText
Public Property Get IInvoiceLineItem_ColumnDictionary() As Scripting.Dictionary
    Set IInvoiceLineItem_ColumnDictionary = SetColumnDictionary()
End Property
Public Property Get IInvoiceLineItem_BoldDictionary() As Scripting.Dictionary
    Set IInvoiceLineItem_BoldDictionary = GetBoldData()
End Property
Public Property Get IInvoiceLineItem_TextColour() As Long
    IInvoiceLineItem_TextColour = m_TextColour
End Property
Public Property Get IInvoiceLineItem_TotalCost() As Long
    IInvoiceLineItem_TotalCost = GetTotalCost()
End Property
Public Sub IInvoiceLineItem_Initialize(GroupedSamples As Collection, TextColour As Long, Logger As clsLoggingSystem, AccessDB As clsAccessDatabase, PriceRetriever As clsInvoicePriceRetriever, SLIMDict As clsSLIMDictionaryManager)
    Dim Submission As clsTestingRequestSubmission
    Const PRECILAB_BLUE As Long = 16711680

    Set m_Logger = Logger
    Set Submission = GroupedSamples(1)("Submission")
    
    Call m_LineItemBase.Initialize(GroupedSamples, PriceRetriever, SLIMDict, m_Logger)
    
    m_RequestNumber = GetRequestNumber(Submission.POInformation)
    m_Quantity = GetQuantity() ' Customer A samples are never grouped and so the quantity is always '1'
    m_Description = GetDescription(GroupedSamples)
    m_UnitPrices = GetUnitPrices(GroupedSamples, PriceRetriever)
    m_ExtendedCost = m_UnitPrices
    
    If IsNumeric(Submission.POInformation) Then ' Check that there is a numeric po number to prevent errors due to mod calculation
        m_TextColour = IIf(Submission.POInformation Mod 2 = 0, vbBlack, PRECILAB_BLUE)
    End If
    
    Set m_ColumnDictionary = SetColumnDictionary()
        
End Sub
Private Function GetRequestNumber(POInformation As String) As clsInvoiceLineItemFormattedText()
    Dim FormattedText(0) As New clsInvoiceLineItemFormattedText
    
    Call FormattedText(0).Initialize(POInformation, m_Logger)
    
    ' Return the formatted text to the calling sub
    GetRequestNumber = FormattedText
End Function
Private Function GetQuantity() As clsInvoiceLineItemFormattedText()
    Dim FormattedText(0) As New clsInvoiceLineItemFormattedText
    
    Call FormattedText(0).Initialize(CStr(1), m_Logger)
    
    ' Return the formatted text to the calling sub
    GetQuantity = FormattedText
End Function
Private Function GetDescription(GroupedSamples As Collection) As clsInvoiceLineItemFormattedText()
    Dim Submission As clsTestingRequestSubmission
    Dim Sample As clsTestingRequestSample
    Dim Analysis As clsTestingRequestAnalysis
    Dim SampleString As String
    Dim ProcessingTime() As clsInvoiceLineItemFormattedText
    Dim ProcessingTimeStr As String
    Dim Description(0) As clsInvoiceLineItemFormattedText
    Dim AnalysisDescription As String
    
    ' Explicitly set the variables
    Set Sample = GroupedSamples(1)("Sample")
    Set Submission = GroupedSamples(1)("Submission")
    Set Analysis = GroupedSamples(1)("Analysis")
    
    ' Build the description
    SampleString = format(Submission.DateReceived, "mmddyy") & "-" & Sample.Chemical.ChemicalName & "-" & Submission.Customer.CustomerName & "-" & Sample.SampleName
    If Analysis.AnalysisDescription = "Additional Element" Then
        AnalysisDescription = m_LineItemBase.ReplaceAdditionalElements(Analysis.AnalysisDescription, Sample)
    Else
        AnalysisDescription = Analysis.AnalysisDescription
    End If
    ProcessingTime = m_LineItemBase.ProcessingTime
    ProcessingTimeStr = ProcessingTime(0).Text
    Set Description(0) = New clsInvoiceLineItemFormattedText
    
    Call Description(0).Initialize(SampleString & vbNewLine & AnalysisDescription & IIf(ProcessingTimeStr = "Next Day", "", " " & ProcessingTimeStr), m_Logger)
    ' Set the colour dictionary in the formatted text object if required
    If ProcessingTimeStr <> "Next Day" Then
        Call Description(0).AddColourDictEntry(vbRed, Len(Description(0).Text) - Len(ProcessingTimeStr))
    End If
    
    ' Return the description to the calling sub
    GetDescription = Description
    
End Function
Private Function GetUnitPrices(GroupedSamples As Collection, PriceRetriever As clsInvoicePriceRetriever) As clsInvoiceLineItemFormattedText()
    Dim PricingDict As Scripting.Dictionary
    Dim Submission As clsTestingRequestSubmission
    Dim Sample As clsTestingRequestSample
    Dim Analysis As clsTestingRequestAnalysis
    Dim Key As Variant
    Dim UnitPrices(0) As clsInvoiceLineItemFormattedText
    Dim InvoicePriceDetail As clsInvoicePriceDetail
    Dim Multiplier As Double

    
    Set Submission = GroupedSamples(1)("Submission")
    Set Sample = GroupedSamples(1)("Sample")
    Set Analysis = GroupedSamples(1)("Analysis")
    
    Set PricingDict = PriceRetriever.RetrievePrices(Submission, Sample, 1)
    
    
    For Each Key In PricingDict.Keys
        If Key = Analysis.AnalysisName Then
            Set UnitPrices(0) = New clsInvoiceLineItemFormattedText
            Set InvoicePriceDetail = PricingDict(Key)
            If Analysis.AnalysisDescription = "Additional Element" Then
                Multiplier = m_LineItemBase.MultiplierForAdditionalElements(Sample)
            Else
                Multiplier = 1
            End If
            Call UnitPrices(0).Initialize(format(InvoicePriceDetail.Price * Multiplier, "#,###.00"), m_Logger)
            ' Add red colour for anything above Next day
            If Sample.ProcessingTime > 2 Then
                Call UnitPrices(0).AddColourDictEntry(vbRed)
            End If
            
            Exit For
        End If
    Next Key
    
    ' Return the unit prices to the calling sub
    GetUnitPrices = UnitPrices
    
End Function
Private Function SetColumnDictionary() As Scripting.Dictionary
    Dim ColumnDictionary As New Scripting.Dictionary
    
    ColumnDictionary.Add 1, m_RequestNumber
    ColumnDictionary.Add 2, m_Quantity
    ColumnDictionary.Add 3, m_Description
    ColumnDictionary.Add 4, m_UnitPrices
    ColumnDictionary.Add 5, m_ExtendedCost
    
    ' Return the dictionary to the calling sub
    Set SetColumnDictionary = ColumnDictionary

End Function
Private Function GetBoldData() As Scripting.Dictionary
    Dim BoldDictionary As New Scripting.Dictionary
'    Dim iIndex As Long
'    Dim BoldArr() As Long
'
'    ' Ensure there is more than one row before resizing
'    If UBound(m_Description) > 0 Then
'        ' Create array for non-bold row indices (skip first row)
'        ReDim BoldArr(UBound(m_Description) - 1)
'
'        ' Populate array with row indices (start at 1, skipping first row)
'        For iIndex = 1 To UBound(m_Description)
'            BoldArr(iIndex - 1) = iIndex
'        Next iIndex
'
'        ' Add to dictionary for column 3 (description column)
'        BoldDictionary.Add 3, BoldArr
'    End If
    
    ' Return the dictionary
    Set GetBoldData = BoldDictionary
End Function
Private Function GetTotalCost() As Long
    Dim iIndex As Long
    Dim RunningTotal As Long
    
    For iIndex = 0 To UBound(m_ExtendedCost)
        If Not m_ExtendedCost(iIndex).Text = "" Then
            RunningTotal = RunningTotal + CLng(m_ExtendedCost(iIndex).Text)
        End If
    Next iIndex
    
    ' Return the cost to the calling sub
    GetTotalCost = RunningTotal
    
End Function
Public Sub IInvoiceLineItem_FillOutMinCharge(MinimumCharge As Long, TotalCost As Long, TextColour As Long, Logger As clsLoggingSystem)
    ' Place holder
End Sub
Private Sub Class_Initialize()
    Set m_LineItemBase = New clsInvoiceLineItemBase
End Sub

Private Sub Class_Terminate()
    Set m_LineItemBase = Nothing
End Sub
