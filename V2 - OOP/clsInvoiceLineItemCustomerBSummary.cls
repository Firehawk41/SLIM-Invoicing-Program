VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceLineItemCustomerBSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IInvoiceLineItem
' MacroName: clsInvoiceLineItemCustomerB
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-11
' Description: Specialized Line Item object for Customer B invoices
' DependsOn: clsLoggingSystem, clsInvoiceLineItemBase, clsInvoiceLineItemFormattedText, clsInvoiceContact, clsAccessDatabase, clsInvoicePriceRetriever, clsSLIMDictionaryManager, clsTestingRequestSubmission, clsTestingRequestSample
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Private m_Logger As clsLoggingSystem
Private m_LineItemBase As clsInvoiceLineItemBase
Private m_Quantity() As clsInvoiceLineItemFormattedText
Private m_LogInDate() As clsInvoiceLineItemFormattedText
Private m_Description() As clsInvoiceLineItemFormattedText
Private m_Contact As clsInvoiceContact
Private m_FormattedContact() As clsInvoiceLineItemFormattedText
Private m_ProcessingTime() As clsInvoiceLineItemFormattedText
Private m_UnitPrices() As clsInvoiceLineItemFormattedText
Private m_ExtendedCost() As clsInvoiceLineItemFormattedText
Private m_TextColour As Long

Public Property Get IInvoiceLineItem_ColumnDictionary() As Scripting.Dictionary
    Set IInvoiceLineItem_ColumnDictionary = GetColumnData()
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
    
    Set m_Logger = Logger
    Set Submission = GroupedSamples(1)("Submission")
    
    Call m_LineItemBase.Initialize(GroupedSamples, PriceRetriever, SLIMDict, m_Logger)
    
    m_Quantity = GetQuantity(GroupedSamples.Count)
    m_LogInDate = GetLogInDate(Submission.DateReceived)
    m_Description = m_LineItemBase.Description(CLng(m_Quantity(0).Text))
    m_FormattedContact = GetFormattedContact(Submission)

    m_ProcessingTime = m_LineItemBase.ProcessingTime
    m_UnitPrices = m_LineItemBase.UnitPrices(CLng(m_Quantity(0).Text))
    m_ExtendedCost = m_LineItemBase.ExtendedPrice(CLng(m_Quantity(0).Text))
    m_TextColour = TextColour
  
    
End Sub
Private Function GetQuantity(Quantity As Long) As clsInvoiceLineItemFormattedText()
    Dim FormattedText(0) As New clsInvoiceLineItemFormattedText
    
    Call FormattedText(0).Initialize(CStr(Quantity), m_Logger)
    
    ' Return the formatted text to the calling sub
    GetQuantity = FormattedText
End Function
Private Function GetLogInDate(LogInDate As Date) As clsInvoiceLineItemFormattedText()
    Dim FormattedText(0) As New clsInvoiceLineItemFormattedText
    
    Call FormattedText(0).Initialize(format(LogInDate, "mm-dd-yyyy"), m_Logger)
    
    ' Return the formatted text to the calling sub
    GetLogInDate = FormattedText
End Function
Private Function GetFormattedContact(Submission As clsTestingRequestSubmission) As clsInvoiceLineItemFormattedText()
    Dim FormattedText(0) As New clsInvoiceLineItemFormattedText
    Dim FormattedContact As String
    
    Call m_Contact.InitializeFromSubmission(Submission, m_Logger)
    
    Call FormattedText(0).Initialize(m_Contact.FormattedContact, m_Logger)
    
    ' Return the formatted text to the calling sub
    GetFormattedContact = FormattedText
End Function

Private Function GetColumnData() As Scripting.Dictionary
    Dim ColumnDictionary As New Scripting.Dictionary
    
    ColumnDictionary.Add 1, m_LogInDate
    ColumnDictionary.Add 2, m_Quantity
    ColumnDictionary.Add 3, m_Description
    ColumnDictionary.Add 4, m_FormattedContact
    ColumnDictionary.Add 5, m_ProcessingTime
    ColumnDictionary.Add 6, m_UnitPrices
    ColumnDictionary.Add 7, m_ExtendedCost
    
    ' Return the dictionary to the calling sub
    Set GetColumnData = ColumnDictionary
End Function
Private Function GetBoldData() As Scripting.Dictionary
    Dim BoldDictionary As New Scripting.Dictionary
    Dim iIndex As Long
    Dim BoldArr() As Long
    
    ' Ensure there is more than one row before resizing
    If UBound(m_Description) > 0 Then
        ' Create array for non-bold row indices (skip first row)
        ReDim BoldArr(UBound(m_Description) - 1)
        
        ' Populate array with row indices (start at 1, skipping first row)
        For iIndex = 1 To UBound(m_Description)
            BoldArr(iIndex - 1) = iIndex
        Next iIndex
        
        ' Add to dictionary for column 3 (description column)
        BoldDictionary.Add 3, BoldArr
    End If
    
    ' Return the dictionary
    Set GetBoldData = BoldDictionary
End Function
Private Function GetTotalCost() As Long
    Dim iIndex As Long
    Dim RunningTotal As Long
    
    For iIndex = 0 To UBound(m_ExtendedCost)
        If Not m_ExtendedCost(iIndex) Is Nothing Then
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
    Set m_Contact = New clsInvoiceContact
End Sub

Private Sub Class_Terminate()
    Set m_LineItemBase = Nothing
    Set m_Contact = Nothing
End Sub
