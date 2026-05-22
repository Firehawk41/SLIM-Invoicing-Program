VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceLineItemRegIndividual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceLineItemRegIndividual
' Version: 1.1.1
' Author: JT
' Created: ?
' LastModified: 2025-09-30
' Description: Speciality line item for regular invoices
' DependsOn: clsLoggingSystem, clsInvoiceLineItemBase, clsInvoiceLineItemFormattedText, IInvoiceLineItem, clsTestingRequestSample, clsAccessDatabase, clsInvoicePriceRetriever, clsSLIMDictionaryManager
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
'   - 1.1.1 - 2025-09-30 - Modified Initialize so that unit prices cell would be red if a rush
Implements IInvoiceLineItem
Private m_Logger As clsLoggingSystem
Private m_LineItemBase As clsInvoiceLineItemBase
Private m_Quantity() As clsInvoiceLineItemFormattedText
Private m_Description() As clsInvoiceLineItemFormattedText
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
    Set m_Logger = Logger

    Call m_LineItemBase.Initialize(GroupedSamples, PriceRetriever, SLIMDict, m_Logger)
    
    m_Quantity = GetQuantity(GroupedSamples.Count)
    m_Description = m_LineItemBase.Description(CLng(m_Quantity(0).Text))
    m_ProcessingTime = m_LineItemBase.ProcessingTime
    m_UnitPrices = m_LineItemBase.UnitPrices(CLng(m_Quantity(0).Text))

    ' Sets the unit price cell to red if the processing time is a rush
    Call SetUnitTimeColour
        
    m_ExtendedCost = m_LineItemBase.ExtendedPrice(CLng(m_Quantity(0).Text))
    m_TextColour = TextColour
        
End Sub
Private Function GetQuantity(Quantity As Long) As clsInvoiceLineItemFormattedText()
    Dim FormattedText(0) As New clsInvoiceLineItemFormattedText
    
    Call FormattedText(0).Initialize(CStr(Quantity), m_Logger)
    
    ' Return the formatted text to the calling sub
    GetQuantity = FormattedText
End Function
Private Sub SetUnitTimeColour()
    Const SAME_DAY_RUSH As String = "Same Day RUSH"
    Const CALL_IN_RUSH As String = "Call-in RUSH"
    Dim iIndex As Long
    
    ' Sets the unit price cell to red if the processing time is a rush
    If m_ProcessingTime(0).Text = SAME_DAY_RUSH Or m_ProcessingTime(0).Text = CALL_IN_RUSH Then
        For iIndex = LBound(m_UnitPrices) To UBound(m_UnitPrices)
            Call m_UnitPrices(iIndex).AddColourDictEntry(vbRed)
        Next iIndex
    End If
    
End Sub
Private Function GetColumnData() As Scripting.Dictionary
    Dim ColumnDictionary As New Scripting.Dictionary
    
    ' Create dictionary with column number as key and contents of that column as the paired value
    ColumnDictionary.Add 1, m_Quantity
    ColumnDictionary.Add 2, m_Description
    ColumnDictionary.Add 3, m_ProcessingTime
    ColumnDictionary.Add 4, m_UnitPrices
    ColumnDictionary.Add 5, m_ExtendedCost
    
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
        
        ' Add to dictionary for column 2 (description column)
        BoldDictionary.Add 2, BoldArr
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
    Set m_Logger = Logger
    
    m_Quantity = PopulateFormattedTextObj("1", True)
    m_Description = PopulateFormattedTextObj("Minimum charge adjustment", True)
    m_ProcessingTime = PopulateFormattedTextObj("N.A.", True)
    m_UnitPrices = PopulateFormattedTextObj(CStr(MinimumCharge - TotalCost & ".00"), True)
    m_ExtendedCost = PopulateFormattedTextObj(CStr(MinimumCharge - TotalCost & ".00"), True)
    If Not TextColour = vbBlack Then
        m_TextColour = vbBlack
    Else
        m_TextColour = TextColour
    End If
End Sub
Private Function PopulateFormattedTextObj(Value As String, Optional IsBold As Boolean = False) As clsInvoiceLineItemFormattedText()
    Dim TextObj(0) As New clsInvoiceLineItemFormattedText
    
    ' initialized the text object
    Call TextObj(0).Initialize(Value, m_Logger, IsBold)
    
    ' Return the object to calling sub
    PopulateFormattedTextObj = TextObj
End Function
Private Sub Class_Initialize()
    Set m_LineItemBase = New clsInvoiceLineItemBase
End Sub

Private Sub Class_Terminate()
    Set m_LineItemBase = Nothing
End Sub

