VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceFactory
' Version: 1.0.3
' Author: JT
' Created: ?
' LastModified: 2025-10-27
' Description: Factory class for Invoices
' DependsOn: clsLoggingSystem, modEnums, IInvoiceSubmissionManager, IInvoice, clsAccessDatabase, clsTestingRequestSubmission, clsInvoicePriceRetriever, clsSLIMDictionaryManager, IInvoiceLineItem, IInvoiceSampleRow, clsTestingRequestSample, clsTestingRequestAnalysis
' ChangeLog:
'   - 1.0.1 - 2025-06-05 - Added metadata
'   - 1.0.2 - 2025-07-23 - Added 'LineCount' to GroupByAnalysis to fix bug where a request with 2 samples with the same id and analyses didn't invoice the second sample
'   - 1.0.3 - 2025-10-27 - Switched 'OrderedSampleDict' to a collection to fix bug where multiple instances of the same sample ID where skipped
Private m_Logger As clsLoggingSystem
Private m_InvoiceType As InvoiceTypeEnum
Public Sub Initialize(InvoiceType As InvoiceTypeEnum, Logger As clsLoggingSystem)
    m_InvoiceType = InvoiceType
    Set m_Logger = Logger
End Sub
Public Function GetInvoiceSubmissionManager() As IInvoiceSubmissionManager
    Select Case m_InvoiceType
        Case InvoiceTypeEnum.CustomerBSummary
            Set GetInvoiceSubmissionManager = New clsInvoiceSubmissionManagerCustomerB
        Case InvoiceTypeEnum.CustomerASummary
            Set GetInvoiceSubmissionManager = New clsInvoiceSubmissionManagerCustomerA
        Case InvoiceTypeEnum.DefaultIndividual
            Set GetInvoiceSubmissionManager = New clsInvoiceSubmissionManagerReg
        Case Else
        Err.Raise vbObjectError + 1, "InvoiceFactory", "Unknown Invoice Type"
    End Select
End Function
Public Function GetInvoiceSetupMethod() As String
    Select Case m_InvoiceType
        Case InvoiceTypeEnum.CustomerBSummary
            GetInvoiceSetupMethod = "SetCustomerBSubmissions"
        Case InvoiceTypeEnum.CustomerASummary
            GetInvoiceSetupMethod = "SetCustomerASubmissions"
        Case InvoiceTypeEnum.DefaultIndividual
            GetInvoiceSetupMethod = "SetDefaultIndividualSubmission"
        Case Else
        Err.Raise vbObjectError + 1, "InvoiceFactory", "Unknown Invoice Type"
    End Select

End Function
Public Function GetPricingDict(RequestType As TestingRequestTypeEnum) As Scripting.Dictionary
    Dim PricingDict As Scripting.Dictionary
    
    ' Retrieve the correct pricing dictionary based on the request type
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical
            Set PricingDict = PricingCache.ChemicalPricingDict
        Case TestingRequestTypeEnum.Water
            Set PricingDict = PricingCache.WaterPricingDict
        Case TestingRequestTypeEnum.Wafer
            Set PricingDict = PricingCache.WaferPricingDict
        Case Else
            Err.Raise vbObjectError + 1, "InvoiceFactory", "Unknown Invoice Type"
            Exit Function
    End Select
    
    Set GetPricingDict = PricingDict
End Function
Public Function CreateInvoice() As IInvoice
    Select Case m_InvoiceType
        Case InvoiceTypeEnum.CustomerBSummary
            Set CreateInvoice = New clsInvoiceCustomerBSummary
        Case InvoiceTypeEnum.CustomerASummary
            Set CreateInvoice = New clsInvoiceCustomerASummary
        Case InvoiceTypeEnum.DefaultIndividual
            Set CreateInvoice = New clsInvoiceDefaultIndividual
        Case Else
            Err.Raise vbObjectError + 1, "InvoiceFactory", "Unknown Invoice Type"
    End Select
End Function
Public Function SetLineItemCollection(Submissions As Collection, AccessDB As clsAccessDatabase, PriceRetriever As clsInvoicePriceRetriever, SLIMDict As clsSLIMDictionaryManager) As Collection
    Dim SampleGroupDict As Scripting.Dictionary
    Dim Key As Variant
    Dim LineItem As IInvoiceLineItem
    Dim LineItems As New Collection
    Dim TextColour As Long

    
    ' Set the dictionary that will be used to sort the line items
    Select Case m_InvoiceType
        Case InvoiceTypeEnum.CustomerBSummary, InvoiceTypeEnum.DefaultIndividual
            Set SampleGroupDict = GroupSamplesByAnalysesProcessingTimeDateReceived(Submissions)
        Case InvoiceTypeEnum.CustomerASummary
            Set SampleGroupDict = GroupByAnalysis(Submissions)
        Case Else
            Err.Raise vbObjectError + 1, "InvoiceFactory", "Unknown Invoice Type"
    End Select
    
    ' Loop through collection in the dictionary
    For Each Key In SampleGroupDict
        ' Alternate color for next line item
        TextColour = AlternateTextColour(TextColour)
        
        Set LineItem = CreateLineItem
        Call LineItem.Initialize(SampleGroupDict(Key), TextColour, m_Logger, AccessDB, PriceRetriever, SLIMDict)
        LineItems.Add LineItem
        
    Next Key
    
    ' Return the collection to the calling sub
    Set SetLineItemCollection = LineItems
    
End Function
Private Function GroupSamplesByAnalysesProcessingTimeDateReceived(Submissions As Collection) As Scripting.Dictionary
    Dim Submission As clsTestingRequestSubmission
    Dim Sample As clsTestingRequestSample
    Dim ProcessingTime As String
    Dim Analyses As String
    Dim AdditionalElements As String
    Dim DateReceived As Date
    Dim MatchKey As String
    Dim SampleGroup As New Scripting.Dictionary

    For Each Submission In Submissions
        For Each Sample In Submission.Samples
            
            ' Generate the match key
            ProcessingTime = Sample.ProcessingTime
            Analyses = Sample.AnalysesStr
            AdditionalElements = Sample.AdditionalElementStr
            DateReceived = Submission.DateReceived
            MatchKey = ProcessingTime & "|" & Analyses & "|" & AdditionalElements & "|" & DateReceived
            If MatchKey = "1|||9/16/2025" Then
                Debug.Print MatchKey
            End If
            Call AddSampleDataDictToCollection(SampleGroup, MatchKey, Sample, Submission)
            
        Next Sample
    Next Submission
    
    ' Assign the dictionary to the function name
    Set GroupSamplesByAnalysesProcessingTimeDateReceived = SampleGroup
End Function
Private Function CollectionToString(Collection As Collection) As String
    Dim iIndex As Long
    Dim Str As String
    
    ' Loop through the collection and compile a string
    For iIndex = 1 To Collection.Count
        Str = Str & Collection(iIndex).AnalysisName & "/"
    Next iIndex
    
    ' Remove the trailing "/"
    If Len(Str) > 0 Then
        Str = Left(Str, Len(Str) - 1)
    End If
    
    ' Return the string to the calling sub
    CollectionToString = Str
        
End Function
Private Function GroupByAnalysis(Submissions As Collection) As Scripting.Dictionary
    Dim Submission As clsTestingRequestSubmission
    Dim Sample As clsTestingRequestSample
    Dim Analysis As clsTestingRequestAnalysis
    Dim ProcessingTime As String
    Dim Analyses As String
    Dim MatchKey As String
    Dim SampleGroup As New Scripting.Dictionary
    Dim LineCount As Integer
    
    ' Initialize Line Count - this ensures that each sample and analysis gets a unique match key
    LineCount = 0
    
    For Each Submission In Submissions
        For Each Sample In Submission.Samples
            For Each Analysis In Sample.Analyses
                ' Increment the line count
                LineCount = LineCount + 1
                
                ' Generate the match key
                MatchKey = Submission.POInformation & "|" & Sample.SampleName & "|" & Analysis.AnalysisName & "|" & LineCount
                
                Call AddSampleDataDictToCollection(SampleGroup, MatchKey, Sample, Submission, Analysis)
            Next Analysis
        Next Sample
    Next Submission
    
    ' Assign the dictionary to the function name
    Set GroupByAnalysis = SampleGroup
End Function
Private Function GroupBySample(Submissions As Collection) As Scripting.Dictionary
    Dim Submission As clsTestingRequestSubmission
    Dim Sample As clsTestingRequestSample
    Dim MatchKey As String
    Dim SampleGroup As New Scripting.Dictionary
   
    
    For Each Submission In Submissions
        For Each Sample In Submission.Samples
            
            ' Generate the match key
            MatchKey = Submission.POInformation & "|" & Sample.SampleName
            
            Call AddSampleDataDictToCollection(SampleGroup, MatchKey, Sample, Submission)

        Next Sample
    Next Submission
    
    ' Assign the dictionary to the function name
    Set GroupBySample = SampleGroup
End Function
Public Function CreateLineItem() As IInvoiceLineItem
    Select Case m_InvoiceType
        Case InvoiceTypeEnum.CustomerBSummary
            Set CreateLineItem = New clsInvoiceLineItemCustomerBSummary
        Case InvoiceTypeEnum.CustomerASummary
            Set CreateLineItem = New clsInvoiceLineItemCustomerASummary
        Case InvoiceTypeEnum.DefaultIndividual
            Set CreateLineItem = New clsInvoiceLineItemRegIndividual
        Case Else
            Err.Raise vbObjectError + 1, "InvoiceFactory", "Unknown LineItem Type"
    End Select
End Function
Private Function AlternateTextColour(TextColour As Long) As Long
    Const PRECILAB_BLUE As Long = 16711680
    If TextColour = vbBlack Then
        AlternateTextColour = PRECILAB_BLUE
    Else
        AlternateTextColour = vbBlack
    End If
End Function
Public Function SetSampleRowCollection(Submissions As Collection) As Collection
    Dim SampleGroupColl As Collection
    Dim SampleDataDict As Scripting.Dictionary
    Dim SampleRow As IInvoiceSampleRow
    Dim SampleRows As New Collection
    Dim TextColour As Long
    Dim Submission As clsTestingRequestSubmission
    Dim Sample As clsTestingRequestSample
    
    ' Set the collection that will be used to sort the line items
    Select Case m_InvoiceType
        Case InvoiceTypeEnum.CustomerBSummary, InvoiceTypeEnum.DefaultIndividual
            Set SampleGroupColl = ReorderSamplesForSampleID(GroupSamplesByAnalysesProcessingTimeDateReceived(Submissions))
        Case InvoiceTypeEnum.CustomerASummary
            Set SampleGroupColl = ReorderSamplesForSampleID(GroupBySample(Submissions))
        Case Else
            Err.Raise vbObjectError + 1, "InvoiceFactory", "Unknown Invoice Type"
    End Select
    
    ' Loop through collection
    For Each SampleDataDict In SampleGroupColl
        
        ' Retrieve variables from the dictionary
        Set Submission = SampleDataDict("Submission")
        Set Sample = SampleDataDict("Sample")
        TextColour = SampleDataDict("TextColour")
        
        Set SampleRow = CreateSampleRow
        Call SampleRow.Initialize(Submission, Sample, TextColour, m_Logger)
        SampleRows.Add SampleRow
        
    Next SampleDataDict
    
    ' Return the collection to the calling sub
    Set SetSampleRowCollection = SampleRows
    
End Function
Private Function ReorderSamplesForSampleID(LineItemDict As Scripting.Dictionary) As Collection
    Dim OrderedSampleColl As New Collection
    Dim Key As Variant
    Dim SampleDataDict As Scripting.Dictionary
    Dim Sample As clsTestingRequestSample
    Dim TextColour As Long

    ' Loop through the already grouped and ordered dictionary
    For Each Key In LineItemDict.Keys
        TextColour = AlternateTextColour(TextColour)
        For Each SampleDataDict In LineItemDict(Key)
            ' Add the text colour to the dict
            SampleDataDict.Add "TextColour", TextColour
            ' Add the dict to the new collection
            OrderedSampleColl.Add SampleDataDict
        Next SampleDataDict
    Next Key

    ' Return the correctly ordered collection
    Set ReorderSamplesForSampleID = OrderedSampleColl
End Function
Public Function AddMinimumCostLineItem() As IInvoiceLineItem
    Select Case m_InvoiceType
        Case InvoiceTypeEnum.DefaultIndividual
            Set AddMinimumCostLineItem = New clsInvoiceLineItemRegIndividual
        Case InvoiceTypeEnum.CustomerBSummary
            Set AddMinimumCostLineItem = New clsInvoiceLineItemCustomerBSummary
        Case InvoiceTypeEnum.CustomerASummary
            Set AddMinimumCostLineItem = New clsInvoiceLineItemCustomerASummary
        Case Else
            Err.Raise vbObjectError + 1, "InvoiceFactory", "Unknown InvoiceLineItem Type"
    End Select
End Function
Public Function CreateSampleRow() As IInvoiceSampleRow
    Select Case m_InvoiceType
        Case InvoiceTypeEnum.CustomerASummary
            Set CreateSampleRow = New clsInvoiceSampleRowCustomerA
        Case InvoiceTypeEnum.DefaultIndividual, InvoiceTypeEnum.CustomerBSummary
            Set CreateSampleRow = New clsInvoiceSampleRowReg
        Case Else
            Err.Raise vbObjectError + 1, "InvoiceFactory", "Unknown SampleRow Type"
    End Select
End Function

Private Sub AddSampleDataDictToCollection(SampleGroup As Scripting.Dictionary, MatchKey As String, Sample As clsTestingRequestSample, Submission As clsTestingRequestSubmission, Optional Analysis As clsTestingRequestAnalysis = Nothing)
    Dim SampleCollection As Collection
    Dim SampleDataDict As Scripting.Dictionary
    
    If Not SampleGroup.Exists(MatchKey) Then
        Set SampleCollection = New Collection
        SampleGroup.Add MatchKey, SampleCollection
    End If
    
    ' Add the sample to the collection for this key
    Set SampleDataDict = New Scripting.Dictionary
    SampleDataDict.Add "Sample", Sample
    SampleDataDict.Add "Submission", Submission
    SampleDataDict.Add "Analysis", Analysis
    
    SampleGroup(MatchKey).Add SampleDataDict
End Sub
Public Function GetKeyByCustomerType(Submission As clsTestingRequestSubmission) As String
    Dim Key As String
    
    ' Apply filtering or grouping logic based on the m_InvoiceTypeEnum
    Select Case m_InvoiceType
        Case InvoiceTypeEnum.CustomerBSummary
            Key = Submission.Customer.CustomerName ' & "_" & Submission.POInformation
        Case InvoiceTypeEnum.CustomerASummary
            Key = Submission.Customer.CustomerName
        Case InvoiceTypeEnum.DefaultIndividual
            Key = "Individual Invoice" ' ie there are no rules for grouping passing a collection of one submission
        Case Else
             Err.Raise vbObjectError + 1, "InvoiceFactory", "Unknown InvoiceType"
    End Select
    
    ' Return the string to the calling sub
    GetKeyByCustomerType = Key
End Function
