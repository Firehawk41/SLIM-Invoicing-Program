VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceLineItemBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceLineItemBase
' Version: 1.1.1
' Author: JT
' Created: ?
' LastModified: 2025-10-27
' Description: Base Line Item object for shared functions and variables
' DependsOn: clsLoggingSystem, clsInvoicePriceRetriever, clsSLIMDictionaryManager, modEnums, clsInvoiceLineItemFormattedText, clsTestingRequestSample, clsTestingRequestSubmission, clsInvoicePriceDetail
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
'   - 1.1.0 - 2025-07-07 - Updated pricing dictionary to use clsInvoicePriceDetail instead of just a price.  This involved a rewrite to get the logic to work as I needed to set an object
'   - 1.1.1 - 2025-07-08 - Fixed bug where additional element count was not calculated
'   - 1.2.0 - 2025-10-27 - Refactored additional elements multiplier and description into helper functions so that they can be accessed internally and externally
Private m_Logger As clsLoggingSystem
Private m_PriceRetriever As clsInvoicePriceRetriever
Private m_PricingDict As Scripting.Dictionary
Private m_GroupedSamples As Collection
Private m_AnalysisDescriptionDict As Scripting.Dictionary
Private m_Sample As clsTestingRequestSample
Private Enum ArrayTypeEnum
    DescriptionArray = 1
    UnitPriceArray
    ExtendedPriceArray
End Enum
Public Property Get ProcessingTime() As clsInvoiceLineItemFormattedText()
    ProcessingTime = GetProcessingTime()
End Property
Public Property Get ExtendedPrice(Quantity As Long) As clsInvoiceLineItemFormattedText()
    ' Set the pricing dict if not already initiated
    If m_PricingDict Is Nothing Then
        Call RetrievePricingDict(Quantity)
    End If
    
    If m_PricingDict.Count > 1 Then
        ExtendedPrice = CompileArrayFromDict(Quantity, ArrayTypeEnum.ExtendedPriceArray)
    Else
        ExtendedPrice = CompileSingleUnitPrice(Quantity, ArrayTypeEnum.ExtendedPriceArray)
    End If
    
End Property
Public Property Get Description(Quantity As Long) As clsInvoiceLineItemFormattedText()
    Description = CompileArrayFromDict(Quantity, ArrayTypeEnum.DescriptionArray)
End Property
Public Function UnitPrices(Quantity As Long) As clsInvoiceLineItemFormattedText()
        ' Set the pricing dict if not already initiated
    If m_PricingDict Is Nothing Then
        Call RetrievePricingDict(Quantity)
    End If
    
    If m_PricingDict.Count > 1 Then
        UnitPrices = CompileArrayFromDict(Quantity, ArrayTypeEnum.UnitPriceArray)
    Else
        UnitPrices = CompileSingleUnitPrice(Quantity, ArrayTypeEnum.UnitPriceArray)
    End If
        
End Function
Public Function ReplaceAdditionalElements(AnalysisDescription As String, Sample As clsTestingRequestSample) As String
    ReplaceAdditionalElements = Replace(AnalysisDescription, "##", Sample.AdditionalElementStr)
End Function
Public Function MultiplierForAdditionalElements(Sample As clsTestingRequestSample) As Double
   
    ' If additional elements exist, return the count
    If Not Sample.AdditionalElements Is Nothing Then
        MultiplierForAdditionalElements = Sample.AdditionalElements.Count
    End If
    
    ' Default to 1 if the count is less than 1
    If MultiplierForAdditionalElements < 1 Then MultiplierForAdditionalElements = 1
End Function
Public Sub Initialize(GroupedSamples As Collection, PriceRetriever As clsInvoicePriceRetriever, SLIMDict As clsSLIMDictionaryManager, Logger As clsLoggingSystem)

    Set m_Logger = Logger
    Set m_PriceRetriever = PriceRetriever
    Set m_GroupedSamples = GroupedSamples
    Set m_Sample = m_GroupedSamples(1)("Sample")
    Set m_AnalysisDescriptionDict = SLIMDict.AnalysisDescriptionDict
    
End Sub
Private Function CompileSingleUnitPrice(Quantity As Long, ArrayType As ArrayTypeEnum) As clsInvoiceLineItemFormattedText()
    Dim ResultArray(0) As New clsInvoiceLineItemFormattedText
    Dim AdditionalElementsMultiplier As Long
    Dim PriceDetail As clsInvoicePriceDetail
    Dim PriceKey As String
    Dim UnitPrice As Long
    
    ' Get the dictionary key and value
    PriceKey = m_PricingDict.Keys(0)
    Set PriceDetail = m_PricingDict(PriceKey)

    ' Determine the multiplier
    If PriceKey = "Additional Element" And Not PriceDetail.IsFlatRate Then
        AdditionalElementsMultiplier = MultiplierForAdditionalElements(m_Sample)
    Else
        AdditionalElementsMultiplier = 1
    End If

    ' Compute the unit price
    UnitPrice = PriceDetail.Price * AdditionalElementsMultiplier

    ' Populate the array
    Select Case ArrayType
        Case ArrayTypeEnum.UnitPriceArray
            Call ResultArray(0).Initialize(format(UnitPrice, "#,###.00"), m_Logger)
        Case ArrayTypeEnum.ExtendedPriceArray
            Call ResultArray(0).Initialize(format(UnitPrice * Quantity, "#,###.00"), m_Logger)
    End Select

    ' Return array
    CompileSingleUnitPrice = ResultArray
End Function
Private Function CompileArrayFromDict(Quantity As Long, ArrayType As ArrayTypeEnum) As clsInvoiceLineItemFormattedText()
    Dim ResultArray() As clsInvoiceLineItemFormattedText
    Dim Key As Variant
    Dim KeyCount As Long
    Dim AdditionalElementsMultiplier As Long
    Dim PriceDetail As clsInvoicePriceDetail
    Dim Multiplier As Long
    
    ' Set the pricing dict if not already initiated
    If m_PricingDict Is Nothing Then
        Call RetrievePricingDict(Quantity)
    End If
    
    ' Resize array
    ReDim ResultArray(m_PricingDict.Count)
    KeyCount = 1
    
    ' Set the first value of the array as the sample matrix for the description
    If ArrayType = ArrayTypeEnum.DescriptionArray Then
        Set ResultArray(0) = GetSampleMatrices()
    End If
    
    ' Populate the array
    For Each Key In m_PricingDict.Keys
        Set PriceDetail = m_PricingDict(Key)
        
        ' Determine multiplier
        If Key = "Additional Element" And Not PriceDetail.IsFlatRate Then
            Multiplier = MultiplierForAdditionalElements(m_Sample)
        Else
            Multiplier = 1
        End If

        Set ResultArray(KeyCount) = New clsInvoiceLineItemFormattedText

        Select Case ArrayType
            Case ArrayTypeEnum.DescriptionArray
                Call ResultArray(KeyCount).Initialize(vbTab & GetAnalysisDescription(CStr(Key)), m_Logger, False)
            Case ArrayTypeEnum.UnitPriceArray
                Call ResultArray(KeyCount).Initialize(format(PriceDetail.Price * Multiplier, "#,###.00"), m_Logger)
            Case ArrayTypeEnum.ExtendedPriceArray
                Call ResultArray(KeyCount).Initialize(format(PriceDetail.Price * Quantity * Multiplier, "#,###.00"), m_Logger)
        End Select

        KeyCount = KeyCount + 1
    Next Key
    
    ' Return array
    CompileArrayFromDict = ResultArray
End Function
Private Sub RetrievePricingDict(Quantity As Long)
    Dim Sample As clsTestingRequestSample
    Dim Submission As clsTestingRequestSubmission
    
    Set Sample = m_GroupedSamples(1)("Sample")
    Set Submission = m_GroupedSamples(1)("Submission")
    Set m_PricingDict = m_PriceRetriever.RetrievePrices(Submission, Sample, Quantity)
End Sub

Private Function GetAnalysisDescription(AnalysisName As String) As String
    Dim AnalysisArr() As String
    Dim ArrayIndex As Long
    Dim AnalysisDescription As String
    
    ' Convert the string to an array
    AnalysisArr = Split(AnalysisName, "/")
    
    ' Loop through the array to compile a string of analysis descriptions
    For ArrayIndex = LBound(AnalysisArr) To UBound(AnalysisArr)
        If m_AnalysisDescriptionDict.Exists(AnalysisArr(ArrayIndex)) Then
            ' Retrieve the analysis description
            AnalysisDescription = AnalysisDescription & m_AnalysisDescriptionDict(AnalysisArr(ArrayIndex)).AnalysisDescription & " / "
            ' Add element symbols to any "Additional Elements"
            If AnalysisArr(ArrayIndex) = "Additional Element" Then
                AnalysisDescription = ReplaceAdditionalElements(AnalysisDescription, m_Sample)
            End If
        End If
    Next ArrayIndex
    
    ' Trim any trailing slashes
    If Right(AnalysisDescription, 3) = " / " Then
        AnalysisDescription = Left(AnalysisDescription, Len(AnalysisDescription) - 3)
    End If
    
    ' Return the string to the calling sub
    GetAnalysisDescription = AnalysisDescription
    
End Function
Private Function GetSampleMatrices() As clsInvoiceLineItemFormattedText
    Dim UniqueMatrices As Scripting.Dictionary
    Dim SampleMatrices As New clsInvoiceLineItemFormattedText
    
    ' Set dictionary of unique matrices
    Set UniqueMatrices = CompileUniqueMatrices()
    
    ' Compile the keys into a formatted text object
    Call SampleMatrices.Initialize(CompileMatrixString(UniqueMatrices), m_Logger)
    
    ' Return to the calling sub
    Set GetSampleMatrices = SampleMatrices

End Function
Private Function CompileUniqueMatrices() As Scripting.Dictionary
    Dim UniqueMatrices As New Scripting.Dictionary
    Dim Item As Variant
    Dim Sample As clsTestingRequestSample
    Dim SampleMatrix As String
    
        For Each Item In m_GroupedSamples
        ' Set the variables
        Set Sample = Item("Sample")
        SampleMatrix = Sample.FormChemicalName
        
        ' Store in dictionary if not already present
        If Not UniqueMatrices.Exists(SampleMatrix) Then
            UniqueMatrices.Add SampleMatrix, True
        End If
        
    Next Item
    
    ' Return the dictionary to the calling sub
    Set CompileUniqueMatrices = UniqueMatrices
End Function
Private Function CompileMatrixString(UniqueMatrices As Scripting.Dictionary) As String
    Dim MatrixKey As Variant
    Dim TempStr As String
    Dim SampleMatrices As String
    
    Const MAX_CHARACTERS As Long = 80
    
    ' Build the output string with a max length of 80 characters
    For Each MatrixKey In UniqueMatrices.Keys
        TempStr = IIf(SampleMatrices = "", MatrixKey, SampleMatrices & ", " & MatrixKey)

        ' Check if adding the next matrix would exceed 80 characters
        If Len(TempStr) > MAX_CHARACTERS Then
            SampleMatrices = SampleMatrices & ", etc"
            Exit For
        Else
            SampleMatrices = TempStr
        End If
    Next MatrixKey

    ' If all matrices fit within the limit, return normally
    If Right(SampleMatrices, 5) = ", etc" Then
        CompileMatrixString = "Various chemicals (" & SampleMatrices & ")"
    Else
        CompileMatrixString = SampleMatrices
    End If
End Function
Private Function GetProcessingTime() As clsInvoiceLineItemFormattedText()
    Dim Submission As clsTestingRequestSubmission
    Dim Sample As clsTestingRequestSample
    Dim ProcessingTime As ProcessingTimeEnum
    Dim ProcessingTimeStr(0) As New clsInvoiceLineItemFormattedText
    Dim RequestType As TestingRequestTypeEnum
    
    ' Set the processing time variable
    Set Submission = m_GroupedSamples(1)("Submission")
    Set Sample = m_GroupedSamples(1)("Sample")
    RequestType = Submission.RequestType
    ProcessingTime = Sample.ProcessingTime
    
    
    ' Convert the processing time enum to a string
    Select Case ProcessingTime
        Case ProcessingTimeEnum.ExtendedTime
            Call ProcessingTimeStr(0).Initialize(SetExtendedTime(RequestType), m_Logger)
        Case ProcessingTimeEnum.NextDay
            Call ProcessingTimeStr(0).Initialize(SetNextDayTime(RequestType), m_Logger)
        Case ProcessingTimeEnum.TimeLimited
            Call ProcessingTimeStr(0).Initialize("Next Day Time Limited", m_Logger)
        Case ProcessingTimeEnum.SameDayRush
            Call ProcessingTimeStr(0).Initialize("Same Day RUSH", m_Logger)
        Case ProcessingTimeEnum.CallInRush
            Call ProcessingTimeStr(0).Initialize("Call-in RUSH", m_Logger)
        Case Else
            Call m_Logger.LogMessage("SetWaterProcessingStr", LogLevelEnum.LogError, "Invalid Processing Time")
    End Select
    
    ' Return the string to the calling sub
    GetProcessingTime = ProcessingTimeStr
    
End Function
Private Function SetExtendedTime(RequestType As TestingRequestTypeEnum) As String
    Select Case RequestType
        Case TestingRequestTypeEnum.Water
            SetExtendedTime = "Two Days"
        Case TestingRequestTypeEnum.Wafer
            SetExtendedTime = "Up To 3 Working Days"
        Case Else
            Call m_Logger.LogMessage("SetExtendedTime", LogLevelEnum.LogError, "Invalid Request Type.")
    End Select
End Function
Private Function SetNextDayTime(RequestType As TestingRequestTypeEnum) As String
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical, TestingRequestTypeEnum.Water
            SetNextDayTime = "Next Day"
        Case TestingRequestTypeEnum.Wafer
            SetNextDayTime = "Next Day RUSH"
        Case Else
            Call m_Logger.LogMessage("SetNextDayTime", LogLevelEnum.LogError, "Invalid Request Type.")
    End Select
End Function
Private Function GetAdditionalElementsCount() As Long
    Dim Sample As clsTestingRequestSample
    Set Sample = m_GroupedSamples(1)("Sample")

    ' If additional elements exist, return the count
    If Not Sample.AdditionalElements Is Nothing Then
        GetAdditionalElementsCount = Sample.AdditionalElements.Count
    End If
    
    ' Default to 1 if the count is less than 1
    If GetAdditionalElementsCount < 1 Then GetAdditionalElementsCount = 1
End Function

