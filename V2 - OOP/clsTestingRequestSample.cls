VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestingRequestSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsTestingRequestSample
' Version: 1.1.2
' Author: JT
' Created: ?
' LastModified: 2025-09-29
' Description: Stores properties relating to an individual sample
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsSLIMDictionaryManager, modEnums, clsTestingRequestChemical, clsTestingRequestAnalysis, clsTestingRequestElement
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
'   - 1.1.0 - 2025-08-12 - Added sub to initialize from json file
'   - 1.1.1 - 2025-08-19 - Fixed bugs for LoadFromJson
'   - 1.1.2 - 2025-09-29 - Added m_AnalysesDict to prevent duplicates in the list of analyses - *** Convert fully to dictionary ***
Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_SLIMDictManager As clsSLIMDictionaryManager
Private m_DatabaseEntryDate As Date
Private m_SampleID As Long
Private m_SampleName As String
Private m_Chemical As clsTestingRequestChemical
Private m_ProcessingTime As ProcessingTimeEnum
Private m_Analyses As Collection
Private m_AnalysesDict As Scripting.Dictionary
Private m_AdditionalElements As Collection
Private m_FormChemicalName As String
Public Property Get FormChemicalName() As String
    FormChemicalName = m_FormChemicalName
End Property
Public Property Get SampleName() As String
    SampleName = m_SampleName
End Property
Public Property Get ProcessingTime() As String
    ProcessingTime = m_ProcessingTime
End Property
Public Property Get Analyses() As Collection
    Set Analyses = m_Analyses
End Property
Public Property Get AnalysesStr() As String
    AnalysesStr = CollectionToString(m_Analyses, "AnalysisName")
End Property
Public Property Get AdditionalElements() As Collection
    Set AdditionalElements = m_AdditionalElements
End Property
Public Property Get AdditionalElementStr() As String
    AdditionalElementStr = CollectionToString(m_AdditionalElements, "ElementSymbol")
End Property
Public Property Get Chemical() As clsTestingRequestChemical
    Set Chemical = m_Chemical
End Property
Public Sub Initialize(SampleCell As Range, SampleRange As Range, RequestType As TestingRequestTypeEnum, SLIMDictManager As clsSLIMDictionaryManager, Logger As clsLoggingSystem, AccessDB As clsAccessDatabase)
    ' Retrieve the sample properties
    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    Set m_SLIMDictManager = SLIMDictManager
    
    m_SampleName = SampleCell.Value
    m_FormChemicalName = GetFormChemicalName(SampleCell.Worksheet, RequestType, SampleCell)
    Set m_Chemical = GetChemical(RequestType)
    m_ProcessingTime = GetProcessingTime(RequestType, SampleCell)
    
    
    ' Add analyses objects to the sample
    Call AddAnalysesToSample(SampleCell, SampleRange, SLIMDictManager.AnalysisDict)
    
    'Add additional elements to the sample (if any)
    Call AddAdditionalElements(RequestType, SampleCell, SampleRange, SLIMDictManager)
    
End Sub
Private Function GetChemical(RequestType As TestingRequestTypeEnum) As clsTestingRequestChemical
    Dim Chemical As New clsTestingRequestChemical
    
    Call Chemical.SetFromDictionary(m_FormChemicalName, RequestType, m_Logger, m_AccessDB, m_SLIMDictManager)

    ' Return the object to the calling sub
    Set GetChemical = Chemical
    
End Function
Private Function GetProcessingTime(RequestType As TestingRequestTypeEnum, SampleCell As Range) As ProcessingTimeEnum
    
    Dim WkS As Worksheet
    Dim ProcessingTimeColumn As Long
    Dim ProcessingTimeEnum As ProcessingTimeEnum
    
    ' Determine the worksheet from the supplied range
    Set WkS = SampleCell.Worksheet
    
    ' Determine the column the processing time is located in
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical
            ProcessingTimeColumn = 11
        Case TestingRequestTypeEnum.Water
            ProcessingTimeColumn = 14
        Case TestingRequestTypeEnum.Wafer
            ProcessingTimeColumn = 9
    End Select
    
    ' Retrieve the processing time string
    ProcessingTimeEnum = GetProcessingTimeEnum(WkS.Cells(SampleCell.Row, ProcessingTimeColumn))
    
    ' Return the processing time to the calling sub
    GetProcessingTime = ProcessingTimeEnum
    
End Function
Private Function GetProcessingTimeEnum(ProcessingTimeStr As String) As ProcessingTimeEnum
    
    Select Case ProcessingTimeStr
        Case "Two Days", "Up To 3 Working Days"
            GetProcessingTimeEnum = ProcessingTimeEnum.ExtendedTime
        Case "Next Day", "Next Day RUSH"
            GetProcessingTimeEnum = ProcessingTimeEnum.NextDay
        Case "Next Day Time Limited"
            GetProcessingTimeEnum = ProcessingTimeEnum.TimeLimited
        Case "Same Day RUSH"
            GetProcessingTimeEnum = ProcessingTimeEnum.SameDayRush
        Case "Call-in RUSH"
            GetProcessingTimeEnum = ProcessingTimeEnum.CallInRush
        Case Else
            Call m_Logger.LogMessage("GetProcessingTime", LogLevelEnum.LogError, "Invalid Processing Time. Defaulting to Next Day.")
            GetProcessingTimeEnum = ProcessingTimeEnum.NextDay
    End Select
    
End Function
Private Function GetFormChemicalName(WkS As Worksheet, RequestType As TestingRequestTypeEnum, SampleCell As Range) As String
    ' Retrieves the chemical name used on the form
    
    Dim FormChemicalName As String
    
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical
            FormChemicalName = WkS.Cells(SampleCell.Row, 3).Value
        Case TestingRequestTypeEnum.Water
            FormChemicalName = "Water"
        Case TestingRequestTypeEnum.Wafer
            FormChemicalName = WkS.Cells(SampleCell.Row, 4).Value
    End Select

    ' Log the value
    Call m_Logger.LogMessage("GetFormChemicalName", LogLevelEnum.LogDebug, "Form chemical name is " & FormChemicalName)
    
    ' Return the value to the calling sub
    GetFormChemicalName = FormChemicalName
    
End Function
Private Sub AddAnalysesToSample(SampleCell As Range, SampleRange As Range, AnalysisDict As Scripting.Dictionary)
    
    Dim AnalysisCell As Range
    Dim SampleRow As Range
    Dim AnalysisCollection As Collection
    Dim Analysis As clsTestingRequestAnalysis
    
    ' Set the sample row range
    Set SampleRow = SampleCell.Resize(1, SampleRange.Columns.Count)
    
    ' Loop through the sample row
    For Each AnalysisCell In SampleRow.Cells
        If Not IsEmpty(AnalysisCell) Then
            ' Retrieve and create the analysis objects
            If AnalysisDict.Exists(AnalysisCell.Value) Then
                Set AnalysisCollection = AnalysisDict(AnalysisCell.Value)
            End If
            
            ' Add the analysis objects to the sample
            If Not AnalysisCollection Is Nothing Then
                For Each Analysis In AnalysisCollection
                    If Not m_AnalysesDict.Exists(Analysis.AnalysisName) Then
                        Set m_AnalysesDict(Analysis.AnalysisName) = Analysis
                        m_Analyses.Add Analysis
                    End If
                Next Analysis
                Set AnalysisCollection = Nothing
            End If
        End If
    Next AnalysisCell
    
End Sub
Private Sub AddAdditionalElements(RequestType As TestingRequestTypeEnum, SampleCell As Range, SampleRange As Range, SLIMDictManager As clsSLIMDictionaryManager)
    
    Dim AddElColumn As Long
    Dim AddElCell As Range
    Dim AddElLabel As Range
    Dim ElementArr As Variant
    Dim iIndex As Long
    Dim ElementSymbol As String
    
    ' Get the column for additional elements based on the request type
    AddElColumn = GetAdditionalElementColumn(RequestType)
    
    ' Get the cell and label for additional elements
    Set AddElCell = SampleCell.Worksheet.Cells(SampleCell.Row, AddElColumn)
    Set AddElLabel = SampleCell.Worksheet.Cells(SampleRange.Row - 1, AddElColumn)
    
    ' If there are additional elements, process them
    If Not IsEmpty(AddElCell) Then
        ' Add "Additional Elements" to the analysis collection
        Call AddAdditionalElementsAnalysis(AddElLabel.Value, SLIMDictManager.AnalysisDict)
        
        ' Split and add elements
        ElementArr = SplitElements(AddElCell.Value)
        Call AddElementsToSample(ElementArr, SLIMDictManager.ElementDict)
    End If
    
End Sub
Private Function GetAdditionalElementColumn(RequestType As TestingRequestTypeEnum) As Long
    Select Case RequestType
        Case TestingRequestTypeEnum.Chemical, TestingRequestTypeEnum.Water
            GetAdditionalElementColumn = 5
        Case TestingRequestTypeEnum.Wafer
            GetAdditionalElementColumn = 6
        Case Else
            GetAdditionalElementColumn = 0 ' Error case, handle accordingly
    End Select
End Function
Private Sub AddAdditionalElementsAnalysis(AdElLabelValue As String, AnalysisDict As Scripting.Dictionary)
    
    Dim AnalysesCollection As Collection
    Dim Analysis As clsTestingRequestAnalysis
    
    ' Retrieve the analysis collection and add the first analysis to m_Analyses
    If AnalysisDict.Exists(AdElLabelValue) Then
        Set AnalysesCollection = AnalysisDict(AdElLabelValue)
        If AnalysesCollection.Count > 0 Then
            Set Analysis = AnalysesCollection(1)
            m_Analyses.Add Analysis
        End If
    End If
    
End Sub
Private Function SplitElements(ElementString As String) As Variant
    Dim Elements() As String
    
    ' Normalize delimiters and split string
    ElementString = Replace(ElementString, ";", ",")
    ElementString = Replace(ElementString, " ", ",")
    ElementString = Replace(ElementString, vbCrLf, ",")
    ElementString = Replace(ElementString, vbCr, ",")
    ElementString = Replace(ElementString, vbLf, ",")
    Elements = Split(ElementString, ",")
    
    ' Clean and return the trimmed array
    SplitElements = TrimArray(Elements)
    
End Function

' Helper function to trim and filter the array
Private Function TrimArray(Elements() As String) As Variant
    Dim CleanElements() As String
    Dim Element As String
    Dim iIndex As Long, Count As Long
    
    ' Initialize dynamic array
    ReDim CleanElements(LBound(Elements) + 1 To UBound(Elements) + 1)
    
    ' Trim whitespace and collect valid entries
    Count = 0
    For iIndex = LBound(Elements) To UBound(Elements)
        Element = Trim(Elements(iIndex))
        If Element <> "" Then
            Count = Count + 1
            CleanElements(Count) = Element ' to account for 0 based array
        End If
    Next iIndex
    
    ' Resize the array to fit valid entries and return
    If Count > 0 Then
        ReDim Preserve CleanElements(1 To Count)
        TrimArray = CleanElements
    Else
        TrimArray = Array("")
    End If
End Function
Private Sub AddElementsToSample(ElementArr As Variant, ElementDict As Scripting.Dictionary)
    Dim iIndex As Long
    Dim ElementSymbol As String
    
    ' Loop through each element in the array and retrieve/add to the sample
    For iIndex = LBound(ElementArr) To UBound(ElementArr)
        ElementSymbol = ElementArr(iIndex)
        Call RetrieveAndAddElement(ElementSymbol, ElementDict)
    Next iIndex
End Sub

Private Sub RetrieveAndAddElement(ElementSymbol As String, ElementDict As Scripting.Dictionary)
    Dim Element As clsTestingRequestElement
    
    ' Check if the element exists and add it if found
    If ElementDict.Exists(ElementSymbol) Then
        Set Element = ElementDict(ElementSymbol)
        m_AdditionalElements.Add Element
    End If
End Sub
Private Function CollectionToString(Collection As Collection, PropertyName As String) As String
    Dim iIndex As Long
    Dim Str As String
    Dim Item As Object
    Dim PropValue As String
    
    ' Loop through the collection and compile a string
    For iIndex = 1 To Collection.Count
        Set Item = Collection(iIndex)
        
        ' Use CallByName to dynamically get the property value
        On Error Resume Next
        PropValue = CStr(CallByName(Item, PropertyName, VbGet))
        On Error GoTo 0
        
        ' Append the property value to the string
        If Len(PropValue) > 0 Then
            Str = Str & PropValue & "/"
        End If
    Next iIndex
    
    ' Remove the trailing "/"
    If Len(Str) > 0 Then
        Str = Left(Str, Len(Str) - 1)
    End If
    
    ' Return the string
    CollectionToString = Str
End Function
Public Sub LoadFromJson(Json As Dictionary, Logger As clsLoggingSystem)
    Call Logger.LogMessage("clsTestingRequestSample.LoadFromJson", LogLevelEnum.LogDebug, "Loading data from json to Sample object.")
    
    If Not Logger.DebugMode Then On Error GoTo ErrorHandler
    
    Set m_Logger = Logger
    
    m_SampleName = Json("sample_name")
    
    m_ProcessingTime = GetProcessingTimeEnum(Json("processing_time"))
    If Json.Exists("form_chemical") Then m_FormChemicalName = Json("form_chemical")("name")
    
    
    Set m_Chemical = New clsTestingRequestChemical
    Call m_Chemical.LoadFromJson(Json("form_chemical"), Logger)
    
    Dim AnalysisJson As Dictionary
    Dim Analysis As clsTestingRequestAnalysis
    For Each AnalysisJson In Json("form_analyses")
        Set Analysis = New clsTestingRequestAnalysis
        Call Analysis.LoadFromJson(AnalysisJson, Logger)
        m_Analyses.Add Analysis
        ' Check for additional elements
        If AnalysisJson.Exists("elements") Then
            If Not IsNull(AnalysisJson("elements")) Then Call LoadElementsFromJson(AnalysisJson("elements"))
        End If
    Next AnalysisJson
    
    Call Logger.LogMessage("clsTestingRequestSample.LoadFromJson", LogLevelEnum.LogDebug, "Completed loading data from json to Sample object: " & m_SampleName & ".")
    Exit Sub
ErrorHandler:
    Dim ErrorMessage As String
     ErrorMessage = "Error encountered in clsTestingRequestSample.LoadFromJson where the sample is: " & m_SampleName
    Call Logger.LogMessage("clsTestingRequestSample.LoadFromJson", LogLevelEnum.LogError, ErrorMessage)
    Err.Raise vbObjectError + 1000, "clsTestingRequestSample.LoadFromJson", ErrorMessage
End Sub
Private Sub LoadElementsFromJson(Json As Dictionary)
    Dim ElementJson As Dictionary
    Dim Element As clsTestingRequestElement
    Dim ElementStr As String
    
    If Not m_Logger.DebugMode Then On Error GoTo ErrorHandler
    
    Call m_Logger.LogMessage("clsTestingRequestSample.LoadElementsFromJson", LogLevelEnum.LogDebug, "Loading additional element data from json to Sample object.")
    
    
    ' Loop through the dictionary and add any elements
    For Each ElementJson In AnalysisJson("elements")
        Set Element = New clsTestingRequestElement
        Call Element.LoadFromJson(ElementJson, Logger)
        m_AdditionalElements.Add Element
        ElementStr = ElementStr & Element.ElementSymbol & ", "
    Next ElementJson
    
    ' Trim any trailing ", "
    If Len(ElementStr) > 2 Then ElementStr = Left(ElementStr, Len(ElementStr) - 2)
    
    
    Call m_Logger.LogMessage("clsTestingRequestSample.LoadElementsFromJson", LogLevelEnum.LogDebug, "Completed loading additional elements, '" & ElementStr & "'.")
    Exit Sub
ErrorHandler:
    Dim ErrorMsg As String
    ErrorMsg = "Error in clsTestingRequestSample.LoadElementsFromJson.  The elements that were successfully loaded are: " & ElementStr
    Call m_Logger.LogMessage("clsTestingRequestSample.LoadElementsFromJson", LogLevelEnum.LogError, ErrorMsg)
    Err.Raise vbObjectError + 1000, "clsTestingRequestSample", ErrorMsg
End Sub
Public Sub Class_Initialize()
    ' Initialize the Analyses collection
    Set m_Analyses = New Collection
    Set m_AdditionalElements = New Collection
    Set m_AnalysesDict = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    Set m_Analyses = Nothing
    Set m_AdditionalElements = Nothing
    Set m_AnalysesDict = Nothing
End Sub
