VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceQuoteBundle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceQuoteBundle
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-14
' Description: Object to hold multiple analyses for one price
' DependsOn: clsLoggingSystem, clsTestingRequestAnalysis, clsTestingRequestElement
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Private m_Logger As clsLoggingSystem
Private m_Analyses As Collection
Private m_Elements As Collection
Private m_RegularPrice As Long
Private m_BulkPrice As Long
Private m_BulkMin As Long
Public Property Get Analyses() As Collection
    Set Analyses = m_Analyses
End Property
Public Property Get AnalysisStr() As String
    AnalysisStr = ConcatenateAnalyses()
End Property
Public Property Get Elements() As Collection
    Set Elements = m_Elements
End Property
Public Property Get Price(Quantity As Long) As Long
    If m_BulkPrice <> -1 And Not Quantity < m_BulkMin Then
        Price = m_BulkPrice
    Else
        Price = m_RegularPrice
    End If
End Property
Public Property Get IsPriceBundle() As Boolean
    IsPriceBundle = Me.Elements.Count > 0 Or Me.Analyses.Count > 1
End Property
Public Sub Initialize(RegularPrice As Long, BulkPrice As Long, BulkMin As Long, AnalysesCollection As Collection, ElementCollection As Collection, Logger As clsLoggingSystem)
    On Error GoTo ErrorHandler

    Set m_Logger = Logger

    m_RegularPrice = RegularPrice
    m_BulkPrice = BulkPrice
    m_BulkMin = BulkMin
    Set m_Analyses = AnalysesCollection
    Set m_Elements = ElementCollection
    Exit Sub
ErrorHandler:
    Call Logger.LogMessage("Initialize", LogLevelEnum.LogError, "Error initializing quote bundle: " & Err.Number & " " & Err.Description)
    Err.Raise Err.Number, "clsInvoiceQuoteBundle.Initialize", Err.Description

End Sub
Public Function MatchBundle(AnalysesDict As Scripting.Dictionary) As Boolean
    Dim AnalysisName As Variant

    ' Iterate over each analysis name in the bundle's analyses collection
    For Each AnalysisName In Me.Analyses

        ' Check whether the analysis name exists in the dictionary of analyses that the customer ordered
        If AnalysesDict.Exists(AnalysisName) Then

            ' Check whether the analysis has already been matched
            If AnalysesDict(AnalysisName) = True Then
                ' Exit early
                MatchBundle = False
                Exit Function
            Else
                If AnalysisName = "Additional Element" And m_Elements.Count > 0 Then
                    If Not MatchAdditionalElements(AnalysesDict) Then
                        MatchBundle = False
                        Exit Function
                    End If
                End If

                MatchBundle = True
            End If
        Else
            MatchBundle = False
            Exit Function
        End If
    Next AnalysisName

    If MatchBundle = True Then
        Call SetAnalysesMatchToTrue(AnalysesDict)
    End If

End Function
Private Function MatchAdditionalElements(AnalysesDict As Scripting.Dictionary) As Boolean
    Dim ElementSymbol As Variant
    Dim MatchedCount As Long
    Dim ElementDict As Scripting.Dictionary
    Const ElementSymbol_DICT As String = "__ElementDict"
    MatchedCount = 0

    ' Set the dictionary of additional elements
    Set ElementDict = AnalysesDict(ElementSymbol_DICT)


    ' Iterate over each ElementSymbol in the elements collection
    For Each ElementSymbol In m_Elements
        ' Check whether the ElementSymbol exists in the analyses dictionary
        If ElementDict.Exists(ElementSymbol) Then
            ' If found and not already matched, mark as matched
            If ElementDict(ElementSymbol) = False Then
                MatchedCount = MatchedCount + 1
            End If
        End If
    Next ElementSymbol

    ' Only return true if all elements in the bundle are matched
    MatchAdditionalElements = (MatchedCount = m_Elements.Count)
End Function
Public Sub SetAnalysesMatchToTrue(AnalysesDict As Scripting.Dictionary)
    Dim AnalysisName As Variant
    Dim Element As Variant

    ' Set each analysis in the dictionary to true
    For Each AnalysisName In m_Analyses
        If AnalysesDict.Exists(AnalysisName) Then
            AnalysesDict(AnalysisName) = True
        End If
    Next AnalysisName

    ' Set each element in the dictionary to true
    For Each Element In m_Elements
        If AnalysesDict.Exists(Element) Then
            AnalysesDict(Element) = True
        End If
    Next Element

End Sub

Private Function CollectionContains(ByRef TargetCollection As Collection, ByVal SearchValue As Variant) As Boolean
    Dim Item As Variant

    For Each Item In TargetCollection
        If Item = SearchValue Then
            CollectionContains = True
            Exit Function
        End If
    Next Item

    CollectionContains = False
End Function
Private Sub AddAllItemsToCollection(MatchedAnalysis As Collection)
    Dim AnalysisName As Variant

    ' Add each item in the m_analyses collection to the MatchedAnalyses collection
    For Each AnalysisName In m_Analyses
        MatchedAnalysis.Add AnalysisName
    Next AnalysisName

End Sub
Private Function ConcatenateAnalyses() As String
    Dim AnalysisName As Variant
    Dim AnalysisStr As String

    ' Concatenate the analyses
    For Each AnalysisName In m_Analyses
        AnalysisStr = AnalysisStr & AnalysisName & "/"
    Next AnalysisName

    ' Trim any trailing slash
    If Right(AnalysisStr, 1) = "/" Then
        AnalysisStr = Left(AnalysisStr, Len(AnalysisStr) - 1)
    End If

    ' Return the string to the calling sub
    ConcatenateAnalyses = AnalysisStr
End Function


Private Sub Class_Initialize()
    Set m_Analyses = New Collection
    Set m_Elements = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Analyses = Nothing
    Set m_Elements = Nothing
End Sub
