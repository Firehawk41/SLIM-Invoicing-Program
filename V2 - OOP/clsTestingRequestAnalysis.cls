VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestingRequestAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsTestingRequestAnalysis
' Version: 1.1.1
' Author: JT
' Created: ?
' LastModified: 2025-08-19
' Description: Stores properties relating to an individual analysis
' DependsOn: clsLoggingSystem
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
'   - 1.1.0 - 2025-08-12 - Added sub to initialize from json file
'   - 1.1.1 - 2025-08-19 - Fixed bugs for LoadFromJson
Private m_Logger As clsLoggingSystem
Private m_AnalysisID As Long
Private m_AnalysisName As String
Private m_AnalysisDescription As String
Public Property Get AnalysisID() As Long
    AnalysisID = m_AnalysisID
End Property
Public Property Get AnalysisName() As String
    AnalysisName = m_AnalysisName
End Property
Public Property Let AnalysisDescription(Value As String)
    m_AnalysisDescription = Value
End Property
Public Property Get AnalysisDescription() As String
    AnalysisDescription = m_AnalysisDescription
End Property

Public Sub Initialize(AnalysisID As Long, AnalysisName As String, AnalysisDescription As String, Logger As clsLoggingSystem)
    Set m_Logger = Logger
    m_AnalysisID = AnalysisID
    m_AnalysisName = AnalysisName
    m_AnalysisDescription = AnalysisDescription

End Sub

Public Sub LoadFromJson(Json As Dictionary, Logger As clsLoggingSystem)
    Call Logger.LogMessage("clsTestingRequestAnalysis.LoadFromJson", LogLevelEnum.LogDebug, "Loading data from json to Analysis object.")
    
    If Not Logger.DebugMode Then On Error GoTo ErrorHandler
    
    Set m_Logger = Logger
    
    If Json.Exists("analysis") Then
        m_AnalysisID = Json("analysis")("id")
        m_AnalysisName = Json("analysis")("name")
        m_AnalysisDescription = Json("analysis")("description")
    End If
    
    Call Logger.LogMessage("clsTestingRequestAnalysis.LoadFromJson", LogLevelEnum.LogDebug, "Completed loading data from json to Analysis object: " & m_AnalysisName & ".")
    Exit Sub
ErrorHandler:
    Dim ErrorMessage As String
    ErrorMessage = "Error encountered in clsTestingRequestAnalysis.LoadFromJson where the element is: " & m_AnalysisName
    Call Logger.LogMessage("clsTestingRequestAnalysis.LoadFromJson", LogLevelEnum.LogError, ErrorMessage)
    Err.Raise vbObjectError + 1000, "clsTestingRequestAnalysis.LoadFromJson", ErrorMessage
End Sub
