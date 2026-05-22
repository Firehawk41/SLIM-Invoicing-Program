VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestingRequestElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsTestingRequestElement
' Version: 1.1.1
' Author: JT
' Created: ?
' LastModified: 2025-08-19
' Description: Stores properties relating to an individual element
' DependsOn: clsLoggingSystem
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
'   - 1.1.0 - 2025-08-12 - Added sub to initialize from json file
'   - 1.1.1 - 2025-08-19 - Fixed bugs for LoadFromJson
Private m_Logger As clsLoggingSystem
Private m_ElementID As Long
Private m_ElementSymbol As String
Private m_ElementName As String

' Properties for ElementID, ElementSymbol, ElementName
Public Property Get ElementID() As Long
    ElementID = m_ElementID
End Property
Public Property Get ElementSymbol() As String
    ElementSymbol = m_ElementSymbol
End Property
Public Property Get ElementName() As String
    ElementName = m_ElementName
End Property

Public Sub Initialize(ElementID As Long, ElementSymbol As String, ElementName As String, Logger As clsLoggingSystem)
    Set m_Logger = Logger
    m_ElementID = ElementID
    m_ElementSymbol = ElementSymbol
    m_ElementName = ElementName
    
End Sub
Public Sub LoadFromJson(Json As Dictionary, Logger As clsLoggingSystem)
    Call Logger.LogMessage("clsTestingRequestElement.LoadFromJson", LogLevelEnum.LogDebug, "Loading data from json to Element object.")
    
    If Not Logger.DebugMode Then On Error GoTo ErrorHandler
    
    Set m_Logger = Logger
    
    m_ElementID = Json("id")
    m_ElementSymbol = Json("symbol")
    m_ElementName = Json("name")
    
    Call Logger.LogMessage("clsTestingRequestElement.LoadFromJson", LogLevelEnum.LogDebug, "Completed loading data from json to Element object: " & m_ElementName & ".")
    Exit Sub
ErrorHandler:
    Dim ErrorMessage As String
     ErrorMessage = "Error encountered in clsTestingRequestElement.LoadFromJson where the element is: " & m_ElementSymbol
    Call Logger.LogMessage("clsTestingRequestElement.LoadFromJson", LogLevelEnum.LogError, ErrorMessage)
    Err.Raise vbObjectError + 1000, "clsTestingRequestElement.LoadFromJson", ErrorMessage
End Sub
