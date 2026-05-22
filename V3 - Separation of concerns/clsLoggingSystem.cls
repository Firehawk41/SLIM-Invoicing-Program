VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLoggingSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsLoggerSystem
' Version: 1.1.2
' Author: JT
' Created: 2025-??-??
' LastModified: 2025-11-06
' Description: Class used to log the progress of other macros - in a log file and in the immediate window
' DependsOn: modEnums
' ChangeLog:
'   - 1.0.0 - Initial release
'   - 1.1.0 - 2025-04-04    - tidied up formatting for new lines, debug.print, and message boxes; Added feature to 'stack' repeated messages to tidy up log files
'   - 1.1.1 - 2025-05-21    - added metadata
'   - 1.1.2 - 2025-11-06    - Fixed bug on Log Error that refered to err.number etc before the error was raised
Private m_DebugMode As Boolean
Private m_LogFile As Scripting.TextStream
Private m_ComputerName As String
Private m_LogFilePath As String
Private m_LastMessage As String
Private m_LastSubName As String
Private m_LastLevel As LogLevelEnum
Private m_RepeatCount As Long
Private m_LastFormattedMessage As String
Public Property Get DebugMode() As Boolean
    DebugMode = m_DebugMode
End Property
' Initializes the logging system
Public Sub Initialize(SubOrFunctionName As String, DebugMode As Boolean, Optional LogFilePath As String = "\\PRECILAB-SERVER\LabPlusServer\Macros\Log Files\")
    m_DebugMode = DebugMode
    m_ComputerName = Environ("COMPUTERNAME")
    m_LogFilePath = LogFilePath

    ' Initiate log file creation
    Set m_LogFile = InitiateNewLogFile(SubOrFunctionName)
End Sub

' Creates a new log file
Private Function InitiateNewLogFile(SubOrFunctionName As String) As Scripting.TextStream
    Dim FSO As Scripting.FileSystemObject
    Dim LogFile As Scripting.TextStream
    Dim DateStamp As String

    DateStamp = Format(Now, "yyyy-mm-dd hh") & "h" & Format(Now, "nn") & "m" & Format(Now, "ss") & "s"
    Set FSO = New Scripting.FileSystemObject

    ' Open or create the log file
    Set LogFile = FSO.OpenTextFile(m_LogFilePath & m_ComputerName & " - " & SubOrFunctionName & " - " & DateStamp & ".txt", ForWriting, True)
    ' Log initialization message
    Call LogMessage(SubOrFunctionName, LogLevelEnum.LogInfo, "Log initialized for " & m_ComputerName)

    Set InitiateNewLogFile = LogFile
End Function

' Logs a message with the specified level, subroutine name, and message
Public Sub LogMessage(SubOrFunctionName As String, Level As LogLevelEnum, Message As String, Optional MinDisplay As LogLevelEnum = LogLevelEnum.LogInfo, Optional DisplayToUser As Boolean = False)
    Dim FormattedMessage As String
    Dim LogMessageText As String
    Dim FullMessageKey As String

    On Error GoTo ErrorHandler

    FormattedMessage = FormatNewLines(Message)
    LogMessageText = FormatLogEntry(SubOrFunctionName, Level, FormattedMessage)

    ' Create a unique key for this log line
    FullMessageKey = SubOrFunctionName & "|" & Level & "|" & FormattedMessage

    ' Check for repetition
    If FullMessageKey = m_LastMessage Then
        m_RepeatCount = m_RepeatCount + 1

        ' Display it in the debug windwo
        If Level >= MinDisplay And m_DebugMode Then
            Debug.Print Replace(LogMessageText, vbNewLine & Space(64), "")
        End If

        ' Alert the user (if requested)
        If DisplayToUser Then
            MsgBox GetSeverityLabel(Level) & vbNewLine & vbNewLine & FormattedMessage, vbOKOnly, "Log - " & GetSeverityLabel(Level)
        End If

    Else
        ' If this message does not match, but a if the previous message repeated, flush it with count
        If m_RepeatCount > 0 Then
            Call FlushRepeatedMessage
        End If

        ' Log the new message (that doesn't match)
        If m_DebugMode And Level >= MinDisplay Then
            Debug.Print Replace(LogMessageText, vbNewLine & Space(64), "")
        End If

        If Not m_LogFile Is Nothing Then
            m_LogFile.WriteLine Now & " - " & LogMessageText
        End If

        If DisplayToUser Then
            MsgBox GetSeverityLabel(Level) & vbNewLine & vbNewLine & FormattedMessage, vbOKOnly, "Log - " & GetSeverityLabel(Level)
        End If

        ' Update state
        m_LastFormattedMessage = FormattedMessage
        m_LastMessage = FullMessageKey
        m_LastSubName = SubOrFunctionName
        m_LastLevel = Level
        m_RepeatCount = 0
    End If

    Exit Sub

ErrorHandler:
    HandleLogError "Error logging message in " & SubOrFunctionName
End Sub

' Formats a log entry with subroutine name, severity level, and message
Private Function FormatLogEntry(SubOrFunctionName As String, Level As LogLevelEnum, FormattedMessage As String) As String
    FormatLogEntry = Left(SubOrFunctionName & Space(25), 25) & vbTab & GetSeverityLabel(Level) & FormattedMessage
End Function

' Retrieves the severity label for the log level
Private Function GetSeverityLabel(Level As LogLevelEnum) As String
    Select Case Level
        Case LogLevelEnum.LogDebug: GetSeverityLabel = "DEBUG: "
        Case LogLevelEnum.LogInfo: GetSeverityLabel = "INFO: "
        Case LogLevelEnum.LogWarning: GetSeverityLabel = "WARNING: "
        Case LogLevelEnum.LogError: GetSeverityLabel = "ERROR: "
    End Select
End Function

' Handles errors that occur during logging
Private Sub HandleLogError(ErrorMessage As String)
    Debug.Print "Log Error: " & ErrorMessage
    If Not m_LogFile Is Nothing Then
        On Error Resume Next
        m_LogFile.WriteLine Now & " - ERROR: " & ErrorMessage
        On Error GoTo 0
    End If
End Sub
' Formats new lines in the message to align properly and wraps long lines
Private Function FormatNewLines(ByVal InputStr As String) As String
    Const MAX_LINE_LENGTH As Long = 128
    Dim Indent As String
    Dim ResultStr As String
    Dim Lines() As String
    Dim i As Long

    Indent = vbNewLine & Space(64)

    ' Remove leading newline if present
    If Left(InputStr, Len(vbNewLine)) = vbNewLine Then
        InputStr = Mid(InputStr, Len(vbNewLine) + 1)
    End If

    ' Split the input into lines
    Lines = Split(InputStr, vbNewLine)

    ' Loop through lines and apply indent + wrapping
    For i = LBound(Lines) To UBound(Lines)
        Dim Wrapped As String
        Wrapped = WrapLine(Lines(i), MAX_LINE_LENGTH, Indent)

        If i = 0 Then
            ResultStr = Wrapped
        Else
            ResultStr = ResultStr & Indent & Wrapped
        End If
    Next i

    FormatNewLines = ResultStr
End Function

' Helper to wrap a single line to a max length with optional indent for wrapped parts
Private Function WrapLine(ByVal Text As String, ByVal MaxLen As Long, ByVal LineIndent As String) As String
    Dim Result As String
    Dim i As Long
    Dim LineChunk As String

    Do While Len(Text) > MaxLen
        ' Try to break at a space, but if not found, just cut at MaxLen
        i = InStrRev(Left(Text, MaxLen), " ")
        If i = 0 Then i = MaxLen

        LineChunk = Left(Text, i)
        Result = Result & LineChunk & LineIndent

        Text = Mid(Text, i + 1)
    Loop

    Result = Result & Text
    WrapLine = Result
End Function

'' Formats new lines in the message to align properly
'Private Function FormatNewLines(ByVal InputStr As String) As String
'    Dim ResultStr As String
'
'    If Left(InputStr, Len(vbNewLine)) = vbNewLine Then
'        ResultStr = Mid(InputStr, Len(vbNewLine) + 1)
'    Else
'        ResultStr = InputStr
'    End If
'
'    ResultStr = Replace(ResultStr, vbNewLine, vbNewLine & Space(64))
'    FormatNewLines = ResultStr
'End Function
Private Sub FlushRepeatedMessage()
    Dim RepeatText As String

    RepeatText = FormatLogEntry(m_LastSubName, m_LastLevel, FormatNewLines(m_LastFormattedMessage & " (Repeated x" & (m_RepeatCount + 1) & ")"))

    If Not m_LogFile Is Nothing Then
        m_LogFile.WriteLine Now & " - " & RepeatText
    End If
End Sub
' Logs errors in a standard format
Public Sub LogError(SubOrFunctionName As String, ErrorNumber As Long, ErrorDescription As String, Optional PropogateError As Boolean = True, Optional AlertUser As Boolean = False)
    Dim ErrorMessage As String

    ErrorMessage = "Error " & ErrorNumber & ": " & ErrorDescription
    Call LogMessage(SubOrFunctionName, LogLevelEnum.LogError, ErrorMessage, , AlertUser)
    If PropogateError Then Err.Raise ErrorNumber, SubOrFunctionName, ErrorDescription
End Sub
' Closes the log file cleanly
Public Sub CloseLogFile()
    If Not m_LogFile Is Nothing Then

        ' Flush any remaining message
        If m_RepeatCount > 0 Then
            Call FlushRepeatedMessage
        End If

        m_LogFile.Close
        Set m_LogFile = Nothing
    End If
End Sub
Private Sub Class_Terminate()
    Call CloseLogFile
End Sub
