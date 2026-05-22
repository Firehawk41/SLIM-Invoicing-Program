VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStalenessChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsStalenessChecker
' Version: 1.0.0
' Author: JT
' Created: 2025-05-16
' LastModified: 2025-05-21
' Description: Object to check whether a workbook is stale (older than the version on the server)
' DependsOn: clsLoggingSystem
' ChangeLog:
'       - 1.0.0 - 2025-05-21 - Initial release.
Private m_FileName As String
Private m_FilePath As String
Private m_LocalDate As Date
Private m_FileDate As Date
Private m_ToleranceMinutes As Long
Private m_Logger As clsLoggingSystem
Public Property Get ServerTimeStamp() As Date
    ServerTimeStamp = m_FileDate
End Property
Public Property Get IsObsoleteMessage() As String
    IsObsoleteMessage = "The version of '" & m_FileName & "' that's open is obsolete.  Please close all instances of Excel to load the current version."
End Property

Public Sub Initialize(ByVal FileName As String, FilePath As String, ByVal LocalDate As Date, Optional ByRef Logger As clsLoggingSystem = Nothing, Optional ByVal ToleranceMinutes As Long = 5)
    m_FileName = FileName
    m_FilePath = FilePath
    m_LocalDate = LocalDate
    m_ToleranceMinutes = ToleranceMinutes
    Set m_Logger = Logger
End Sub
' Returns True if the server file is within the allowed time tolerance
Public Function IsCurrent() As Boolean
    Dim FileDate As Date

    ' Set default return
    IsCurrent = False

    ' Exit if file path is invalid
    If Len(Trim(m_FilePath)) = 0 Then
        If Not m_Logger Is Nothing Then Call m_Logger.LogMessage("MacroValidator.IsCurrent", LogLevelEnum.LogError, "No file path was passed to the object")
        Exit Function
    End If

    ' Checks if the file exists
    If Dir(m_FilePath) = "" Then
        If Not m_Logger Is Nothing Then Call m_Logger.LogMessage("MacroValidator.IsCurrent", LogLevelEnum.LogError, "The specified file doesn't exist")
        Exit Function
    End If


    ' Compare the last modified date with the current date
    m_FileDate = FileDateTime(m_FilePath)
    IsCurrent = (DateDiff("n", m_LocalDate, m_FileDate) <= m_ToleranceMinutes)

    If Not m_Logger Is Nothing Then
        If IsCurrent Then
            Call m_Logger.LogMessage("MacroValidator.IsCurrent", LogLevelEnum.LogDebug, "The specified file is current")
        Else
            Call m_Logger.LogMessage("MacroValidator.IsCurrent", LogLevelEnum.LogError, Me.IsObsoleteMessage)
        End If
    End If


End Function
