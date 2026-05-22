VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestingRequestEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsTestingRequestEmail
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-11
' Description: Stores properties relating to an individual email
' DependsOn: clsLoggingSystem
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Private m_Logger As clsLoggingSystem
Private m_EmailID As Long
Private m_SubmissionID As Long
Private m_EmailAddress As String

Public Property Get EmailID() As Long
    EmailID = m_EmailID
End Property

Public Property Let EmailID(ByVal Value As Long)
    m_EmailID = Value
End Property

Public Property Get SubmissionID() As Long
    SubmissionID = m_SubmissionID
End Property

Public Property Let SubmissionID(ByVal Value As Long)
    m_SubmissionID = Value
End Property

Public Property Get EmailAddress() As String
    EmailAddress = m_EmailAddress
End Property


Public Sub Initialize(EmailAddress As String, Logger As clsLoggingSystem)
    Set m_Logger = Logger
    m_EmailAddress = EmailAddress

End Sub
' Validate email format (e.g., ensure it contains "@")
Public Function IsValidEmail() As Boolean
    If InStr(1, m_EmailAddress, "@") Then
        IsValidEmail = True
    Else
        IsValidEmail = False
    End If
End Function
