VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceContact
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-11
' Description: Container for contacts for invoices
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsTestingRequestSubmission
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Private m_Name As String
Private m_Title As String
Private m_Email As String
Private m_Phone As String
Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Public Property Get FormattedContact() As String
    FormattedContact = Name & IIf(Title = "", "", vbNewLine & Title) & IIf(Email = "", "", vbNewLine & Email) & IIf(Phone = "", "", vbNewLine & Phone)
End Property
Public Property Get Name() As String
    Name = m_Name
End Property
Public Property Get Title() As String
    Title = m_Title
End Property
Public Property Get Email() As String
    Email = m_Email
End Property
Public Property Get Phone() As String
    Phone = m_Phone
End Property

Public Sub InitializeFromDatabase(CustomerName As String, Field1 As String, Field2 As String, Field3 As String, Field4 As String, AccessDB As clsAccessDatabase, Logger As clsLoggingSystem)
    Dim DataArr As Variant
    Dim SQLQuery As String
    
    Set m_Logger = Logger
    Set m_AccessDB = AccessDB

    'References database
    SQLQuery = "SELECT [" & Field1 & "], [" & Field2 & "], [" & Field3 & "], [" & Field4 & "] FROM Customer_Database WHERE [Customer_Name] = '" & CustomerName & "'"
    
    DataArr = AccessDB.ExecuteQuery(SQLQuery)
    
    On Error Resume Next
    m_Name = DataArr(0, 0)
    m_Title = DataArr(1, 0)
    m_Email = DataArr(2, 0)
    m_Phone = DataArr(3, 0)
    On Error GoTo 0

End Sub
Public Sub InitializeFromSubmission(Submission As clsTestingRequestSubmission, Logger As clsLoggingSystem)

    Set m_Logger = Logger
    
    m_Name = Submission.CustomerContact
    m_Email = FindContactEmail(Submission)
End Sub
Private Function FindContactEmail(Submission As clsTestingRequestSubmission) As String
    
    Dim Email As Variant
    Dim ContactNameArr() As String
    Dim ContactLastName As String
    Dim ContactEmail As String
    
    Call m_Logger.LogMessage("FindContactEmail", LogLevelEnum.LogDebug, "Searching for the contact's email")
    
    If Not Submission.CustomerContact = "" Then
    
        ContactNameArr = Split(Submission.CustomerContact, " ")
        ContactLastName = ContactNameArr(UBound(ContactNameArr))
        
        ' Search through each email in the main distribution list for the contact's email
        For Each Email In Submission.ResultsEmailMain
            If InStr(1, Email.EmailAddress, ContactLastName, vbTextCompare) > 0 Then
                ContactEmail = Email.EmailAddress
                Exit For
            End If
        Next Email
        
        ' If an email wasn't found, search through the CC'd distribution list
        If ContactEmail = "" Then
            For Each Email In Submission.ResultsEmailCC
                If InStr(1, Email.EmailAddress, ContactLastName, vbTextCompare) > 0 Then
                    ContactEmail = Email.EmailAddress
                    Exit For
                End If
            Next Email
        End If
        
        ' If the email still wasn't found, search through some hard-coded names and emails (for customers who routinely don't include their email)
        If ContactEmail = "" Then
            Select Case Submission.CustomerContact
                Case "Blake Saffer"
                    ContactEmail = "basaffer@monumentchemical.com"
                Case "Alexis Falgoust"
                    ContactEmail = "aneschete@customer-b.com"
            End Select
        End If
        
        FindContactEmail = ContactEmail
    
    End If
    
    Call m_Logger.LogMessage("FindContactEmail", LogLevelEnum.LogInfo, IIf(FindContactEmail = "", "Could not find contact email", "Contact's email is: " & FindContactEmail))
    
End Function
