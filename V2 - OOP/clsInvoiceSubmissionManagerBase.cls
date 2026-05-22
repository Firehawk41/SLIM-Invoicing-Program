VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceSubmissionManagerBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceSubmissionManagerBase
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-05
' Description: Object to hold functions common to all Submission Managers
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsSLIMDictionaryManager, clsTestingRequestSubmission
' ChangeLog:
'   - 1.0.1 - 2025-06-05 - Added metadata
Private m_Logger As clsLoggingSystem
Private m_SLIMAccessDB As clsAccessDatabase
Private m_SLIMDictManager As clsSLIMDictionaryManager
Public Property Get Submissions() As Collection
    Set Submissions = Nothing
End Property
Public Sub Initialize(StartDate As Date, SelectedFilePath As String, Logger As clsLoggingSystem)
    Dim MethodName As String
    Const SLIM_DATABASE_PATH As String = "\\PRECILAB-SERVER\LabPlusServer\Documents_In_Works\Thomson\Sample Login.accdb"
    
    Set m_Logger = Logger
    Call m_SLIMAccessDB.Initialize(SLIM_DATABASE_PATH, Logger)
            
End Sub
Public Function SetIndividualCollection(SelectedFilePath As String) As Collection
    Dim Submission As clsTestingRequestSubmission
    Dim Submissions As New Collection

    Set Submission = SetSubmission(SelectedFilePath)
    Submissions.Add Submission
    
    ' Return the collection to the callings sub
    Set SetIndividualCollection = Submissions
    
End Function

Public Function SetSubmissionCollection(StartDate As Date, EndDate As Date, FilePath As String) As Collection
    
    Dim FileName As String
    Dim FileDate As Date
    Dim WBk As Workbook
    Dim TestingRequest As Worksheet
    Dim Submission As clsTestingRequestSubmission
    Dim SubmissionCollection As New Collection

    
    ' Log the function
    Call m_Logger.LogMessage("SetSubmissionCollection", LogLevelEnum.LogDebug, "Initiating function.")
    
    ' Set the dictionary manager so that it doesn't have to be created for every submission in the collection
    Set m_SLIMDictManager = New clsSLIMDictionaryManager
    Call m_SLIMDictManager.Initialize(m_Logger, m_SLIMAccessDB)
    
    'Sets the file name as every xlsx file in the specified folder
    FileName = Dir(FilePath & "*.xlsx")
    
    'Loops through each file in the specified folder
    Do While FileName <> ""
        ' Check whether the first 6 digits are numeric
        If IsNumeric(Left(FileName, 6)) Then
            FileDate = DateSerial(Mid(FileName, 5, 2), Left(FileName, 2), Mid(FileName, 3, 2))
            If FileDate >= StartDate And FileDate <= EndDate Then
                
                ' Set the submission from the filepath
                Set Submission = SetSubmission(FilePath & "\" & FileName)

                If Not Submission Is Nothing Then
                    SubmissionCollection.Add Submission
                End If

            End If
        End If
        'Proceed to the next file name
        FileName = Dir()
    Loop
    
    Set SetSubmissionCollection = SubmissionCollection
    
End Function
Private Function SetSubmission(FilePath As String) As clsTestingRequestSubmission
    Dim TestingRequest As Worksheet
    Dim Submission As clsTestingRequestSubmission
    
    ' Set the Testing Request WkS
    Set TestingRequest = FindTestingRequest(FilePath)
                
    If Not TestingRequest Is Nothing Then
        
        Set Submission = New clsTestingRequestSubmission
        Call Submission.Initialize(TestingRequest, m_Logger, m_SLIMAccessDB, m_SLIMDictManager)
        
        TestingRequest.Parent.Close SaveChanges:=False
        
        ' Return the submission to the calling sub
        Set SetSubmission = Submission
        
    End If
    
End Function
Private Function FindTestingRequest(File As String) As Worksheet

    Dim WBk As Workbook
    Dim WkS As Worksheet

    Set WBk = Workbooks.Open(File, ReadOnly:=True)
    
    ' Log the function
    Call m_Logger.LogMessage("FindTestingRequest", LogLevelEnum.LogDebug, "Finding testing request for: " & WBk.Name)
    
    For Each WkS In WBk.Sheets
        If WkS.Range("D3").Value Like "*Testing Request Form" Then
            Set FindTestingRequest = WkS
            Call m_Logger.LogMessage("FindTestingRequest", LogLevelEnum.LogInfo, "Testing request found.")
            Exit Function
        End If
    Next WkS
    
    ' Log failure
    Call m_Logger.LogMessage("FindTestingRequest", LogLevelEnum.LogWarning, "No testing request found. Closing workbook.")
    
    ' Close the workbook
    WBk.Close SaveChanges:=False
    
End Function

Private Sub Class_Initialize()
    Set m_SLIMAccessDB = New clsAccessDatabase
End Sub
Private Sub Class_Terminate()
    Set m_SLIMAccessDB = Nothing
End Sub
