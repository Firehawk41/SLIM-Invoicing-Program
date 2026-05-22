VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceReportGatekeeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceReportGatekeeper
' Author: JT
' Created: 2026-02-13
' Description: Stateless minimal pre-flight validation
' ChangeLog:
'   - 0.1.0 - 2026-02-12 - Initial prototype

' ==========
' Public API
' ==========
Public Function ValidateBeforeInvoice(FilePath As String, ErrorMessage As String) As Boolean
    Dim FileName As String
    Dim CleanPath As String

    Debug.Assert FilePath <> ""

    If FilePath = "" Then
        ErrorMessage = "No source workbook provided."
        Exit Function
    End If

    ' Remove trailing slash if present
    CleanPath = FilePath
    If Right$(CleanPath, 1) = "\" Then
        CleanPath = Left$(CleanPath, Len(CleanPath) - 1)
    End If

    ' Extract filename
    FileName = Mid$(CleanPath, InStrRev(CleanPath, "\") + 1)

    If Not ValidateFileType(FileName, ErrorMessage) Then Exit Function
    If Not ValidateFilenameStructure(FileName, ErrorMessage) Then Exit Function


    ValidateBeforeInvoice = True

End Function

' ===============
' Private Helpers
' ===============
Private Function ValidateFileType(FileName As String, ErrorMessage As String) As Boolean
    Dim Extension As String

    Extension = LCase$(Mid$(FileName, InStrRev(FileName, ".") + 1))

    If Extension <> "xlsx" Then
        ErrorMessage = "Invalid file type. Only Excel testing request files allowed."
        Exit Function
    End If

    ValidateFileType = True

End Function

Private Function ValidateFilenameStructure(FileName As String, ErrorMessage As String) As Boolean


    If InStr(1, FileName, "Results") = 0 Then
        ErrorMessage = "Filename does not contain 'Results'."
        Exit Function
    End If

    If InStr(1, FileName, "Partial") > 0 Then
        ErrorMessage = "File appears to be partial or draft. Cannot invoice."
        Exit Function
    End If

    If Not IsNumeric(Left(FileName, 6)) Then
        ErrorMessage = "File name does not begin with date."
        Exit Function
    End If

    ValidateFilenameStructure = True

End Function
