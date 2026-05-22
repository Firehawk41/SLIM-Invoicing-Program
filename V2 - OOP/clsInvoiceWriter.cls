VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceWriter
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-16
' Description: Object for writing invoices to word document
' DependsOn: clsLoggingSystem, clsInvoiceSampleRowBase, clsTestingRequestSubmission, clsTestingRequestSample
' ChangeLog:
'   - 1.0.1 - 2025-06-16 - Added metadata
Private m_Logger As clsLoggingSystem
Public Sub Initialize(Logger As clsLoggingSystem)
    Set m_Logger = Logger
End Sub
' Write to a specific bookmark
Public Sub WriteToBookmark(WordDoc As Word.Document, BookmarkName As String, Value As String, Optional TextSize As Long = 12, Optional TextColour As Long = 0, Optional IsBold As Boolean = False, Optional TextFont As String = "Tahoma")

    Dim BookMarkRange As Word.Range
    
    If WordDoc.Bookmarks.Exists(BookmarkName) Then
        Set BookMarkRange = WordDoc.Bookmarks(BookmarkName).Range
        BookMarkRange.Text = Value

        ' Apply formatting
        With BookMarkRange.Font
            .Name = TextFont
            .Color = TextColour
            .Size = TextSize
            .Bold = IsBold
        End With

        ' Re-add the bookmark (since writing text removes it)
        WordDoc.Bookmarks.Add BookmarkName, BookMarkRange

        m_Logger.LogMessage "WriteToBookmark", LogLevelEnum.LogInfo, "Updated bookmark " & BookmarkName

    Else
        m_Logger.LogMessage "WriteToBookmark", LogLevelEnum.LogWarning, "Bookmark " & BookmarkName & " not found"
    End If
End Sub

' Write data to a Word table
Public Sub WriteToTable(WordTable As Word.Table, DataDict As Scripting.Dictionary)
    Dim Key As Variant
    Dim CellCoords() As String
    Dim Row As Integer, Col As Integer
    Dim Cell As Word.Cell

    m_Logger.LogMessage "WriteToTable", LogLevelEnum.LogDebug, "Writing data to table"

    ' Loop through each key-value pair in the dictionary
    For Each Key In DataDict.Keys
        CellCoords = Split(Key, ",") ' Split "Row,Column" into two numbers
        Row = CInt(CellCoords(0))
        Col = CInt(CellCoords(1))
        
        ' Ensure the row exists in the table
        If Row <= WordTable.Rows.Count And Col <= WordTable.Columns.Count Then
            Set Cell = WordTable.Cell(Row, Col)
            Cell.Range.Text = DataDict(Key)
        Else
            m_Logger.LogMessage "WriteToTable", LogLevelEnum.LogWarning, "Cell (" & Row & ", " & Col & ") is out of table range."
        End If
    Next Key

    ' Log completion
    m_Logger.LogMessage "WriteToTable", LogLevelEnum.LogInfo, "Table successfully updated"
End Sub

