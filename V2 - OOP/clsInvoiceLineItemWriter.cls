VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceLineItemWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceLineItemWriter
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-14
' Description: Object in charge of writing line items to the Word Invoice
' DependsOn: clsLoggingSystem
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Private m_Logger As clsLoggingSystem
Public Sub Initialize(Logger As clsLoggingSystem)
    Call Logger.LogMessage("Initialize", LogLevelEnum.LogDebug, "Initializing Line Item Writer")
    Set m_Logger = Logger
End Sub

Public Sub WriteLineItem(Table As Word.Table, LineNumber As Long, TextColour As Long, ColumnMap As Scripting.Dictionary)
    Dim Cell As Word.Cell
    Dim Key As Variant
    Dim Value As Variant
    Dim NotBoldRows As Variant
    
    Call m_Logger.LogMessage("WriteLineItem", LogLevelEnum.LogDebug, "Initiating filling line item in row " & LineNumber)
    
    For Each Key In ColumnMap.Keys
        ' Ensure column exists in the table
        If Key <= Table.Columns.Count Then
            Set Cell = Table.Cell(LineNumber, Key)
            Value = ColumnMap(Key)
            
            Call ProcessCell(Cell, Value, TextColour)
        End If
    Next Key
    
    Call m_Logger.LogMessage("WriteLineItem", LogLevelEnum.LogInfo, "Successfully filled line item in row " & LineNumber)
    
End Sub
Private Sub ProcessCell(Cell As Word.Cell, CellValue As Variant, TextColour As Long)

    Call m_Logger.LogMessage("ProcessColumn", LogLevelEnum.LogDebug, "Initiating processing cell on row " & Cell.RowIndex & " and column " & Cell.ColumnIndex)
    
    If IsArray(CellValue) Then

        Call ProcessSplitCell(Cell, CellValue, TextColour)
    Else
        Call ProcessRegularCell(Cell, CellValue, TextColour)
    End If

    Call m_Logger.LogMessage("ProcessColumn", LogLevelEnum.LogInfo, "Finished processing cell on row " & Cell.RowIndex & " and column " & Cell.ColumnIndex)
End Sub
Private Sub ProcessSplitCell(Cell As Word.Cell, CellValue As Variant, TextColour As Long)
    Dim RowIndex As Long
    Dim NewCell As Word.Cell
    Dim FormattedText As clsInvoiceLineItemFormattedText
    
    ' Split the cell
    If UBound(CellValue) > 0 Then
        Cell.Split NumRows:=UBound(CellValue) + 1, NumColumns:=1
    End If

    ' Populate the split cells
    For RowIndex = 0 To UBound(CellValue)
        Set NewCell = Cell.Parent.Cell(Cell.RowIndex + RowIndex, Cell.ColumnIndex)
        Set FormattedText = CellValue(RowIndex)
        
        ' Don't write if the contents are '0'
        If Not FormattedText Is Nothing Then
            NewCell.Range.Text = FormattedText.Text
            Call SetCellFontColor(NewCell, TextColour, FormattedText.IsBold, FormattedText.ColourDict)
        End If
        
        ' Apply consistent formatting
        Call SetCellBorders(NewCell, False, False)
        
        
    Next RowIndex
    
    ' Ensure the last split cell has a bottom border
    Call SetCellBorders(NewCell, False, True)
    Call SetCellBorders(Cell, True, False)
    
End Sub

Private Sub ProcessRegularCell(Cell As Word.Cell, CellValue As Variant, TextColour As Long)
    Dim FormattedText As clsInvoiceLineItemFormattedText
    
    Set FormattedText = CellValue
    ' Fill in and format non-split cell
    Cell.Range.Text = FormattedText.Text
    Call SetCellBorders(Cell, True, True)
    Call SetCellFontColor(Cell, TextColour, FormattedText.IsBold)

End Sub
Private Sub SetCellBorders(Cell As Word.Cell, IsTopBorder As Boolean, IsBottomBorder As Boolean)

    ' Set the top border
    If IsTopBorder Then
        Cell.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
    Else
        Cell.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    End If

    ' Set the bottom border
    If IsBottomBorder Then
        Cell.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
    End If

End Sub
Private Sub SetCellFontColor(Cell As Word.Cell, TextColour As Long, Optional IsBold As Boolean = True, Optional ColourDict As Scripting.Dictionary = Nothing)
    Const SAME_DAY_RUSH As String = "Same Day RUSH"
    Const CALL_IN_RUSH As String = "Call-in RUSH"
    Dim Key As Variant
    Dim StartingPos As Long
    Dim EndingPos As Long
    Dim Parts() As String
    
    ' Set the text colour but overrule if the sample is rush
    If CleanCellText(Cell.Range.Text) = SAME_DAY_RUSH Or CleanCellText(Cell.Range.Text) = CALL_IN_RUSH Then
        Cell.Range.Font.Color = vbRed
    Else
        Cell.Range.Font.Color = TextColour
    End If
    
    If Not ColourDict Is Nothing Then
        For Each Key In ColourDict.Keys
            Parts = Split(Key, "-")
            StartingPos = CLng(Parts(0))
            EndingPos = CLng(Parts(1))
            
            With Cell.Range
                .Start = Cell.Range.Start + StartingPos - 1
                .End = Cell.Range.Start + EndingPos
                .Font.Color = ColourDict(Key)
            End With
        Next Key
    End If
    
    ' Set the bold state
    Cell.Range.Font.Bold = IsBold

    Call m_Logger.LogMessage("SetCellFontColor", LogLevelEnum.LogDebug, "Set cell colour")

End Sub
Private Function CleanCellText(ByVal Txt As String) As String
    CleanCellText = Trim(Replace(Replace(Replace(Txt, vbCr, ""), vbNewLine, ""), Chr(7), ""))
End Function

