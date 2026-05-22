VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceLineItemFormattedText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceLineItemFormattedText
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-11
' Description: Object to store formatted text for line items
' DependsOn: clsLoggingSystem
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Private m_Logger As clsLoggingSystem
Private m_Text As String
Private m_IsBold As Boolean
Private m_ColourDict As Scripting.Dictionary
Public Property Get Text() As String
    Text = m_Text
End Property
Public Property Get IsBold() As Boolean
    IsBold = m_IsBold
End Property
Public Property Get ColourDict() As Scripting.Dictionary
    Set ColourDict = m_ColourDict
End Property
Public Sub Initialize(Text As String, Logger As clsLoggingSystem, Optional IsBold As Boolean = True, Optional ColourDict As Scripting.Dictionary)
    Set m_Logger = Logger
    m_Text = Text
    m_IsBold = IsBold
    If Not ColourDict Is Nothing Then
        Set m_ColourDict = ColourDict
    End If
End Sub
Public Sub AddColourDictEntry(Colour As Long, Optional StartingCharacter As Long = 1, Optional FinishingCharacter As Long = -1)
    Dim Key As String
    
    ' Check that the starting and finishing characters are valid
    If (StartingCharacter > FinishingCharacter And Not FinishingCharacter = -1) Or StartingCharacter < 1 Or FinishingCharacter < -1 Then
        Call m_Logger.LogMessage("InitializeColourDict", LogLevelEnum.LogError, "Invalid starting and/or finishing character!")
        Exit Sub
    End If
    
    ' Replace placeholder value with the number of characters in the text
    If FinishingCharacter = -1 Then FinishingCharacter = Len(m_Text)
    
    ' Set the character range
    Key = CStr(StartingCharacter) & "-" & CStr(FinishingCharacter)
    
    ' Add the range of characters due to be coloured to the dictionary
    If Not m_ColourDict.Exists(Key) Then
        m_ColourDict.Add Key, Colour
    End If
    
End Sub

Private Sub Class_Initialize()
    Set m_ColourDict = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    Set m_ColourDict = Nothing
End Sub
