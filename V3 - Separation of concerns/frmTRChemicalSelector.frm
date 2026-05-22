VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTRChemicalSelector 
   Caption         =   "Get Chemical"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4485
   OleObjectBlob   =   "frmTRChemicalSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTRChemicalSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' Name: frmTRChemicalSelector
' Author: JT
' Created: 2026-04-23
' Description: Dumb display form. Presents a choose-existing or
'              enter-new UI for resolving an unmatched chemical
'              string from a testing request form. All validation
'              and service calls are owned by clsTRFormInputResolver.
'              Includes a dual-listbox for KEDS element selection.
' DependsOn: clsLoggingSystem
' ChangeLog:
'   - 1.0.0 - 2026-04-23 - Initial release
'
' ================================================================
' LAYER: UI - no service or DB dependencies.
'
' LIFECYCLE:
'   1. New frmTRChemicalSelector
'   2. .Initialize(AllChemicals, MetalsPrepList, SiliconPrepList,
'                  IonsPrepList, AllElements, UnmatchedName, Logger)
'   3. .Show                        <- blocks until OK or Cancel
'   4. Read .Cancelled              <- True if user dismissed
'      Read .IsNewChemical          <- True if enter-new selected
'      Read .SelectedChemicalID     <- if IsNewChemical = False
'      Read .NewChemicalDict        <- if IsNewChemical = True
'   5. .ShowValidationError(Msg)    <- called by resolver on failure
'      .Show again                  <- resolver re-shows for correction
'
' HIDE CONTRACT:
'   All close paths call Me.Hide, never Unload.
'
' CANCELLED DEFAULT:
'   m_Cancelled initialises True. Only btnOK_Click sets it False.
'   The X button is trapped by QueryClose and routes to Cancel.
'
' DROPDOWN ID CONTRACT:
'   Chemical IDs stored in m_ChemicalIDs(), parallel to cboExisting.
'   Element IDs stored in m_AvailableElementIDs() and
'   m_SelectedElementIDs(), parallel to lstAvailable and lstSelected.
'
' KEDS DUAL-LISTBOX:
'   lstAvailable  - all elements not yet selected
'   lstSelected   - elements chosen for this chemical
'   btnAddElement    (->)  moves item from Available to Selected
'   btnRemoveElement (<-)  moves item from Selected back to Available
'   NewChemicalDict includes ked_elements key -> Dictionary Long->True
' ================================================================

' ----------------------------------------------------------------
' Controls expected on the form:
'   lblPrompt              Label
'   lblValidationError     Label   - hidden by default
'   fraChoose              Frame
'   optChooseExisting      Radio
'   cboExisting            ComboBox
'   fraNew                 Frame
'   optEnterNew            Radio
'   txtChemicalName        TextBox
'   cboMetalsPrep          ComboBox
'   cboSiliconPrep         ComboBox
'   cboIonsPrep            ComboBox
'   lstAvailable           ListBox  - available elements
'   lstSelected            ListBox  - selected KEDS elements
'   btnAddElement          CommandButton  - moves to selected
'   btnRemoveElement       CommandButton  - moves to available
'   btnOK                  CommandButton
'   btnCancel              CommandButton
' ----------------------------------------------------------------

' ----------------------------------------------------------------
' Private state
' ----------------------------------------------------------------
Private m_Logger               As clsLoggingSystem
Private m_Cancelled            As Boolean
Private m_ChemicalIDs()        As Long     ' parallel to cboExisting
Private m_AvailableElementIDs() As Long    ' parallel to lstAvailable
Private m_SelectedElementIDs() As Long     ' parallel to lstSelected
Private m_AvailableCount       As Long
Private m_SelectedCount        As Long

' ----------------------------------------------------------------
' Initialization  (called by resolver before .Show)
' ----------------------------------------------------------------
Public Sub Initialize(AllChemicals As Collection, _
                      MetalsPrepList As Variant, _
                      SiliconPrepList As Variant, _
                      IonsPrepList As Variant, _
                      AllElements As Collection, _
                      UnmatchedName As String, _
                      Logger As clsLoggingSystem)

    Debug.Assert Not Logger Is Nothing
    Debug.Assert Not AllChemicals Is Nothing
    Debug.Assert Not AllElements Is Nothing
    Debug.Assert m_Logger Is Nothing    ' enforce single-init

    Set m_Logger = Logger
    m_Cancelled = True                  ' safe default

    ' Prompt label
    Dim PromptText As String
    PromptText = "Chemical not found: """ & UnmatchedName & """"
    PromptText = PromptText & vbNewLine & _
                 "Select an existing chemical or enter a new one."
    lblPrompt.Caption = PromptText

    ' Validation label hidden by default
    lblValidationError.Caption = ""
    lblValidationError.Visible = False

    ' Populate dropdowns
    PopulateExistingDropdown AllChemicals
    PopulatePrepDropdown cboMetalsPrep, MetalsPrepList
    PopulatePrepDropdown cboSiliconPrep, SiliconPrepList
    PopulatePrepDropdown cboIonsPrep, IonsPrepList
    PopulateAvailableElements AllElements

    ' Pre-fill name with unmatched input
    TxtChemicalName.Text = UnmatchedName

    ' Default to choose-existing
    optChooseExisting.Value = True
    SetChooseExistingMode

End Sub

' ----------------------------------------------------------------
' Public outputs  (read by resolver after .Show returns)
' ----------------------------------------------------------------
Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property

Public Property Get IsNewChemical() As Boolean
    IsNewChemical = optEnterNew.Value
End Property

Public Property Get SelectedChemicalID() As Long
    If cboExisting.ListIndex < 0 Then
        SelectedChemicalID = 0
        Exit Property
    End If
    SelectedChemicalID = m_ChemicalIDs(cboExisting.ListIndex)
End Property

Public Property Get NewChemicalDict() As Scripting.Dictionary
    Dim D As New Scripting.Dictionary
    D.Add "chemical_name", Trim(TxtChemicalName.Text)
    D.Add "metals_prep", cboMetalsPrep.Text
    D.Add "silicon_prep", cboSiliconPrep.Text
    D.Add "ions_prep", cboIonsPrep.Text

    ' Build ked_elements Dictionary Long->True from selected list
    Dim KEDElements As New Scripting.Dictionary
    Dim i As Long
    For i = 0 To m_SelectedCount - 1
        KEDElements.Add m_SelectedElementIDs(i), True
    Next i
    D.Add "ked_elements", KEDElements

    Set NewChemicalDict = D
End Property

' ----------------------------------------------------------------
' Validation feedback  (called by resolver on failure, before re-Show)
' ----------------------------------------------------------------
Public Sub ShowValidationError(Message As String)
    lblValidationError.Caption = Message
    lblValidationError.Visible = True
End Sub

Public Sub ClearValidationError()
    lblValidationError.Caption = ""
    lblValidationError.Visible = False
End Sub

' ----------------------------------------------------------------
' Button events
' ----------------------------------------------------------------
Private Sub btnOK_Click()
    If optChooseExisting.Value Then
        If cboExisting.ListIndex < 0 Then
            ShowValidationError "Please select an existing chemical."
            Exit Sub
        End If
    Else
        If Len(Trim(TxtChemicalName.Text)) = 0 Then
            ShowValidationError "Chemical name is required."
            Exit Sub
        End If
        If cboMetalsPrep.ListIndex < 0 Then
            ShowValidationError "Metals preparation type is required."
            Exit Sub
        End If
        If cboSiliconPrep.ListIndex < 0 Then
            ShowValidationError "Silicon preparation type is required."
            Exit Sub
        End If
        If cboIonsPrep.ListIndex < 0 Then
            ShowValidationError "IC preparation type is required."
            Exit Sub
        End If
    End If

    m_Cancelled = False
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    m_Cancelled = True
    Me.Hide
End Sub

' ----------------------------------------------------------------
' Trap the X button - route to Cancel
' ----------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnCancel_Click
    End If
End Sub

' ----------------------------------------------------------------
' Radio button toggle
' ----------------------------------------------------------------
Private Sub optChooseExisting_Click()
    SetChooseExistingMode
End Sub

Private Sub optEnterNew_Click()
    SetEnterNewMode
End Sub

Private Sub SetChooseExistingMode()
    cboExisting.Enabled = True
    TxtChemicalName.Enabled = False
    cboMetalsPrep.Enabled = False
    cboSiliconPrep.Enabled = False
    cboIonsPrep.Enabled = False
    lstAvailable.Enabled = False
    lstSelected.Enabled = False
    btnAddElement.Enabled = False
    btnRemoveElement.Enabled = False
    ClearValidationError
End Sub

Private Sub SetEnterNewMode()
    cboExisting.Enabled = False
    TxtChemicalName.Enabled = True
    cboMetalsPrep.Enabled = True
    cboSiliconPrep.Enabled = True
    cboIonsPrep.Enabled = True
    lstAvailable.Enabled = True
    lstSelected.Enabled = True
    btnAddElement.Enabled = True
    btnRemoveElement.Enabled = True
    ClearValidationError
End Sub

' ----------------------------------------------------------------
' KEDS dual-listbox events
' ----------------------------------------------------------------
Private Sub btnAddElement_Click()
    Dim idx As Long
    idx = lstAvailable.ListIndex
    If idx < 0 Then Exit Sub

    ' Add to selected list
    lstSelected.AddItem lstAvailable.List(idx)
    m_SelectedCount = m_SelectedCount + 1
    ReDim Preserve m_SelectedElementIDs(0 To m_SelectedCount - 1)
    m_SelectedElementIDs(m_SelectedCount - 1) = m_AvailableElementIDs(idx)

    ' Remove from available list
    RemoveFromAvailable idx
End Sub

Private Sub btnRemoveElement_Click()
    Dim idx As Long
    idx = lstSelected.ListIndex
    If idx < 0 Then Exit Sub

    ' Return to available list
    lstAvailable.AddItem lstSelected.List(idx)
    m_AvailableCount = m_AvailableCount + 1
    ReDim Preserve m_AvailableElementIDs(0 To m_AvailableCount - 1)
    m_AvailableElementIDs(m_AvailableCount - 1) = m_SelectedElementIDs(idx)

    ' Remove from selected list
    RemoveFromSelected idx
End Sub

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub PopulateExistingDropdown(AllChemicals As Collection)
    Dim Count As Long
    Count = AllChemicals.Count

    If Count = 0 Then
        cboExisting.Enabled = False
        optEnterNew.Value = True
        SetEnterNewMode
        Exit Sub
    End If

    ReDim m_ChemicalIDs(0 To Count - 1)

    Dim i As Long
    Dim C As clsChemical
    i = 0

    For Each C In AllChemicals
        cboExisting.AddItem C.Name
        m_ChemicalIDs(i) = C.ID
        i = i + 1
    Next C

    cboExisting.ListIndex = -1
End Sub

Private Sub PopulatePrepDropdown(Cbo As MSForms.ComboBox, PrepList As Variant)
    Dim i As Long
    For i = LBound(PrepList) To UBound(PrepList)
        Cbo.AddItem PrepList(i)
    Next i
    Cbo.ListIndex = 0   ' default to first option (typically N/A)
End Sub

Private Sub PopulateAvailableElements(AllElements As Collection)
    Dim Count As Long
    Count = AllElements.Count

    If Count = 0 Then Exit Sub

    ReDim m_AvailableElementIDs(0 To Count - 1)
    m_AvailableCount = 0

    Dim E As clsElement
    For Each E In AllElements
        lstAvailable.AddItem E.Symbol & " - " & E.Name
        m_AvailableElementIDs(m_AvailableCount) = E.ID
        m_AvailableCount = m_AvailableCount + 1
    Next E

    m_SelectedCount = 0
    ReDim m_SelectedElementIDs(0 To 0)  ' empty - will grow on add
End Sub

Private Sub RemoveFromAvailable(idx As Long)
    ' Rebuild available list and ID array without the removed item
    Dim i As Long
    Dim NewCount As Long
    NewCount = m_AvailableCount - 1

    Dim NewIDs() As Long
    If NewCount > 0 Then
        ReDim NewIDs(0 To NewCount - 1)
    End If

    lstAvailable.RemoveItem idx

    Dim NewIdx As Long
    NewIdx = 0
    For i = 0 To m_AvailableCount - 1
        If i <> idx Then
            NewIDs(NewIdx) = m_AvailableElementIDs(i)
            NewIdx = NewIdx + 1
        End If
    Next i

    m_AvailableCount = NewCount
    If NewCount > 0 Then
        m_AvailableElementIDs = NewIDs
    End If
End Sub

Private Sub RemoveFromSelected(idx As Long)
    ' Rebuild selected list and ID array without the removed item
    Dim i As Long
    Dim NewCount As Long
    NewCount = m_SelectedCount - 1

    Dim NewIDs() As Long
    If NewCount > 0 Then
        ReDim NewIDs(0 To NewCount - 1)
    End If

    lstSelected.RemoveItem idx

    Dim NewIdx As Long
    NewIdx = 0
    For i = 0 To m_SelectedCount - 1
        If i <> idx Then
            NewIDs(NewIdx) = m_SelectedElementIDs(i)
            NewIdx = NewIdx + 1
        End If
    Next i

    m_SelectedCount = NewCount
    If NewCount > 0 Then
        m_SelectedElementIDs = NewIDs
    End If
End Sub


