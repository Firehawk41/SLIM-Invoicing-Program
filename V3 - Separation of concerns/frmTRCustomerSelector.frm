VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTRCustomerSelector 
   Caption         =   "Customer Selector"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   OleObjectBlob   =   "frmTRCustomerSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTRCustomerSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' Name: frmTRCustomerSelector
' Author: JT
' Created: 2026-04-23
' Description: Dumb display form. Presents a choose-existing or
'              enter-new UI for resolving an unmatched customer
'              string from a testing request form. All validation
'              and service calls are owned by clsTRFormInputResolver.
' DependsOn: clsLoggingSystem
' ChangeLog:
'   - 1.0.0 - 2026-04-23 - Initial release
'
' ================================================================
' LAYER: UI - no service or DB dependencies.
'
' LIFECYCLE:
'   1. New frmTRCustomerSelector
'   2. .Initialize(AllCustomers, UnmatchedName, UnmatchedAddress, Logger)
'   3. .Show                        <- blocks until OK or Cancel
'   4. Read .Cancelled              <- True if user dismissed
'      Read .IsNewCustomer          <- True if enter-new selected
'      Read .SelectedCustomerID     <- if IsNewCustomer = False
'      Read .NewCustomerDict        <- if IsNewCustomer = True
'   5. .ShowValidationError(Msg)    <- called by resolver on failure
'      .Show again                  <- resolver re-shows for correction
'
' HIDE CONTRACT:
'   All close paths call Me.Hide, never Unload.
'   The resolver reads properties after Show returns - the form
'   must remain in memory until the resolver is done with it.
'
' CANCELLED DEFAULT:
'   m_Cancelled initialises True. Only btnOK_Click sets it False.
'   The X button is trapped by QueryClose and routes to Cancel.
'
' DROPDOWN ID CONTRACT:
'   Customer IDs are stored in m_CustomerIDs(), a private Long
'   array parallel to cboExisting. cboExisting.ListIndex indexes
'   into both. Display text is "Name - FormattedAddress".
' ================================================================

' ----------------------------------------------------------------
' Controls expected on the form:
'   lblPrompt          Label   - shows unmatched input + instructions
'   lblValidationError Label   - hidden by default, shown on failure
'   fraChoose          Frame   - contains optChooseExisting + cboExisting
'   optChooseExisting  Radio
'   cboExisting        ComboBox
'   Frame1             Frame   - contains optEnterNew + new-record fields
'   optEnterNew        Radio
'   txtName            TextBox - Customer Name
'   txtStreetAddress   TextBox - Street Address
'   txtCity            TextBox - City
'   txtState           TextBox - State
'   txtPostalCode      TextBox - Postal Code
'   txtCountry         TextBox - Country
'   btnOK              CommandButton
'   btnCancel          CommandButton
' ----------------------------------------------------------------

' ----------------------------------------------------------------
' Private state
' ----------------------------------------------------------------
Private m_Logger        As clsLoggingSystem
Private m_Cancelled     As Boolean
Private m_CustomerIDs() As Long     ' parallel to cboExisting list

' ----------------------------------------------------------------
' Initialization  (called by resolver before .Show)
' ----------------------------------------------------------------
Public Sub Initialize(AllCustomers As Collection, _
                      UnmatchedName As String, _
                      UnmatchedAddress As String, _
                      Logger As clsLoggingSystem)

    Debug.Assert Not Logger Is Nothing
    Debug.Assert Not AllCustomers Is Nothing
    Debug.Assert m_Logger Is Nothing    ' enforce single-init

    Set m_Logger = Logger
    m_Cancelled = True                  ' safe default

    ' Prompt label - show what was on the form
    Dim PromptText As String
    PromptText = "Customer not found: """ & UnmatchedName & """"
    If Len(Trim(UnmatchedAddress)) > 0 Then
        PromptText = PromptText & " / """ & UnmatchedAddress & """"
    End If
    PromptText = PromptText & vbNewLine & _
                 "Select an existing customer or enter a new one."
    lblPrompt.Caption = PromptText

    ' Validation label hidden by default
    lblValidationError.Caption = ""
    lblValidationError.Visible = False

    ' Populate dropdown
    PopulateExistingDropdown AllCustomers

    ' Pre-fill new-record fields with unmatched input
    txtName.Text = UnmatchedName
    txtStreetAddress.Text = UnmatchedAddress

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

Public Property Get IsNewCustomer() As Boolean
    IsNewCustomer = optEnterNew.Value
End Property

Public Property Get SelectedCustomerID() As Long
    If cboExisting.ListIndex < 0 Then
        SelectedCustomerID = 0
        Exit Property
    End If
    SelectedCustomerID = m_CustomerIDs(cboExisting.ListIndex)
End Property

Public Property Get NewCustomerDict() As Scripting.Dictionary
    Dim D As New Scripting.Dictionary
    D.Add "customer_name", Trim(txtName.Text)
    D.Add "street_address", Trim(txtStreetAddress.Text)
    D.Add "city", Trim(txtCity.Text)
    D.Add "state", Trim(txtState.Text)
    D.Add "postal_code", Trim(txtPostalCode.Text)
    D.Add "country", Trim(txtCountry.Text)
    Set NewCustomerDict = D
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
    ' Guard - must have a valid selection before allowing OK
    If optChooseExisting.Value Then
        If cboExisting.ListIndex < 0 Then
            ShowValidationError "Please select an existing customer."
            Exit Sub
        End If
    Else
        If Len(Trim(txtName.Text)) = 0 Then
            ShowValidationError "Customer name is required."
            Exit Sub
        End If
    End If

    m_Cancelled = False
    Me.Hide     ' NOT Unload - resolver reads properties after Show returns
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
    txtName.Enabled = False
    txtStreetAddress.Enabled = False
    txtCity.Enabled = False
    txtPostalCode.Enabled = False
    txtCountry.Enabled = False
    ClearValidationError
End Sub

Private Sub SetEnterNewMode()
    cboExisting.Enabled = False
    txtName.Enabled = True
    txtStreetAddress.Enabled = True
    txtCity.Enabled = True
    txtPostalCode.Enabled = True
    txtCountry.Enabled = True
    ClearValidationError
End Sub

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub PopulateExistingDropdown(AllCustomers As Collection)
    Dim Count As Long
    Count = AllCustomers.Count

    If Count = 0 Then
        cboExisting.Enabled = False
        optEnterNew.Value = True
        SetEnterNewMode
        Exit Sub
    End If

    ReDim m_CustomerIDs(0 To Count - 1)

    Dim i As Long
    Dim C As clsCustomer
    i = 0

    For Each C In AllCustomers
        cboExisting.AddItem C.Name & " - " & C.FormattedAddress
        m_CustomerIDs(i) = C.ID
        i = i + 1
    Next C

    cboExisting.ListIndex = -1  ' no default selection - user must choose
End Sub


