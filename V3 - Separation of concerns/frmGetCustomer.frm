VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGetCustomer 
   Caption         =   "Get Customer"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   OleObjectBlob   =   "frmGetCustomer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGetCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
' MacroName: frmGetCustomer
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-05
' Description: Allows the user to select a chemical or enter a new one
' DependsOn: clsLoggingSystem, clsSLIMDictionaryManager, clsTestingRequestElement, modEnums, clsTestingRequestChemical
' ChangeLog:
'   - 1.0.1 - 2025-06-05 - Added metadata
Private m_Cancelled As Boolean
Private m_Logger As clsLoggingSystem
Private m_CustomerList As Variant
Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property
Public Property Get CustomerName() As String
    If btnExistingCustomer.Value Then
        CustomerName = cboCustomerName.Text
    Else
        CustomerName = txtCustomerName.Text
    End If
End Property
Public Property Get StreetAddress() As String
    If btnNewCustomer.Value Then StreetAddress = Me.txtStreetAddress.Value
End Property
Public Property Get City() As String
    If btnNewCustomer.Value Then City = Me.txtCity.Value
End Property
Public Property Get State() As String
    If btnNewCustomer.Value Then State = Me.txtState.Value
End Property
Public Property Get PostalCode() As String
    If btnNewCustomer.Value Then PostalCode = Me.txtPostalCode.Value
End Property
Public Property Get Country() As String
    If btnNewCustomer.Value Then Country = Me.txtCountry.Value
End Property
Public Sub Initialize(FormCustomerName As String, FormCustomerAddress As String, CustomerDict As Scripting.Dictionary, Logger As clsLoggingSystem)

    btnExistingCustomer = True

    ' Sets up variables for userform
    Set m_Logger = Logger
    m_CustomerList = SetCustomerList(CustomerDict)
    Me.cboCustomerName.List = m_CustomerList
    Me.lblCaption.Caption = "This customer name (" & FormCustomerName & ") and address (" & FormCustomerAddress & ") is not found in the database." & vbNewLine & vbNewLine & "Please select a current customer or enter a new one."
    Me.txtCustomerName = FormCustomerName

End Sub
Private Sub btnExistingCustomer_Click()
    Call SetCustomerControlsState(True)
End Sub

Private Sub btnNewCustomer_Click()
    Call SetCustomerControlsState(False)
End Sub
Private Sub SetCustomerControlsState(IsExisting As Boolean)
    ' Sets the format

    cboCustomerName.Enabled = IsExisting
    cboCustomerName.BackColor = IIf(IsExisting, vbWhite, &H8000000F)

    txtCustomerName.Enabled = Not IsExisting
    txtCustomerName.BackColor = IIf(Not IsExisting, vbWhite, &H8000000F)

End Sub
Private Sub CancelButton_Click()
    m_Cancelled = True
    Me.Hide

    Call m_Logger.LogMessage("GetCustomerForm", LogLevelEnum.LogInfo, "User cancelled the userform.")

End Sub

Private Sub OKButton_Click()

    If ValidateInputs() Then
        Me.Hide
    Else
        Call m_Logger.LogMessage("GetCustomerForm", LogLevelEnum.LogWarning, "Validation failed. Please correct the highlighted errors.", , True)
    End If
End Sub

Private Function ValidateInputs() As Boolean

    Dim iIndex As Long

    ' Check if the existing chemical is recognized
    If btnExistingCustomer.Value Then

        ' Initialize validation flag
        ValidateInputs = False

        ' Loop through the chemical list
        For iIndex = LBound(m_CustomerList) To UBound(m_CustomerList)
            If m_CustomerList(iIndex) = cboCustomerName.Text Then
                ' if a match is found, set the input to valid
                ValidateInputs = True
                Exit For
            End If
        Next iIndex

        ' Colour the combobox based on whether the input was valid
        If ValidateInputs Then
            cboCustomerName.BackColor = vbWhite
        Else
            cboCustomerName.BackColor = vbRed
        End If

    Else

        ' Checks whether the PRECILAB Chemical Box has a valid input
        If txtCustomerName.Text = "" Then
            txtCustomerName.BackColor = vbRed
            ValidateInputs = False
        Else
            txtCustomerName.BackColor = vbWhite
            ValidateInputs = True
        End If


    End If

End Function

Private Function SetCustomerList(CustomerDict As Scripting.Dictionary) As Variant

    Dim Key As Variant
    Dim CustomerList() As Variant
    Dim DictIndex As Long

    ReDim CustomerList(CustomerDict.Count - 1)
    For Each Key In CustomerDict
        CustomerList(DictIndex) = Key
        DictIndex = DictIndex + 1
    Next Key

    ' Return the list to the calling sub
    SetCustomerList = CustomerList
End Function
