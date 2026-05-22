VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTRFormInputResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ================================================================
' Name: clsTRFormInputResolver
' Author: JT
' Created: 2026-04-23
' Description: Resolves unmatched form input strings for customer
'              and chemical domains. Launches selector UI on cache
'              miss and returns a fully hydrated domain object.
'              Single coordination point for both parsers and any
'              future consumers.
' DependsOn: clsCustomerService, clsChemicalService, clsElementService,
'            frmTRCustomerSelector, frmTRChemicalSelector,
'            clsLoggingSystem
' ChangeLog:
'   - 1.0.0 - 2026-04-23 - Initial release
'
' ================================================================
' LAYER: Application — coordinates domain services and UI.
'        Domain services stay pure; all UI coupling lives here.
'
' LIFECYCLE:
'   1. New clsTRFormInputResolver
'   2. .Initialize(CustomerSvc, ChemicalSvc, ElementSvc, Logger)
'   3a. .ResolveCustomer(Name, Address, Address2) As clsCustomer
'   3b. .ResolveChemical(Name)          As clsChemical
'
' RESOLUTION CONTRACT:
'   Fast path — form input matched in cache, returns object
'               immediately with no UI interaction.
'   Slow path — cache miss, launches selector form. User may
'               select an existing record or enter a new one.
'               New records are validated for name uniqueness
'               before creation. Returns Nothing if cancelled.
'
' RETURN CONTRACT:
'   Returns full domain object in all non-cancelled cases.
'   Caller reads .ID if only the ID is needed.
'   Returns Nothing if user cancels — caller decides whether
'   to skip or abort depending on context (batch vs individual).
'   Current caller treats this as fatal.
'
' INJECTED INTO:
'   clsTRSubmissionFormParser  — resolves customer
'   clsTRSampleFormParser      — resolves chemical
'
' PERSISTENCE:
'   After UI resolution (both select-existing and create-new paths),
'   the form name mapping is written to the DB via the service.
'   Next session the same string resolves via fast path with no UI.
' ================================================================
Private m_Logger As clsLoggingSystem
Private m_CustomerSvc As clsCustomerService
Private m_ChemicalSvc As clsChemicalService
Private m_ElementSvc As clsElementService

Public Sub Initialize(CustomerSvc As clsCustomerService, ChemicalSvc As clsChemicalService, ElementSvc As clsElementService, Logger As clsLoggingSystem)
    Debug.Assert Not Logger Is Nothing
    Debug.Assert Not CustomerSvc Is Nothing
    Debug.Assert Not ChemicalSvc Is Nothing
    Debug.Assert Not ElementSvc Is Nothing
    Debug.Assert m_Logger Is Nothing ' enforce single-init

    If Not m_Logger Is Nothing Then Err.Raise vbObjectError + 1, , "Resolver is already initialized"

    Set m_Logger = Logger
    Set m_CustomerSvc = CustomerSvc
    Set m_ChemicalSvc = ChemicalSvc
    Set m_ElementSvc = ElementSvc

End Sub
Public Function ResolveCustomer(Name As String, Address As String, Address2 As String) As clsCustomer
    AssertInit "ResolveCustomer"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    ' Fast path — form name and address in DB cache
    Dim CustomerID As Long
    CustomerID = m_CustomerSvc.GetIDByFormCustomer(Name, Address, Address2)
    If CustomerID > 0 Then
        Dim C As clsCustomer
        Set C = m_CustomerSvc.LoadCustomer(CustomerID)
        Set ResolveCustomer = C
        Exit Function
    End If

    ' Slow path — launch UI
    Dim Frm As frmTRCustomerSelector
    Set Frm = New frmTRCustomerSelector
    Frm.Initialize m_CustomerSvc.AllCustomers, Name, Address, m_Logger
    Frm.Show

    Do
        If Frm.Cancelled Then
            Set ResolveCustomer = Nothing
            Exit Function
        End If

        If Frm.IsNewCustomer Then
            If m_CustomerSvc.IsUniqueName(Frm.NewCustomerDict("customer_name")) Then
                Set C = m_CustomerSvc.CreateCustomer(Frm.NewCustomerDict)
                m_CustomerSvc.AddFormKey Name, Address, Address2, C.ID
                Set ResolveCustomer = C
                Exit Function
            Else
                Frm.ShowValidationError "Customer name already exists"
                Frm.Show
            End If
        Else
            Set C = m_CustomerSvc.LoadCustomer(Frm.SelectedCustomerID)
            m_CustomerSvc.AddFormKey Name, Address, Address2, C.ID
            Set ResolveCustomer = C
            Exit Function
        End If
    Loop
    Exit Function

EH:                             HandleError "ResolveCustomer", "Name: " & Name & " | Address: " & Address
End Function
Public Function ResolveChemical(Name As String) As clsChemical
    AssertInit "ResolveChemical"
    If Not m_Logger.DebugMode Then On Error GoTo EH

    ' Fast path — form name in DB cache
    Dim ChemicalID As Long
    ChemicalID = m_ChemicalSvc.GetIDByFormName(Name)
    If ChemicalID > 0 Then
        Dim C As clsChemical
        Set C = m_ChemicalSvc.LoadChemical(ChemicalID)
        Set ResolveChemical = C
        Exit Function
    End If

    ' Slow path — launch UI
    Dim Frm As frmTRChemicalSelector
    Set Frm = New frmTRChemicalSelector
    Frm.Initialize m_ChemicalSvc.AllChemicals, _
                   m_ChemicalSvc.MetalsPrepList, _
                   m_ChemicalSvc.SiliconPrepList, _
                   m_ChemicalSvc.IonsPrepList, _
                   m_ElementSvc.AllElements, _
                   Name, _
                   m_Logger
    Frm.Show

    Do
        If Frm.Cancelled Then
            Set ResolveChemical = Nothing
            Exit Function
        End If

        If Frm.IsNewChemical Then
            If m_ChemicalSvc.IsUniqueName(Frm.NewChemicalDict("chemical_name")) Then
                Set C = m_ChemicalSvc.CreateChemical(Frm.NewChemicalDict)
                ' Persist form name mapping and update session cache
                m_ChemicalSvc.AddFormName Name, C.ID
                Set ResolveChemical = C
                Exit Function
            Else
                Frm.ShowValidationError "Chemical name already exists"
                Frm.Show
            End If
        Else
            Set C = m_ChemicalSvc.LoadChemical(Frm.SelectedChemicalID)
            ' Persist form name mapping and update session cache
            m_ChemicalSvc.AddFormName Name, C.ID
            Set ResolveChemical = C
            Exit Function
        End If
    Loop
    Exit Function

EH:                             HandleError "ResolveChemical", "Name: " & Name
End Function

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------
Private Sub AssertInit(Src As String)
    If m_Logger Is Nothing Then
        Err.Raise vbObjectError + 1, "clsTRFormInputResolver." & Src, "Not initialized"
    End If
End Sub

Private Sub HandleError(Src As String, Optional Context As String = "")
    m_Logger.LogError "clsTRFormInputResolver." & Src, Err.Number, Err.Description & " | " & Context
End Sub
