VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceCustomerBSiteInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceCustomerBSiteInfo
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-11
' Description: Object for separating Customer B site
' DependsOn:
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Private m_Site As String
Private m_SiteContact As String
Private m_CompanyCode As String
Private m_CostCenter As String
Private m_Plant As String
Public Property Get Site() As String
    Site = m_Site
End Property
Public Property Get SiteContact() As String
    SiteContact = m_SiteContact
End Property
Public Property Get CompanyCode() As String
    CompanyCode = m_CompanyCode
End Property
Public Property Get CostCenter() As String
    CostCenter = m_CostCenter
End Property
Public Property Get Plant() As String
    Plant = m_Plant
End Property
Public Sub Initialize(CustomerName As String, Chemical As String)
    Dim LookupKey As String
    
    ' Handle Hahnville special case
    If LCase(CustomerName) = "customer b hahnville" Then
        If LCase(Chemical) = "methyl carbitol" Then
            LookupKey = "Customer B Hahnville - EOEG"
        Else
            LookupKey = "Customer B Hahnville - Amines"
        End If
    Else
        LookupKey = CustomerName
    End If

    ' Populate based on key
    Call LoadInfoFromKey(LookupKey)
End Sub
Private Sub LoadInfoFromKey(Key As String)
    Select Case Key
        Case "Customer B Carteret"
            SetFields "Customer B Carteret, NJ", "fergussn@customer-b.com", "0001", "56539", "D254"
        Case "Customer B Monument"
            SetFields "Customer B Monument, TX", "fergussn@customer-b.com", "0001", "56539", "D229"
        Case "Customer B Hahnville - Amines"
            SetFields "Hahnville, LA (SCO)", "kqhandy@customer-b.com", "2000", "11978", "E105"
        Case "Customer B Hahnville - EOEG"
            SetFields "Hahnville, LA (SCO)", "tsylve@customer-b.com", "2000", "37095", "E094"
        Case "Customer B Seadrift"
            SetFields "Seadrift, TX (SDO)", "atbrennan@customer-b.com", "2000", "11778", "E270"
        Case "Customer B Texas City"
            SetFields "Texas City, TX (TCO)", "alissargarza@customer-b.com", "2000", "37032", "A052"
        Case "Customer B Freeport"
            SetFields "VERESENE, Freeport, TX", "masomme@customer-b.com", "0001", "68235", "D232"
        Case "Customer B Plaquemine"
            SetFields "Louisiana Operations, Plaquemine, LA", "dciccarelli1@customer-b.com", "0001", "60035", "E042"
        Case Else
            Err.Raise vbObjectError + 1001, "clsInvoiceCustomerBSiteInfo", "Unknown site: " & Key
    End Select
End Sub
Private Sub SetFields(Site As String, Contact As String, CompanyCode As String, CostCenter As String, Plant As String)
    m_Site = Site
    m_SiteContact = Contact
    m_CompanyCode = CompanyCode
    m_CostCenter = CostCenter
    m_Plant = Plant
End Sub
