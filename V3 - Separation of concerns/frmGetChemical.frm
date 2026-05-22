VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGetChemical 
   Caption         =   "Get Chemical"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4485
   OleObjectBlob   =   "frmGetChemical.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGetChemical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' MacroName: frmGetChemical
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
Private m_ElementDict As Scripting.Dictionary
Private m_ChemicalDict As Scripting.Dictionary
Private m_ChemicalList As Variant
Private m_MetalsPrepList As Variant
Private m_SiliconPrepList As Variant
Private m_IonsPrepList As Variant
Private m_ChemicalDefault As String
Private m_MetalsDefault As String
Private m_SiliconDefault As String
Private m_IonsDefault As String
Private m_KEDDefault As String
Const elementList As String = "Li,Be,B,Na,Mg,Al,K,Ca,Ti,V,Cr,Mn,Fe,Co,Ni,Cu,Zn,Ga,Ge,As,Sr,Zr,Nb,Mo,Ag,Cd,Sn,Sb,Ba,Ta,W,Pt,Au,Tl,Pb,Bi"
Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property
Public Property Let KEDDefault(Value As String)
    Dim ElementArr() As String

    m_KEDDefault = Value
    ElementArr = Split(m_KEDDefault, ",")
    Me.ListBox2.List = ElementArr
End Property

Public Property Get ChemName() As String
    If ExistingChemicalButton.Value = True Then
        ChemName = ExistingChemicalBox.Value
    Else
        ChemName = PRECILABChemicalBox.Value
    End If
End Property
Public Property Get ICPMSPrep() As String
    If NewChemicalButton.Enabled = True Then ICPMSPrep = ICPMSPrepBox.Text
End Property
Public Property Get ICPrep() As String
    If NewChemicalButton.Enabled = True Then ICPrep = ICPrepBox.Text
End Property
Public Property Get SiliconPrep() As String
    If NewChemicalButton.Enabled = True Then SiliconPrep = SiPrepBox.Text
End Property
Public Property Get KEDSMode() As Collection
    If NewChemicalButton.Enabled = True Then Set KEDSMode = CreateKEDElementCollection
End Property
Public Sub Initialize(AllCustomers As Collection, UnmatchedName As String, Logger As clsLoggingSystem)

    Debug.Assert Not Logger Is Nothing
    Debug.Assert m_Logger Is Nothing ' assert only 1 init

    Set m_Logger = Logger


    ' Sets up variables for userform
    Set m_Logger = Logger
    Set m_ElementDict = SLIMDictManager.ElementDict
    Set m_ChemicalDict = SLIMDictManager.ChemicalDict

    m_ChemicalList = SetChemicalList(SLIMDictManager.ChemicalDict)
    Me.ExistingChemicalBox.List = m_ChemicalList
    Me.GetChemMessageBox.Caption = FormChemicalName & " not found in the database. If it's in the database as another name, please select its match. If it's a new chemical, please enter in its information."
    Me.PRECILABChemicalBox.Value = FormChemicalName

    ' Sets the prep lists and defaults
    Select Case RequestType
        Case TestingRequestTypeEnum.Wafer
            Me.ICPMSPrepBox.List = Array("LP-ICPMS")
            Me.SiPrepBox.List = Array("N/A")
            Me.ICPrepBox.List = Array("N/A")
            Me.ICPMSPrepBox.Value = "LP-ICPMS"
            Me.SiPrepBox.Value = "N/A"
            Me.ICPrepBox.Value = "N/A"

        Case TestingRequestTypeEnum.Water, TestingRequestTypeEnum.Chemical

            Me.ICPMSPrepBox.List = Array("N/A", "Evaporation", "Dilute and Shoot", "Organic Dilute & Shoot", "Standard Addition")
            Me.SiPrepBox.List = Array("N/A", "Evaporation", "Dilute and Shoot")
            Me.ICPrepBox.List = Array("N/A", "Dilute and Shoot", "Evaporation", "Neutralizing", "Extraction", "Standard Addition")
            Me.ICPMSPrepBox.Value = "Evaporation"
            Me.SiPrepBox.Value = "Evaporation"
            Me.ICPrepBox.Value = "Dilute and Shoot"

    End Select

End Sub
Private Function SetChemicalList(ChemicalDict As Scripting.Dictionary) As Variant
    Dim ChemicalList() As Variant
    Dim Chemical As clsTestingRequestChemical
    Dim DictIndex As Long
    Dim Key As Variant

    ReDim ChemicalList(ChemicalDict.Count - 1)


    For Each Key In ChemicalDict.Keys
        ChemicalList(DictIndex) = Key
        DictIndex = DictIndex + 1
    Next Key

    SetChemicalList = ChemicalList
    Call m_Logger.LogMessage("SetChemicalList", LogLevelEnum.LogDebug, UBound(ChemicalList) + 1 & " chemicals found in the database.")

End Function

Private Function CreateKEDElementCollection() As Collection
    Dim Element As Variant
    Dim KEDElementCollection As New Collection
    Dim ElementObject As clsTestingRequestElement
    Dim ListIndex As Long

    For ListIndex = 0 To ListBox2.ListCount - 1
        Element = ListBox2.List(ListIndex)

        ' Retrieve the corresponding object from the dictionary
        If m_ElementDict.Exists(Element) Then
            Set ElementObject = m_ElementDict(Element)
            ' Add the element object to the collection
            KEDElementCollection.Add ElementObject
        Else
            ' Handle the case where the element is not found in the dictionary
            Call m_Logger.LogMessage("CreateKEDElementCollection", LogLevelEnum.LogWarning, "Element not found in dictionary: " & Element)
        End If

    Next ListIndex

    ' Return to the calling sub
    Set CreateKEDElementCollection = KEDElementCollection

End Function

Private Sub AddElementButton_Click()

    Dim ItemExists As Boolean
    Dim iIndex As Long
    Dim ItemValue As String

    If ListBox1.ListIndex <> -1 Then

        ItemValue = ListBox1.List(ListBox1.ListIndex)
        ItemExists = False
        For iIndex = 0 To ListBox2.ListCount - 1
            If ListBox2.List(iIndex) = ItemValue Then
                ItemExists = True
                Exit For
            End If
        Next iIndex

        If Not ItemExists Then
            ListBox2.AddItem ItemValue
        End If
    End If
End Sub

Private Sub CancelButton_Click()
    m_Cancelled = True
    Me.Hide

    Call m_Logger.LogMessage("GetChemicalForm", LogLevelEnum.LogInfo, "User cancelled userform.")
End Sub
Private Sub ExistingChemicalButton_Click()
    Call SetChemicalControlsState(True)
End Sub


Private Sub NewChemicalButton_Click()
    Call SetChemicalControlsState(False)
End Sub
Private Sub SetChemicalControlsState(IsExisting As Boolean)
    ' Sets the format

    ExistingChemicalBox.Enabled = IsExisting
    ExistingChemicalBox.BackColor = IIf(IsExisting, vbWhite, &H8000000F)

    PRECILABChemicalBox.Enabled = Not IsExisting
    PRECILABChemicalBox.BackColor = IIf(Not IsExisting, vbWhite, &H8000000F)

    ICPMSPrepBox.Enabled = Not IsExisting
    ICPMSPrepBox.BackColor = IIf(Not IsExisting, vbWhite, &H8000000F)

    SiPrepBox.Enabled = Not IsExisting
    SiPrepBox.BackColor = IIf(Not IsExisting, vbWhite, &H8000000F)

    ICPrepBox.Enabled = Not IsExisting
    ICPrepBox.BackColor = IIf(Not IsExisting, vbWhite, &H8000000F)

    ListBox1.Enabled = Not IsExisting
    ListBox1.BackColor = IIf(Not IsExisting, vbWhite, &H8000000F)

    ListBox2.Enabled = Not IsExisting
    ListBox2.BackColor = IIf(Not IsExisting, vbWhite, &H8000000F)

End Sub
Private Sub OKButton_Click()

    If ValidateInputs() Then
        Me.Hide
    Else
        Call m_Logger.LogMessage("GetChemicalForm", LogLevelEnum.LogWarning, "Validation failed. Please correct the highlighted errors.", , True)
    End If
End Sub

Private Function ValidateInputs() As Boolean

    Dim iIndex As Long

    ' Check if the existing chemical is recognized
    If ExistingChemicalButton.Value Then

        ' Initialize validation flag
        ValidateInputs = False

        ' Loop through the chemical list
        For iIndex = LBound(m_ChemicalList) To UBound(m_ChemicalList)
            If m_ChemicalList(iIndex) = ExistingChemicalBox.Text Then
                ' if a match is found, set the input to valid
                ValidateInputs = True
                Exit For
            End If
        Next iIndex

        ' Colour the combobox based on whether the input was valid
        If ValidateInputs Then
            ExistingChemicalBox.BackColor = vbWhite
        Else
            ExistingChemicalBox.BackColor = vbRed
        End If

    Else

        ' Initialize validation flag
        ValidateInputs = True

        ' Checks whether the PRECILAB Chemical Box has a valid input
        If PRECILABChemicalBox.Text = "" Then
            PRECILABChemicalBox.BackColor = vbRed
            ValidateInputs = False
        Else
            PRECILABChemicalBox.BackColor = vbWhite
        End If

        ' Checks whether the ICPMS Prep Box has a valid input
        If ICPMSPrepBox.Text = "" Then
            ICPMSPrepBox.BackColor = vbRed
            ValidateInputs = False
        Else
            ICPMSPrepBox.BackColor = vbWhite
        End If

        ' Checks whether the Si Prep Box has a valid input
        If SiPrepBox.Text = "" Then
            SiPrepBox.BackColor = vbRed
            ValidateInputs = False
        Else
            SiPrepBox.BackColor = vbWhite
        End If

        ' Checks whether the IC Prep Box has a valid input
        If ICPrepBox.Text = "" Then
            ICPrepBox.BackColor = vbRed
            ValidateInputs = False
        Else
            ICPrepBox.BackColor = vbWhite
        End If

    End If

End Function


Private Sub RemoveElementButton_Click()

    If Not ListBox2.ListIndex = -1 Then
        ListBox2.RemoveItem ListBox2.ListIndex
    End If
End Sub

Private Sub UserForm_Initialize()
    ExistingChemicalButton = True
    Call PopulateListBox

End Sub
Private Sub PopulateListBox()
    Dim ElementArr() As String
    Dim Element As Variant
    Dim iIndex As Integer
    Dim InListBox2 As Boolean

    ' Split the ElementList string into an array
    ElementArr = Split(elementList, ",")

    ' Clear ListBox1 before populating
    ListBox1.Clear
    ListBox2.Clear

    ' Loop through each element in the array
    For Each Element In ElementArr
        InListBox2 = False

        ' Check if the element is in ListBox2
        For iIndex = 0 To ListBox2.ListCount - 1
            If ListBox2.List(iIndex) = Element Then
                InListBox2 = True
                Exit For
            End If
        Next iIndex

        ' If the element is not in ListBox2, add it to ListBox1
        If Not InListBox2 Then
            ListBox1.AddItem Element
        End If
    Next Element

End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        m_Cancelled = True
        Cancel = True
        Me.Hide
    End If
End Sub


