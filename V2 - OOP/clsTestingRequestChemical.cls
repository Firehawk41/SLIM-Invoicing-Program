VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestingRequestChemical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsTestingRequestChemical
' Version: 1.1.1
' Author: JT
' Created: ?
' LastModified: 2025-08-19
' Description: Stores properties relating to an individual chemical
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsTestingRequestElement, modEnums, clsSLIMDictionaryManager, frmGetChemical
' ChangeLog:
'   - 1.0.1 - 2025-06-05 - Added metadata
'   - 1.1.0 - 2025-08-12 - Added Sub to import data from JSON file
'   - 1.1.1 - 2025-08-19 - Fixed bugs for LoadFromJson
Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_ChemicalID As Long
Private m_ChemicalName As String
Private m_NameOnForm As String
Private m_MetalsPrep As String
Private m_SiliconPrep As String
Private m_IonsPrep As String
Private m_KEDElements As Collection
Private m_DatabaseEntryDate As Date
Public Property Get ChemicalID() As Long
    ChemicalID = m_ChemicalID
End Property
Public Property Get ChemicalName() As String
    ChemicalName = m_ChemicalName
End Property
Public Property Get MetalsPrep() As String
    MetalsPrep = m_MetalsPrep
End Property
Public Property Get SiliconPrep() As String
    SiliconPrep = m_SiliconPrep
End Property
Public Property Get IonsPrep() As String
    IonsPrep = m_IonsPrep
End Property
Public Property Get KEDElements() As Collection
    Set KEDElements = m_KEDElements
End Property
Public Property Set KEDElements(Value As Collection)
    Set m_KEDElements = KEDElements
End Property
Public Property Get NameOnForm() As String
    NameOnForm = m_NameOnForm
End Property
Public Sub Initialize(ChemicalName As String, MetalsPrep As String, SiliconPrep As String, IonsPrep As String, NameOnForm As String, Logger As clsLoggingSystem, Optional ChemicalID As Long = 0)
    Set m_Logger = Logger
    
    m_ChemicalID = ChemicalID
    m_ChemicalName = ChemicalName
    m_MetalsPrep = MetalsPrep
    m_SiliconPrep = SiliconPrep
    m_IonsPrep = IonsPrep
    m_NameOnForm = NameOnForm
    
End Sub
Public Sub SetFromDictionary(NameOnForm As String, RequestType As TestingRequestTypeEnum, Logger As clsLoggingSystem, AccessDB As clsAccessDatabase, SLIMDictionaryManager As clsSLIMDictionaryManager, Optional UsePartialMatching As Boolean = False)
    Dim Key As Variant
    Dim ChemicalCollection As Collection
    Dim Chemical As clsTestingRequestChemical
    Dim IsChemicalSet As Boolean

    Call Logger.LogMessage("SetFromDictionary", LogLevelEnum.LogDebug, "Attempting to set chemical object from the chemical dictionary.")
    
    ' Initialize properties
    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    m_NameOnForm = NameOnForm
    IsChemicalSet = False
    
    ' Search for a matching customer in the dictionary
    If Not SetChemical(SLIMDictionaryManager.ChemicalDict) Then
        ' Fallback if no matching customer was found
        Call RetrieveChemicalFromUserform(RequestType, SLIMDictionaryManager)
    End If
    
    Call m_Logger.LogMessage("SetFromDictionary", LogLevelEnum.LogDebug, "Finished initializing chemical object.")
    
End Sub
Private Function SetChemical(ChemicalDict As Scripting.Dictionary) As Boolean
    Dim Key As Variant
    Dim ChemicalCollection As Collection
    Dim Chemical As clsTestingRequestChemical
    
    ' Search for a matching customer in the dictionary
    For Each Key In ChemicalDict

        ' Validate that the value is a collection of customers
        If TypeOf ChemicalDict(Key) Is Collection Then
            Set ChemicalCollection = ChemicalDict(Key)
            
            ' Search within the chemical collection
            Set Chemical = FindChemicalInCollection(ChemicalCollection)
            If Not Chemical Is Nothing Then
                m_ChemicalName = Chemical.ChemicalName
                m_MetalsPrep = Chemical.MetalsPrep
                m_SiliconPrep = Chemical.SiliconPrep
                m_IonsPrep = Chemical.IonsPrep
                SetChemical = True
                Call m_Logger.LogMessage("SetChemical", LogLevelEnum.LogInfo, "Found the matching chemical object in the dictionary and successfully set chemical properties.")
                Exit Function

            End If

        Else
            ' Handle unexpected data types in the dictionary
            Call m_Logger.LogMessage("SetChemical", LogLevelEnum.LogWarning, "Unexpected value in ChemicalDict for Key: " & Key & ", Type: " & TypeName(ChemicalDict(Key)))
        End If
    Next Key
    
    ' Return false if no match
    SetChemical = False
    
End Function
Private Function FindChemicalInCollection(ChemicalCollection As Collection) As clsTestingRequestChemical
    Dim Chemical As clsTestingRequestChemical
    
    ' Inner loop through the collection
    For Each Chemical In ChemicalCollection
        If Chemical.NameOnForm = NameOnForm Then
            Set FindChemicalInCollection = Chemical
            Exit Function
        End If
    Next Chemical
    
    ' Return nothing if no match is found
    Set FindChemicalInCollection = Nothing
End Function
Private Sub RetrieveChemicalFromUserform(RequestType As TestingRequestTypeEnum, SLIMDictManager As clsSLIMDictionaryManager, Optional IsNewChemical As Boolean = False)

    Dim UF As frmGetChemical
    Dim UseExistingChemical As Boolean
    Dim NewChemical As clsTestingRequestChemical
    Dim ChemicalDict As Scripting.Dictionary
    Dim UFCancelled As Boolean

    
    ' Initialize and show the userform
    Set UF = New frmGetChemical
    Call UF.Initialize(m_NameOnForm, RequestType, SLIMDictManager, m_Logger)
    If IsNewChemical Then
        UF.GetChemMessageBox.Caption = "Please enter the details for the new chemical"
        UF.ExistingChemicalButton.Enabled = False
        UF.NewChemicalButton.Value = True
    End If
    UF.Show
    
    ' Retrieve variables from the userform
    m_ChemicalName = UF.ChemName
    m_MetalsPrep = UF.ICPMSPrep
    m_SiliconPrep = UF.SiliconPrep
    m_IonsPrep = UF.ICPrep
    Set m_KEDElements = UF.KEDSMode
    UseExistingChemical = UF.ExistingChemicalButton.Value
    UFCancelled = UF.Cancelled
    
    ' Unload the userform
    Unload UF
    
    ' Raise an error if the user cancelled the form
    If UFCancelled Then
        Err.Raise 900, "GetChemical", "User cancelled the Get Chemical userform"
    End If
    
    ' Retrieve or insert chemical in the database
    If UseExistingChemical Or SLIMDictManager.ChemicalDict.Exists(LCase(m_ChemicalName)) Then
        m_ChemicalID = GetExistingChemicalID(SLIMDictManager.ChemicalDict)
    Else
        m_ChemicalID = InsertNewChemical(SLIMDictManager.ChemicalDict)
        Call AddElementToDatabase(KEDElements, "ked_elements", "chemical_id", m_ChemicalID)
    End If
    
    ' Add the form chemical name to the database
    Call InsertFormChemical
    
End Sub
Private Function GetExistingChemicalID(ChemicalDict As Scripting.Dictionary) As Long
    Dim ChemicalCollection As Collection
    Dim Chemical As clsTestingRequestChemical
    
    If ChemicalDict.Exists(LCase(m_ChemicalName)) Then
        If TypeOf ChemicalDict(LCase(m_ChemicalName)) Is Collection Then
            Set ChemicalCollection = ChemicalDict(LCase(m_ChemicalName))
            Set Chemical = ChemicalCollection(1)
            GetExistingChemicalID = Chemical.ChemicalID
            
            ' Add the chemical to the current collection
            ChemicalCollection.Add Me
        
        End If
        
    Else
        ' Handle case where chemical is not in the dictionary
        Err.Raise 901, "GetExistingChemicalID", "Chemical not found in dictionary"
    End If
    
End Function
Private Function InsertNewChemical(ChemicalDict As Scripting.Dictionary) As Long
    Dim SQLQuery As String
    Dim DBChemicalID As Long
    Dim ChemicalCollection As New Collection
    
    SQLQuery = "INSERT INTO chemicals (database_entry_date, chemical_name, metals_prep, silicon_prep, ions_prep) " & _
               "VALUES ('" & format(Now, "yyyy-mm-dd hh:mm:ss") & "', '" & m_ChemicalName & "', '" & m_MetalsPrep & "', '" & m_SiliconPrep & "', '" & m_IonsPrep & "')"
    DBChemicalID = m_AccessDB.ExecuteNonQuery(SQLQuery)
    
    ' Add a new collection to the dictionary with the chemical object in it
    ChemicalCollection.Add Me
    ChemicalDict.Add LCase(m_ChemicalName), ChemicalCollection
    
    ' Return the ID to the calling sub
    InsertNewChemical = DBChemicalID
End Function

Private Sub InsertFormChemical()
    Dim SQLQuery As String
    
    SQLQuery = "INSERT INTO form_chemicals (chemical_id, form_chemical_name) VALUES ('" & m_ChemicalID & "', '" & m_NameOnForm & "')"
    Call m_AccessDB.ExecuteNonQuery(SQLQuery)
End Sub
Private Sub AddElementToDatabase(ElementCollection As Collection, DBTable As String, DBColumn As String, LinkingID As Long)
    ' Loops through the analyses and adds them to the database
    
    Dim Element As clsTestingRequestElement
    Dim ElementID As Long
    Dim SQLCommand As String
    Dim NewID As Variant
    
    For Each Element In ElementCollection
        ' Set variables
        ElementID = Element.ElementID
        
        ' Set the SQL command and log it
        SQLCommand = "INSERT INTO " & DBTable & " (" & DBColumn & ", element_id) VALUES ('" & LinkingID & "', '" & ElementID & "')"
        
        ' Execute the SQL command and the new ID
        NewID = m_AccessDB.ExecuteNonQuery(SQLCommand)
        
        If IsNull(NewID) Then
            ' Raises an error in the case of failed SQL execution
            Err.Raise 804, "AddElementToDatabase", "Failed to add element to the database."
        End If
        
    Next Element
    
End Sub
Public Sub InitializeNewChemical(RequestType As TestingRequestTypeEnum, Logger As clsLoggingSystem, AccessDB As clsAccessDatabase, SLIMDictionaryManager As clsSLIMDictionaryManager)
    Call Logger.LogMessage("InitializeNewChemical", LogLevelEnum.LogDebug, "Initializing new chemical.")
    
    ' Initialize properties
    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    
    Call RetrieveChemicalFromUserform(RequestType, SLIMDictionaryManager, True)
    
    Call Logger.LogMessage("InitializeNewChemical", LogLevelEnum.LogInfo, "Initializing new chemical.")
End Sub
Public Sub LoadFromJson(Json As Dictionary, Logger As clsLoggingSystem)
    Call Logger.LogMessage("clsTestingRequestChemical.LoadFromJson", LogLevelEnum.LogDebug, "Loading data from json to Chemical object.")
    
    If Not Logger.DebugMode Then On Error GoTo ErrorHandler
    
    Set m_Logger = Logger
    
    m_ChemicalID = Json("id")
    m_NameOnForm = Json("name")
    If Json.Exists("chemical") Then
        m_ChemicalName = Json("chemical")("name")
        m_MetalsPrep = Json("chemical")("metals_prep")
        m_SiliconPrep = Json("chemical")("silicon_prep")
        m_IonsPrep = Json("chemical")("ions_prep")
    End If
    
    Call Logger.LogMessage("clsTestingRequestChemical.LoadFromJson", LogLevelEnum.LogDebug, "Completed loading data from json to Chemical object: " & m_ChemicalName & ".")
    Exit Sub
ErrorHandler:
    Dim ErrorMessage As String
     ErrorMessage = "Error encountered in clsTestingRequestChemical.LoadFromJson where the chemical is: " & m_ChemicalName
    Call Logger.LogMessage("clsTestingRequestChemical.LoadFromJson", LogLevelEnum.LogError, ErrorMessage)
    Err.Raise vbObjectError + 1000, "clsTestingRequestChemical.LoadFromJson", ErrorMessage
End Sub
Public Sub Class_Initialize()
    ' Initialize the Analyses collection
    Set m_KEDElements = New Collection

End Sub

Private Sub Class_Terminate()
    Set m_KEDElements = Nothing
End Sub
