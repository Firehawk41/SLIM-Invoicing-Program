VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSLIMDictionaryManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsSLIMDictionaryManager
' Version: 1.1.1
' Author: JT
' Created: 2025-??-??
' LastModified: 2025-06-04
' Description: Class used to access an Access Database by dependency injection
' DependsOn: clsLoggingSystem, clsAccessDatabase
' ChangeLog:
'   - 1.1.1 - 2025-05-28    - added metadata
Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_WasAccessDBObjectPassed As Boolean
Private m_ChemicalDict As Scripting.Dictionary
Private m_KEDDict As Scripting.Dictionary
Private m_ElementDict As Scripting.Dictionary
Private m_CustomerDict As Scripting.Dictionary
Private m_AnalysisDict As Scripting.Dictionary
Private m_AnalysisDescriptionDict As Scripting.Dictionary
Private Enum ObjectTypeEnum
    ObjectTypeElement = 1
    ObjectTypeChemical
    ObjectTypeCustomer
    ObjectTypeAnalysis
    ObjectTypeKEDElement
End Enum
' Public property to get the Chemical Dictionary
Public Property Get ChemicalDict() As Scripting.Dictionary
    If m_ChemicalDict Is Nothing Then
        Set m_ChemicalDict = CreateChemicalDictionary()
    End If
    Set ChemicalDict = m_ChemicalDict
End Property

' Public property to get the KED Dictionary
Public Property Get KEDDict() As Scripting.Dictionary
    If m_KEDDict Is Nothing Then
        Set m_KEDDict = CreateKEDDictionary()
    End If
    Set KEDDict = m_KEDDict
End Property

Public Property Get ElementDict() As Scripting.Dictionary
    If m_ElementDict Is Nothing Then
        Set m_ElementDict = CreateElementDictionary()
    End If
    Set ElementDict = m_ElementDict
End Property

Public Property Get CustomerDict() As Scripting.Dictionary
    If m_CustomerDict Is Nothing Then
        Set m_CustomerDict = CreateCustomerDictionary()
    End If
    Set CustomerDict = m_CustomerDict
End Property

Public Property Get AnalysisDict() As Scripting.Dictionary
    If m_AnalysisDict Is Nothing Then
        Set m_AnalysisDict = CreateAnalysisDictionary()
    End If
    Set AnalysisDict = m_AnalysisDict
End Property

Public Property Get AnalysisDescriptionDict() As Scripting.Dictionary
    If m_AnalysisDescriptionDict Is Nothing Then
        Set m_AnalysisDescriptionDict = CreateAnalysisDescriptionDictionary()
    End If
    Set AnalysisDescriptionDict = m_AnalysisDescriptionDict
End Property

Public Sub Initialize(Logger As clsLoggingSystem, Optional AccessDB As clsAccessDatabase)
    Const DATABASE_FILE_PATH As String = "\\PRECILAB-SERVER\LabPlusServer\Documents_In_Works\Thomson\Sample Login.accdb"
    
    Set m_Logger = Logger
    
    If AccessDB Is Nothing Then
        Set m_AccessDB = New clsAccessDatabase
        Call m_AccessDB.Initialize(DATABASE_FILE_PATH, Logger)
        m_WasAccessDBObjectPassed = False
    Else
        Set m_AccessDB = AccessDB
        m_WasAccessDBObjectPassed = True
    End If

End Sub
' Generic Error Handling Subroutine
Private Sub HandleError(FunctionName As String, ErrMsg As String)
    Call m_Logger.LogMessage(FunctionName, LogDebug, ErrMsg)
End Sub
' Helper Function to Initialize and Add Items to a Dictionary
Private Sub AddSingleItemToDictionary(Dict As Scripting.Dictionary, Key As String, Item As Object)
    If Not m_Logger.DebugMode Then
        On Error GoTo ErrorHandler
    End If
    
    ' Handle the case where items are single items
    If Not Dict.Exists(Key) Then
        Dict.Add Key, Item
    Else
        Call m_Logger.LogMessage("AddToDictionaryIfNotExists", LogLevelEnum.LogWarning, "Duplicate entry detected for key: " & Key)
    End If

    Exit Sub
ErrorHandler:
    Call HandleError("AddSingleItemToDictionary", "Error adding to dictionary: " & Err.Description)
End Sub
' Helper Function to Initialize and Add Collections to a Dictionary
Private Sub AddItemToCollectionInDictionary(Dict As Scripting.Dictionary, Key As String, Item As Object)
    Dim ItemCollection As Collection
    
    If Not m_Logger.DebugMode Then
        On Error GoTo ErrorHandler
    End If
    
    If Not Dict.Exists(Key) Then
        ' Create a new collection and add the item
        Set ItemCollection = New Collection
        ItemCollection.Add Item
        Dict.Add Key, ItemCollection
    Else
        ' Add the item to the existing collection
        Set ItemCollection = Dict(Key)
        ItemCollection.Add Item
    End If

    Exit Sub
ErrorHandler:
    Call HandleError("AddItemToCollectionInDictionary", "Error adding to dictionary: " & Err.Description)
End Sub
' Function to Create a Generic Dictionary from Query Results
Private Function CreateDictionaryFromQuery(SQLQuery As String, ObjectType As ObjectTypeEnum, Optional ExpectCollection As Boolean = False, Optional KeyFieldIndex As Long = 1, Optional UseLowerCaseKeys As Boolean = False) As Scripting.Dictionary
    Dim ResultArray As Variant
    Dim Dict As Scripting.Dictionary
    Dim ArrayIndex As Long
    Dim Item As Object
    Dim Key As String

    If Not m_Logger.DebugMode Then
        On Error GoTo ErrorHandler
    End If

    ' Initialize the dictionary
    Set Dict = New Scripting.Dictionary

    ' Execute SQL and get the array
    ResultArray = m_AccessDB.ExecuteQuery(SQLQuery)

    ' Loop through each record in the array and create objects
    For ArrayIndex = 0 To UBound(ResultArray, 2)
        ' Initialize the item based on the passed ObjectType
        Set Item = CreateObjectByType(ObjectType, ResultArray, ArrayIndex)
        
        ' Determine the key, using lowercase if specified
        Key = ResultArray(KeyFieldIndex, ArrayIndex)
        If UseLowerCaseKeys Then Key = LCase(Key)

        ' Add the item to the dictionary using a suitable key
        If ExpectCollection Then
            Call AddItemToCollectionInDictionary(Dict, Key, Item)
        Else
            Call AddSingleItemToDictionary(Dict, Key, Item)
        End If
        
    Next ArrayIndex

    Set CreateDictionaryFromQuery = Dict
    Exit Function
    
ErrorHandler:
    Call HandleError("CreateDictionaryFromQuery", "Error creating dictionary from query: " & Err.Description)
    Set CreateDictionaryFromQuery = Nothing
End Function
Private Function CreateObjectByType(ObjectType As ObjectTypeEnum, ResultArray As Variant, Index As Long) As Object
    Select Case ObjectType
        Case ObjectTypeElement
            Set CreateObjectByType = InitializeElement(ResultArray, Index)
        Case ObjectTypeChemical
            Set CreateObjectByType = InitializeChemical(ResultArray, Index)
        Case ObjectTypeCustomer
            Set CreateObjectByType = InitializeCustomer(ResultArray, Index)
        Case ObjectTypeAnalysis
            Set CreateObjectByType = InitializeAnalysis(ResultArray, Index)
        Case ObjectTypeKEDElement
            Set CreateObjectByType = InitializeKEDElement(ResultArray, Index)
        Case Else
            Err.Raise 1999, "CreateObjectByType", "Unrecognized object type."
    End Select
End Function

' Callbacks for Creating Specific Objects
Private Function InitializeElement(ResultArray As Variant, Index As Long) As clsTestingRequestElement
    Dim Element As clsTestingRequestElement
    Dim ElementNumber As Long
    Dim ElementSymbol As String
    Dim ElementName As String
    
    Set Element = New clsTestingRequestElement
    ElementNumber = ResultArray(0, Index)
    ElementSymbol = ResultArray(1, Index)
    ElementName = ResultArray(2, Index)
    
    Call Element.Initialize(ElementNumber, ElementSymbol, ElementName, m_Logger)
    Set InitializeElement = Element
End Function

Private Function InitializeChemical(ResultArray As Variant, Index As Long) As clsTestingRequestChemical
    Dim Chemical As clsTestingRequestChemical
    Dim ChemicalID As Long
    Dim ChemicalName As String
    Dim MetalsPrep As String
    Dim SiliconPrep As String
    Dim IonsPrep As String
    Dim NameOnForm As String
    
    Set Chemical = New clsTestingRequestChemical
    ChemicalID = ResultArray(0, Index)
    ChemicalName = ResultArray(1, Index)
    MetalsPrep = ResultArray(2, Index)
    SiliconPrep = ResultArray(3, Index)
    IonsPrep = ResultArray(4, Index)
    NameOnForm = ResultArray(5, Index)
    
    Call Chemical.Initialize(ChemicalName, MetalsPrep, SiliconPrep, IonsPrep, NameOnForm, m_Logger, ChemicalID)
    ' If the KED Dictionary is nothing, set it from the database
    If m_KEDDict Is Nothing Then
        Set m_KEDDict = CreateKEDDictionary()
    End If
    ' If the chemical has KED elements assigned to it, return them
    If m_KEDDict.Exists(ChemicalName) Then
        Set Chemical.KEDElements = m_KEDDict(ChemicalName)
    End If
    Set InitializeChemical = Chemical
End Function

Private Function InitializeCustomer(ResultArray As Variant, Index As Long) As clsTestingRequestCustomer
    Dim Customer As clsTestingRequestCustomer
    Dim CustomerID As Long
    Dim CustomerName As String
    Dim StreetAddress As String
    Dim City As String
    Dim State As String
    Dim PostalCode As String
    Dim Country As String
    Dim NameOnForm As String
    Dim AddressOnForm As String
    
    Set Customer = New clsTestingRequestCustomer
    CustomerID = NullOrValue(ResultArray(0, Index))
    CustomerName = NullOrValue(ResultArray(1, Index))
'    On Error Resume Next ' leaves the following properties blank if they are blank in the database (technically they are Null, which causes and error because Null is not a string)
    StreetAddress = NullOrValue(ResultArray(2, Index))
    City = NullOrValue(ResultArray(3, Index))
    State = NullOrValue(ResultArray(4, Index))
    PostalCode = NullOrValue(ResultArray(5, Index))
    Country = NullOrValue(ResultArray(6, Index))
'    On Error GoTo 0 ' reset error handling
    NameOnForm = NullOrValue(ResultArray(7, Index))
    AddressOnForm = NullOrValue(ResultArray(8, Index))
    
    Call Customer.Initialize(CustomerName, StreetAddress, City, State, PostalCode, Country, NameOnForm, AddressOnForm, m_Logger, CustomerID)
    Set InitializeCustomer = Customer
End Function
Private Function NullOrValue(Value As Variant) As String
    If IsNull(Value) Then
        NullOrValue = ""
    Else
        NullOrValue = Value
    End If
    
End Function
Private Function InitializeAnalysis(ResultArray As Variant, Index As Long) As clsTestingRequestAnalysis
    Dim Analysis As clsTestingRequestAnalysis
    Dim AnalysisID As Long
    Dim AnalysisName As String
    Dim AnalysisDescription As String
    
    Set Analysis = New clsTestingRequestAnalysis
    AnalysisID = ResultArray(0, Index)
    AnalysisName = ResultArray(1, Index)
    AnalysisDescription = ResultArray(2, Index)
    
    Call Analysis.Initialize(AnalysisID, AnalysisName, AnalysisDescription, m_Logger)
    Set InitializeAnalysis = Analysis
End Function
Private Function InitializeKEDElement(ResultArray As Variant, Index As Long) As clsTestingRequestElement
    Dim Element As clsTestingRequestElement
    Dim ElementNumber As Long
    Dim ElementSymbol As String
    Dim ElementName As String
    
    Set Element = New clsTestingRequestElement
    ElementNumber = ResultArray(0, Index)
    ElementSymbol = ResultArray(2, Index)
    ElementName = ResultArray(3, Index)
    
    Call Element.Initialize(ElementNumber, ElementSymbol, ElementName, m_Logger)
    Set InitializeKEDElement = Element
End Function

' Main Functions to Create Specific Dictionaries
Private Function CreateElementDictionary() As Scripting.Dictionary
    Dim SQLQuery As String
    SQLQuery = "SELECT ID, element_symbol, element_name FROM elements"
    Set CreateElementDictionary = CreateDictionaryFromQuery(SQLQuery, ObjectTypeElement)
    Call m_Logger.LogMessage("CreateElementDictionary", LogLevelEnum.LogDebug, "Successfully compiled the dictionary of elements.")
End Function

Private Function CreateChemicalDictionary() As Scripting.Dictionary
    Dim SQLQuery As String
    SQLQuery = "SELECT chemicals.ID, chemicals.chemical_name, chemicals.metals_prep, chemicals.silicon_prep, chemicals.ions_prep, form_chemicals.form_chemical_name " & _
               "FROM form_chemicals INNER JOIN chemicals ON form_chemicals.chemical_id = chemicals.ID " & _
               "ORDER BY chemicals.chemical_name ASC"
    Set CreateChemicalDictionary = CreateDictionaryFromQuery(SQLQuery, ObjectTypeChemical, True, , True)
    Call m_Logger.LogMessage("CreateChemicalDictionary", LogLevelEnum.LogDebug, "Successfully compiled the dictionary of chemicals.")
End Function

Private Function CreateCustomerDictionary() As Scripting.Dictionary
    Dim SQLQuery As String
    SQLQuery = "SELECT customers.ID, customers.customer_name, customers.street_address, customers.city, customers.state, customers.postal_code, customers.country, form_customers.form_customer, form_customers.form_address " & _
               "FROM form_customers INNER JOIN customers ON form_customers.customer_id = customers.ID " & _
               "ORDER BY customers.customer_name ASC"
    Set CreateCustomerDictionary = CreateDictionaryFromQuery(SQLQuery, ObjectTypeCustomer, True)
    Call m_Logger.LogMessage("CreateCustomerDictionary", LogLevelEnum.LogDebug, "Successfully compiled the dictionary of customers.")
End Function

Private Function CreateAnalysisDictionary() As Scripting.Dictionary
    Dim SQLQuery As String
    SQLQuery = "SELECT analyses.id, analyses.analysis_name, analyses.analysis_description, form_analyses.form_analysis_name " & _
               "FROM analyses INNER JOIN (form_analyses INNER JOIN form_analysis_mapping ON form_analyses.ID = form_analysis_mapping.form_analysis_id) " & _
               "ON analyses.id = form_analysis_mapping.database_analysis_id"
    Set CreateAnalysisDictionary = CreateDictionaryFromQuery(SQLQuery, ObjectTypeAnalysis, True, 3)
    Call m_Logger.LogMessage("CreateAnalysisDictionary", LogLevelEnum.LogDebug, "Successfully compiled the dictionary of analyses.")
End Function
Private Function CreateAnalysisDescriptionDictionary() As Scripting.Dictionary
    Dim SQLQuery As String
    SQLQuery = "SELECT analyses.id, analyses.analysis_name, analyses.analysis_description FROM analyses"
    Set CreateAnalysisDescriptionDictionary = CreateDictionaryFromQuery(SQLQuery, ObjectTypeAnalysis)
    Call m_Logger.LogMessage("CreateAnalysisDescriptionDictionary", LogLevelEnum.LogDebug, "Successfully compiled the dictionary of analyses descriptions.")
End Function
Private Function CreateKEDDictionary() As Scripting.Dictionary
    Dim SQLQuery As String
    SQLQuery = "SELECT elements.ID, elements.element_symbol, elements.element_name, chemicals.chemical_name " & _
                "FROM elements " & _
                "INNER JOIN (chemicals " & _
                "INNER JOIN ked_elements ON chemicals.ID = ked_elements.chemical_id) " & _
                "ON elements.ID = ked_elements.element_id"
    Set CreateKEDDictionary = CreateDictionaryFromQuery(SQLQuery, ObjectTypeElement, True, 3)
    Call m_Logger.LogMessage("CreateKEDDictionary", LogLevelEnum.LogDebug, "Successfully compiled the dictionary of KED Elements.")
End Function
Private Sub Class_Terminate()
    Set m_ChemicalDict = Nothing
    Set m_CustomerDict = Nothing
    Set m_AnalysisDict = Nothing
    Set m_KEDDict = Nothing
    Set m_ElementDict = Nothing
    If Not m_WasAccessDBObjectPassed Then
        Set m_AccessDB = Nothing
    End If
End Sub
