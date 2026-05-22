VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestingRequestCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsTestingRequestCustomer
' Version: 1.1.1
' Author: JT
' Created: ?
' LastModified: 2025-08-19
' Description: Stores properties relating to an individual customer
' DependsOn: clsLoggingSystem, clsAccessDatabase, clsSLIMDictionaryManager
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
'   - 1.1.0 - 2025-08-12 - Added sub to import data from json
'   - 1.1.1 - 2025-08-19 - Fixed bugs for LoadFromJson
Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_CustomerID As Long
Private m_CustomerName As String
Private m_StreetAddress As String
Private m_City As String
Private m_State As String
Private m_PostalCode As String
Private m_Country As String
Private m_NameOnForm As String
Private m_AddressOnForm As String
Private m_DatabaseEntryDate As Date
Public Property Get CustomerID() As Long
    CustomerID = m_CustomerID
End Property
Public Property Get CustomerName() As String
    CustomerName = m_CustomerName
End Property
Public Property Get StreetAddress() As String
    StreetAddress = m_StreetAddress
End Property
Public Property Get City() As String
    City = m_City
End Property
Public Property Get State() As String
    State = m_State
End Property
Public Property Get PostalCode() As String
    PostalCode = m_PostalCode
End Property
Public Property Get Country() As String
    Country = m_Country
End Property
Public Property Get FormattedCustomerAddress() As String
    FormattedCustomerAddress = FormatAddress(vbCrLf, True)
End Property
Public Property Get NameOnForm() As String
    NameOnForm = m_NameOnForm
End Property
Public Property Get AddressOnForm() As String
    AddressOnForm = m_AddressOnForm
End Property
Public Sub Initialize(CustomerName As String, StreetAddress As String, City As String, State As String, PostalCode As String, Country As String, NameOnForm As String, AddressOnForm As String, Logger As clsLoggingSystem, Optional CustomerID As Long = 0)
    Set m_Logger = Logger
    
    m_CustomerID = CustomerID
    m_CustomerName = CustomerName
    m_StreetAddress = StreetAddress
    m_City = City
    m_State = State
    m_PostalCode = PostalCode
    m_Country = Country
    m_NameOnForm = NameOnForm
    m_AddressOnForm = AddressOnForm
    
End Sub
Private Function FormatAddress(FormatSpacer As String, Optional IncludeCustomerName As Boolean = False) As String
    Dim Address As String
    
    If IncludeCustomerName Then
        Address = m_CustomerName & FormatSpacer
    End If
        
    Address = Address & m_StreetAddress & FormatSpacer & m_City
    
    If m_State <> "" Then
        Address = Address & ", " & m_State
    End If
    If m_PostalCode <> "" Then
        Address = Address & " " & m_PostalCode
    End If
    If m_Country <> "USA" And m_Country <> "" Then
        Address = Address & FormatSpacer & m_Country
    End If
    FormatAddress = Address
End Function

Public Sub SetFromDictionary(NameOnForm As String, AddressOnForm As String, Logger As clsLoggingSystem, AccessDB As clsAccessDatabase, CustomerDict As Scripting.Dictionary, Optional CustomerID As Long = 0)

    Call Logger.LogMessage("SetFromDictionary", LogLevelEnum.LogDebug, "Attempting to set customer object from the customer dictionary.")
    
    ' Initialize properties
    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    m_NameOnForm = NameOnForm
    m_AddressOnForm = AddressOnForm
    m_CustomerID = CustomerID
    
    ' Search for a matching customer in the dictionary
    If Not SetCustomer(CustomerDict) Then
        ' Fallback if no matching customer was found
        Call RetrieveCustomerFromUserform(CustomerDict)
    End If
    
    Call m_Logger.LogMessage("SetFromDictionary", LogLevelEnum.LogDebug, "Finished initializing customer object.")
    
End Sub
Private Function SetCustomer(CustomerDict As Scripting.Dictionary) As Boolean
    Dim Key As Variant
    Dim CustomerCollection As Collection
    Dim MatchedCustomer As clsTestingRequestCustomer
    
    ' Search for a matching customer in the dictionary
    For Each Key In CustomerDict
        ' Validate that the value is a collection of customers
        If TypeOf CustomerDict(Key) Is Collection Then
            Set CustomerCollection = CustomerDict(Key)
            
            ' Get the matching customer object
            Set MatchedCustomer = FindCustomerInCollection(CustomerCollection)
            
            If Not MatchedCustomer Is Nothing Then
                ' Copy all relevant properties
                m_CustomerID = MatchedCustomer.CustomerID
                m_CustomerName = MatchedCustomer.CustomerName
                m_StreetAddress = MatchedCustomer.StreetAddress
                m_City = MatchedCustomer.City
                m_State = MatchedCustomer.State
                m_PostalCode = MatchedCustomer.PostalCode
                m_Country = MatchedCustomer.Country

                SetCustomer = True
                Call m_Logger.LogMessage("SetFromDictionary", LogLevelEnum.LogInfo, "Found the matching customer object in the dictionary and successfully set customer name as " & m_CustomerName & ".")
                Exit Function
            End If
        Else
            ' Handle unexpected data types in the dictionary
            Call m_Logger.LogMessage("SetCustomer", LogLevelEnum.LogWarning, "Unexpected value in CustomerDict for Key: " & Key & ", Type: " & TypeName(CustomerDict(Key)))
        End If
    Next Key
    
    ' Return false if no match is found
    SetCustomer = False
    
End Function
Private Function FindCustomerInCollection(CustomerCollection As Collection) As clsTestingRequestCustomer
    Dim Customer As clsTestingRequestCustomer
    
    ' Inner loop through the collection
    For Each Customer In CustomerCollection
        If Customer.NameOnForm = m_NameOnForm And Customer.AddressOnForm = m_AddressOnForm Then
            Set FindCustomerInCollection = Customer
            Exit Function
        End If
    Next Customer
    
    ' Return nothing if no match is found
    Set FindCustomerInCollection = Nothing
End Function
Private Sub RetrieveCustomerFromUserform(CustomerDict As Scripting.Dictionary, Optional IsNewCustomer As Boolean = False)

    Dim UF As frmGetCustomer
    Dim ExistingCustomerButton As Boolean
    Dim UFCancelled As Boolean
    Dim CustomerCollection As Collection
    Dim Customer As clsTestingRequestCustomer
    Dim DBCustomerID As Long
    Dim SQLQuery As String
    
    ' Initialize the userform
    Set UF = New frmGetCustomer
    Call UF.Initialize(m_NameOnForm, m_AddressOnForm, CustomerDict, m_Logger)
    If IsNewCustomer Then
        UF.lblCaption.Caption = "Please enter the details for the new customer"
        UF.btnExistingCustomer.Enabled = False
        UF.btnNewCustomer.Value = True
    End If
    UF.Show

    ' Retrieve variables from the userform
    m_CustomerName = UF.CustomerName
    ExistingCustomerButton = UF.btnExistingCustomer.Value
    m_StreetAddress = UF.StreetAddress
    m_City = UF.City
    m_State = UF.State
    m_PostalCode = UF.PostalCode
    m_Country = UF.Country
    If IsNewCustomer Then
        m_NameOnForm = m_CustomerName
        m_AddressOnForm = FormatAddress(" ")
    End If
    
    UFCancelled = UF.Cancelled
    
    ' Unload the userform
    Unload UF
      
    ' Raise an error if the user cancelled the form
    If UFCancelled Then
        Err.Raise 900, "GetCustomer", "User cancelled the Get Customer userform"
    End If
    
    
    If ExistingCustomerButton Or CustomerDict.Exists(m_CustomerName) Then
        m_CustomerID = GetExistingCustomerID(CustomerDict)
    Else
        m_CustomerID = InsertNewCustomer(CustomerDict)
    End If
    
    Call InsertFormCustomer
    
End Sub
Private Function GetExistingCustomerID(CustomerDict As Scripting.Dictionary) As Long
    Dim CustomerCollection As Collection
    Dim Customer As clsTestingRequestCustomer
    
    ' Retrieve the ID for the database customer name
    If CustomerDict.Exists(m_CustomerName) Then
        If TypeOf CustomerDict(m_CustomerName) Is Collection Then
        
            Set CustomerCollection = CustomerDict(m_CustomerName)
            Set Customer = CustomerCollection(1)
            GetExistingCustomerID = Customer.CustomerID
        
            ' Add the customer to the current collection in the dictionary
            CustomerCollection.Add Me
            
        End If
        
    Else
        ' Handle case where customer is not in the dictionary
        Err.Raise 901, "GetCustomer", "Customer not found in dictionary"
    End If
End Function

Private Function InsertNewCustomer(CustomerDict As Scripting.Dictionary) As Long
    Dim SQLQuery As String
    Dim DBCustomerID As Long
    Dim CustomerCollection As New Collection
    
    ' Add a new customer to the database
    SQLQuery = "INSERT INTO customers (database_entry_date, customer_name, street_address, city, state, postal_code, country) VALUES ('" & Now & "', '" & m_CustomerName & "', '" & m_StreetAddress & "', '" & m_City & "', '" & m_State & "', '" & m_PostalCode & "', '" & m_Country & "')"
    DBCustomerID = m_AccessDB.ExecuteNonQuery(SQLQuery)
    
    ' Add a new collection to the dictionary with the customer object in it
    CustomerCollection.Add Me
    CustomerDict.Add m_CustomerName, CustomerCollection
    
    ' Return the ID to the calling sub
    InsertNewCustomer = DBCustomerID
End Function
Private Sub InsertFormCustomer()
    Dim SQLQuery As String
    ' Link the "Form Customer Name" to the actual customer entry in the database
    SQLQuery = "INSERT INTO form_customers (customer_id, database_entry_date, form_customer, form_address) VALUES ('" & m_CustomerID & "', '" & Now & "', '" & m_NameOnForm & "', '" & m_AddressOnForm & "')"
    Call m_AccessDB.ExecuteNonQuery(SQLQuery)
End Sub
Public Sub InitializeNewCustomer(Logger As clsLoggingSystem, AccessDB As clsAccessDatabase, SLIMDictionaryManager As clsSLIMDictionaryManager)
    Call Logger.LogMessage("InitializeNewChemical", LogLevelEnum.LogDebug, "Initializing new chemical.")
    
    ' Initialize properties
    Set m_Logger = Logger
    Set m_AccessDB = AccessDB
    
    Call RetrieveCustomerFromUserform(SLIMDictionaryManager.CustomerDict, True)
    
    Call Logger.LogMessage("InitializeNewChemical", LogLevelEnum.LogInfo, "Initializing new chemical.")
End Sub
Public Sub LoadFromJson(Json As Dictionary, Logger As clsLoggingSystem)
    Call Logger.LogMessage("clsTestingRequestCustomer.LoadFromJson", LogLevelEnum.LogDebug, "Loading data from Json to Customer Object")
    
    If Not Logger.DebugMode Then On Error GoTo ErrorHandler
    
    Set m_Logger = Logger
    
    m_CustomerID = Json("id")
    If Json.Exists("customer") Then
        m_CustomerName = Json("name")
        m_StreetAddress = Json("street_address")
        m_City = Json("city")
        m_State = Json("state")
        m_PostalCode = Json("postal_code")
        m_Country = Json("country")
    End If
    m_NameOnForm = Json("name")
    m_AddressOnForm = Json("address1") & " " & Json("address2")
    
    Call Logger.LogMessage("clsTestingRequestCustomer.LoadFromJson", LogLevelEnum.LogDebug, "Completed loading data from Json to Customer Object")
    Exit Sub
ErrorHandler:
    Dim ErrorMessage As String
    ErrorMessage = "Error encountered in clsTestingRequestCustomer.LoadFromJson where the customer is: " & m_CustomerName
    Call Logger.LogMessage("clsTestingRequestCustomer.LoadFromJson", LogLevelEnum.LogError, ErrorMessage)
    Err.Raise vbObjectError + 1000, "clsTestingRequestCustomer.LoadFromJson", ErrorMessage
End Sub
