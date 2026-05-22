VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAccessDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsAccessDatabase
' Version: 1.1.1
' Author: JT
' Created: 2025-??-??
' LastModified: 2025-06-04
' Description: Class used to access an Access Database by dependency injection
' DependsOn: clsLoggingSystem
' ChangeLog:
'   - 1.1.1 - 2025-05-28    - added metadata
'   - 1.1.2 - 2025-06-04    - added enum reference in GetTableFieldNames
Private m_DatabasePath As String
Private m_Connection As ADODB.Connection
Private m_Logger As clsLoggingSystem
Private m_TransactionActive As Boolean
Private Enum SQLResultProcessorEnum
    Query = 1
    Scalar
End Enum
' Initializes the Access database handler
Public Sub Initialize(DatabasePath As String, Optional Logger As clsLoggingSystem = Nothing)
    m_DatabasePath = DatabasePath
    Set m_Logger = Logger
End Sub

' Opens the database connection
Private Sub OpenConnection(Optional BeginTransaction As Boolean = False)
    If m_Connection Is Nothing Then
        Set m_Connection = New ADODB.Connection
        m_Connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & m_DatabasePath
    End If

    ' Begins the transaction
    If BeginTransaction And Not m_TransactionActive Then
        m_Connection.BeginTrans
        m_TransactionActive = True
    End If
End Sub

' Closes the database connection
Private Sub CloseConnection()
    If Not m_Connection Is Nothing And Not m_TransactionActive Then
        If m_Connection.State = adStateOpen Then m_Connection.Close
        Set m_Connection = Nothing
    End If
End Sub
' Centralized method to execute any query and handle the result processing
Private Function ExecuteSQL(SQLQuery As String, ResultProcessor As SQLResultProcessorEnum, BeginTransaction As Boolean) As Variant
    Dim Rs As ADODB.Recordset
    Dim Result As Variant

    On Error GoTo ErrorHandler

    ' Open the database connection
    Call OpenConnection(BeginTransaction)

    ' Execute the SQL query
    Set Rs = New ADODB.Recordset
    Rs.Open SQLQuery, m_Connection

    ' Process the result based on the ResultProcessor type
    If ResultProcessor = SQLResultProcessorEnum.Query Then
        ' For multiple rows (2D array)
        If Not Rs.EOF And Not Rs.BOF Then
            Result = Rs.GetRows ' Retrieve the data as a 2D array
        Else
            Result = Array() ' No data found, return empty array
        End If
    ElseIf ResultProcessor = SQLResultProcessorEnum.Scalar Then
        ' For single scalar value
        If Not Rs.EOF And Not Rs.BOF Then
            Result = Rs.Fields(0).Value ' Retrieve the first field's value from the first row
        Else
            Result = Null ' No result found, return Null
        End If
    End If

    ' Log success
    LogMessage "ExecuteSQL", LogLevelEnum.LogInfo, "Query executed successfully: " & SQLQuery

    ' Clean up
    If Rs.State = adStateOpen Then Rs.Close
    Set Rs = Nothing
    CloseConnection

    ' Return the result
    ExecuteSQL = Result
    Exit Function

ErrorHandler:
    ' Log error
    LogMessage "ExecuteSQL", LogLevelEnum.LogError, "Error executing query: " & SQLQuery & ". Error: " & Err.Description
    ExecuteSQL = Null ' Return Null or empty array on error
    ' Rollback the transaction if active
    If m_TransactionActive Then
        m_Connection.Rollback
        m_TransactionActive = False
    End If

    ' Close the connection
    CloseConnection
End Function

' Wrapper for ExecuteSQL for executing a query that returns multiple rows
Public Function ExecuteQuery(SQLQuery As String, Optional BeginTransaction As Boolean) As Variant
    ExecuteQuery = ExecuteSQL(SQLQuery, SQLResultProcessorEnum.Query, BeginTransaction)
End Function

' Wrapper for ExecuteSQL for executing a query that returns a single scalar value
Public Function ExecuteScalarQuery(SQLQuery As String, Optional BeginTransaction As Boolean) As Variant
    ExecuteScalarQuery = ExecuteSQL(SQLQuery, SQLResultProcessorEnum.Scalar, BeginTransaction)
End Function


' Executes a non-query SQL statement (e.g., INSERT, UPDATE, DELETE)
Public Function ExecuteNonQuery(SQLStatement As String, Optional BeginTransaction As Boolean) As Long
    On Error GoTo ErrorHandler

    ' Open the database connection
    Call OpenConnection(BeginTransaction)

    ' Execute the statement
    m_Connection.Execute SQLStatement

    ' Retrieve the last inserted ID (assuming it's an AutoNumber/AutoIncrement field in the DB)
    Dim Rs As Object
    Set Rs = m_Connection.Execute("SELECT @@IDENTITY AS NewID")

    ' Log success
    LogMessage "ExecuteNonQuery", LogLevelEnum.LogInfo, "Statement executed successfully: " & SQLStatement

    ' Return the ID
    ExecuteNonQuery = Rs.Fields("NewID").Value

    Rs.Close
    Set Rs = Nothing

    CloseConnection
    Exit Function

ErrorHandler:
    ' Log error
    LogMessage "ExecuteNonQuery", LogLevelEnum.LogError, "Error executing statement: " & SQLStatement & ". Error: " & Err.Description
    ExecuteNonQuery = -1

    ' Rollback the transaction if active
    If m_TransactionActive Then
        RollbackTransaction
        m_TransactionActive = False
    End If

    ' Close the connection
    CloseConnection
End Function

' Retrieves field names from the specified table
Public Function GetTableFieldNames(TableName As String) As Variant
    Dim Rs As ADODB.Recordset
    Dim fieldNames As Variant
    Dim i As Long

    On Error GoTo ErrorHandler

    ' Open the database connection
    OpenConnection

    ' Get field names
    Set Rs = New ADODB.Recordset
    Rs.Open "SELECT TOP 1 * FROM " & TableName, m_Connection, adOpenKeyset, adLockOptimistic

    ' Store field names in an array
    ReDim fieldNames(1 To Rs.Fields.Count)
    For i = 1 To Rs.Fields.Count
        fieldNames(i) = Rs.Fields(i - 1).Name
    Next i

    ' Log success
    LogMessage "GetTableFieldNames", LogLevelEnum.LogInfo, "Retrieved field names from table: " & TableName

    ' Return field names
    GetTableFieldNames = fieldNames

    ' Clean up
    If Rs.State = adStateOpen Then Rs.Close
    Set Rs = Nothing
    CloseConnection
    Exit Function

ErrorHandler:
    ' Log error
    LogMessage "GetTableFieldNames", LogLevelEnum.LogError, "Error retrieving field names from table: " & TableName & ". Error: " & Err.Description
    GetTableFieldNames = Array() ' Return empty array on error
    CloseConnection
End Function

' Logs a message using the logger if it is set
Private Sub LogMessage(Source As String, Level As LogLevelEnum, Message As String)
    If Not m_Logger Is Nothing Then
        Call m_Logger.LogMessage(Source, Level, Message)
    End If
End Sub
' Public sub to commit a transaction
Public Sub CommitTransaction()
    On Error GoTo ErrorHandler
    If m_TransactionActive Then
        m_Connection.CommitTrans
        m_TransactionActive = False
        LogMessage "CommitTransaction", LogLevelEnum.LogInfo, "Transaction committed successfully."
    Else
        LogMessage "RollbackTransaction", LogLevelEnum.LogInfo, "Transaction was not open."
    End If
    Exit Sub

ErrorHandler:
    LogMessage "CommitTransaction", LogLevelEnum.LogError, "Error committing transaction: " & Err.Description
End Sub
' Public sub to rollback a transaction
Public Sub RollbackTransaction()
    On Error GoTo ErrorHandler
    If m_TransactionActive Then
        m_Connection.RollbackTrans
        m_TransactionActive = False
        LogMessage "RollbackTransaction", LogLevelEnum.LogInfo, "Transaction rolled back successfully."
    Else
        LogMessage "RollbackTransaction", LogLevelEnum.LogInfo, "Transaction was not open."
    End If
    Exit Sub

ErrorHandler:
    LogMessage "RollbackTransaction", LogLevelEnum.LogError, "Error rolling back transaction: " & Err.Description
End Sub

' Destructor to ensure the connection is closed
Private Sub Class_Terminate()
    If m_TransactionActive Then
        RollbackTransaction
    End If
    CloseConnection
End Sub
