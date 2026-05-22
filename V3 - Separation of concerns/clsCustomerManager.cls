VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsCustomerManager
' Version: 0.1.0
' Author: JT
' Created: 2025-12-15
' LastModified: 2025-12-15
' Description: Controller object for managing Customers
' DependsOn: clsLoggingSystem, modEnums
' ChangeLog:
'   - 0.1.0 - 2025-12-15 - Initial prototype
Private m_Logger As clsLoggingSystem
Private m_AccessDB As clsAccessDatabase
Private m_Repo As clsCustomerRepository
' ------------------
' Public Properties
' ------------------

' ------------------
' Initialization
' ------------------
Public Sub Initialize(SLIMDict As clsSLIMDictionaryManager, AccessDatabase As clsAccessDatabase, Logger As clsLoggingSystem)
    Call Logger.LogMessage("clsCustomerManager.Initialize", LogLevelEnum.LogDebug, "Initializing Customer Manager")

    Set m_Logger = Logger
    Set m_AccessDatabase = AccessDatabase
    Set m_SLIMDict = SLIMDict
    Set m_Repo = New clsCustomerRepository
    Call m_Repo.Initialize(AccessDatabase, Logger)

    Call Logger.LogMessage("clsCustomerManager.Initialize", LogLevelEnum.LogInfo, "Completed initializing Customer Manager")
End Sub
