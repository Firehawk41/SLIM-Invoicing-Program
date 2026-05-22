Attribute VB_Name = "modEnums"
Option Explicit
' EnumModule
' MacroName: modEnums
' Version: 1.0.6
' Author: JT
' Created: 2025-01-15
' LastModified: 2025-05-21
' Description: Store public enums in one place
' DependsOn:
' ChangeLog:
'   - 1.0.6 - 2025-05-21 added metadata JT
Public Enum InvoiceTypeEnum
    DefaultIndividual = 1
    CustomerASummary
    CustomerBSummary
End Enum
Public Enum LineItemColumnTypeEnum
    LogInDateColumn = 1
    TestingRequestColumn
    QuantityColumn
    DescriptionColumn
    TurnAroundTimeColumn
    UnitPriceColumn
    ExtendedPriceColumn
    RequestedByColumn
End Enum
Public Enum LogLevelEnum
    LogDebug = 1
    LogInfo
    LogWarning
    LogError
End Enum
Public Enum TestingRequestTypeEnum
    Chemical = 1
    Water
    Wafer
End Enum
Public Enum EmailTypeEnum
    ResultsMain = 1
    ResultsCC
    InvoiceMain
    InvoiceCC
End Enum
Public Enum ProcessingTimeEnum
    ExtendedTime = 1
    NextDay
    TimeLimited
    SameDayRush
    CallInRush
End Enum
Public Enum PaymentTypeEnum
    PONumber = 1
    CreditCard
End Enum
Public Enum PythonModeEnum
    ModePrototype = 0
    ModeProduction
End Enum
