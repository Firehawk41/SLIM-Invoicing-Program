Attribute VB_Name = "modEnums"
Option Explicit
' ================================================================
' Name: modEnums
' Author: JT
' Created: 2025-05-21
' Description: Public enums used across the invoice system.
'              All enum definitions live here so dependent modules
'              have a single import target.
' DependsOn:
' ChangeLog:
'   - 1.0.6 - 2025-05-21 - Added metadata
'   - 1.1.0 - 2025-11-18 - Added InvoiceOutputEnum
'   - 1.2.0 - 2026-01-22 - Added InvoiceSelectionEnum
'   - 1.3.0 - 2026-03-25 - Added ProcessingTimeEnum values
'   - 2.0.0 - 2026-05-08 - V3 port: removed InvoiceOutputEnum,
'                           InvoiceTypeEnum, LineItemColumnTypeEnum,
'                           EmailTypeEnum, PythonModeEnum
'
' ================================================================
' ENUMS IN THIS MODULE:
'
' InvoiceSelectionEnum  — Individual or Batch invoice run mode.
'                         Used by frmInvoiceTypeSelector and
'                         modInvoiceSystem.
'
' LogLevelEnum          — Log severity levels. Used by
'                         clsLoggingSystem and all classes that
'                         call LogMessage/LogError.
'
' TestingRequestTypeEnum — Sample request category. Used by
'                          clsTRSubmission and the SO pipeline.
'
' ProcessingTimeEnum    — Turnaround time options per sample.
'                         Used by clsTRSubmission,
'                         clsSalesOrderPricingEngine, and
'                         modUtilities.
'
' PaymentTypeEnum       — Payment method on a submission. Used
'                         by clsTRSubmission.
' ================================================================
Public Enum InvoiceSelectionEnum
    Individual = 1
    Batch
End Enum

Public Enum LogLevelEnum
    LogDebug = 1
    LogInfo
    LogWarning
    LogError
End Enum

Public Enum TestingRequestTypeEnum
    RequestType_Min = 1
    Chemical = 1
    Water
    Wafer
    RequestType_Max = 3
End Enum

Public Enum ProcessingTimeEnum
    ProcessingTime_Min = 1
    ExtendedTime = 1
    NextDay
    TimeLimited
    SameDayRush
    CallInRush
    TwoDays
    ThreeDays
    FiveDays
    ProcessingTime_Max = 8
End Enum

Public Enum PaymentTypeEnum
    PONumber = 1
    CreditCard
End Enum
