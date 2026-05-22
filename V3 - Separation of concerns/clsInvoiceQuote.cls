VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoiceQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoiceQuote
' Version: 1.0.1
' Author: JT
' Created: ?
' LastModified: 2025-06-14
' Description: Object to hold quotes for customers
' DependsOn: clsLoggingSystem, modEnums
' ChangeLog:
'   - 1.0.1 - 2025-06-11 - Added metadata
Private m_Logger As clsLoggingSystem
Private m_PriceIndex As Variant
Private m_Analysis As String
Private m_RequestType As TestingRequestTypeEnum
Private m_ExtendedTimePrice As Long
Private m_NextDayPrice As Long
Private m_TimeLimitedPrice As Long
Private m_SameDayRushPrice As Long
Private m_CallInRushPrice As Long
Public Property Get PriceIndex() As Variant
    PriceIndex = m_PriceIndex
End Property
Public Property Get Analysis() As String
    Analysis = m_Analysis
End Property
Public Property Get NextDayPrice() As Long
    NextDayPrice = m_NextDayPrice
End Property
Public Property Get TimeLimitedPrice() As Long
    TimeLimitedPrice = m_TimeLimitedPrice
End Property
Public Property Get SameDayRushPrice() As Long
    SameDayRushPrice = m_SameDayRushPrice
End Property
Public Property Get CallInRushPrice() As Long
    CallInRushPrice = m_CallInRushPrice
End Property
Public Sub Initialize(RequestType As TestingRequestTypeEnum, PriceIndex As Variant, Analysis As String, ExtendedTimePrice As Long, NextDayPrice As Long, NextDayTimeLimitedPrice As Long, SameDayRushPrice As Long, CallInRushPrice As Long, Logger As clsLoggingSystem)
    Call Logger.LogMessage("Initialize", LogLevelEnum.LogDebug, "Initializing quote object.")

    Set m_Logger = Logger
    m_RequestType = RequestType
    m_PriceIndex = PriceIndex
    m_Analysis = Analysis
    m_ExtendedTimePrice = ExtendedTimePrice
    m_NextDayPrice = NextDayPrice
    m_TimeLimitedPrice = TimeLimitedPrice
    m_SameDayRushPrice = SameDayRushPrice
    m_CallInRushPrice = CallInRushPrice

    Call m_Logger.LogMessage("Initialize", LogLevelEnum.LogDebug, "Successfully created " & m_PriceIndex & " " & m_Analysis & " quote.")

End Sub
Public Function GetUnitPrice(ProcessingTime As ProcessingTimeEnum) As Long

    Select Case ProcessingTime
        Case ProcessingTimeEnum.ExtendedTime
            GetUnitPrice = m_ExtendedTimePrice
        Case ProcessingTimeEnum.NextDay
            GetUnitPrice = m_NextDayPrice
        Case ProcessingTimeEnum.TimeLimited
            GetUnitPrice = m_TimeLimitedPrice
        Case ProcessingTimeEnum.SameDayRush
            GetUnitPrice = m_SameDayRushPrice
        Case ProcessingTimeEnum.CallInRush
            GetUnitPrice = m_CallInRushPrice
        Case Else
            Call m_Logger.LogMessage("GetUnitPrice", LogLevelEnum.LogError, "Invalid Processing Time")
    End Select

    Call m_Logger.LogMessage("GetUnitPrice", LogLevelEnum.LogInfo, "Successfully retrieved price: $" & GetUnitPrice & ".")

End Function
