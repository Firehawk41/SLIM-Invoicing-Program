VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoicePriceDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoicePriceDetail
' Version: 1.0.0
' Author: JT
' Created: 2025-07-01
' LastModified: 2025-07-01
' Description: Stores pricing information
' DependsOn: clsLoggingSystem
' ChangeLog:
'   - 1.0.0 - 2025-07-01 - Initial release
Private m_Price As Long
Private m_IsFlatRate As Boolean
Public Property Let Price(Value As Long)
    m_Price = Value
End Property
Public Property Let IsFlatRate(Value As Boolean)
    m_IsFlatRate = Value
End Property
Public Property Get Price() As Long
    Price = m_Price
End Property
Public Property Get IsFlatRate() As Boolean
    IsFlatRate = m_IsFlatRate
End Property

