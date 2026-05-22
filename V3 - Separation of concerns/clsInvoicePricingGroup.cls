VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInvoicePricingGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' MacroName: clsInvoicePricingEngine
' Author: JT
' Created: 2026-03-27
' Description: Container to store properties for analyses grouped together for building a line item
' DependsOn:
' ChangeLog:
'   - 0.1.0 - 2026-03-27 - Initial skeleton

Public Quantity As Long
Public Samples As Collection
Public Analyses As Scripting.Dictionary
Public ProcessingTime As Long
Public CustomerID As Long
Public Price As Currency
