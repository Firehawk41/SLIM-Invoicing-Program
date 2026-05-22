Attribute VB_Name = "modUtilities"
Option Explicit
' ================================================================
' Name: modUtilities
' Author: JT
' Created: 2026-01-22
' Description: Shared utility functions used across the codebase.
'              No business logic. No domain dependencies.
' DependsOn: modEnums
' ChangeLog:
'   - 1.0.0 - 2026-01-22 - Initial release
'   - 1.1.0 - 2026-04-28 - Added ProcessingTime conversion helpers
'   - 1.2.0 - 2026-05-19 - Added CloneCollection
'
' ================================================================
' FUNCTIONS IN THIS MODULE:
'
' NullToDefault(Value, Default)
'   Coalesces a Null variant to a default value. Applied at the
'   DB boundary in all repository classes.
'
' ProcessingTimeToEnum(V)
'   Converts Excel form string values to ProcessingTimeEnum.
'   Raises on unrecognised values — caller must handle.
'
' ProcessingTimeToString(V)
'   Converts ProcessingTimeEnum to display string for CSV output
'   and line item description construction.
'
' ProcessingTimeToDays(PT)
'   Converts ProcessingTimeEnum to integer day count.
'   Used for turnaround time calculations.
'
' CloneCollection(Source)
'   Returns a shallow copy of a scalar Collection.
'   NOT safe for object collections — asserts on object items.
' ================================================================

Public Function NullToDefault(Value As Variant, Default As Variant) As Variant
    If IsNull(Value) Then
        NullToDefault = Default
    Else
        NullToDefault = Value
    End If
End Function

Public Function ProcessingTimeToEnum(V As String) As ProcessingTimeEnum
    Select Case LCase$(Trim$(V))
        Case "extended time": ProcessingTimeToEnum = ProcessingTimeEnum.ExtendedTime
        Case "next day", "next day rush": ProcessingTimeToEnum = ProcessingTimeEnum.NextDay
        Case "time limited", "timelimited", "next day time limited": ProcessingTimeToEnum = ProcessingTimeEnum.TimeLimited
        Case "same day rush", "samedayrush": ProcessingTimeToEnum = ProcessingTimeEnum.SameDayRush
        Case "call in rush", "callinrush": ProcessingTimeToEnum = ProcessingTimeEnum.CallInRush
        Case "two days", "2 days", "2days": ProcessingTimeToEnum = ProcessingTimeEnum.TwoDays
        Case "three days", "3 days", "3days", "up to 3 working days": ProcessingTimeToEnum = ProcessingTimeEnum.ThreeDays
        Case "five days", "5 days", "5days": ProcessingTimeToEnum = ProcessingTimeEnum.FiveDays
        Case Else
            Err.Raise vbObjectError + 500, "modUtilities.ProcessingTimeToEnum", _
                      "Unrecognised processing time string: '" & V & "'"
    End Select
End Function

Public Function ProcessingTimeToString(V As ProcessingTimeEnum) As String
    Select Case V
        Case ProcessingTimeEnum.ExtendedTime: ProcessingTimeToString = "Extended Time"
        Case ProcessingTimeEnum.NextDay: ProcessingTimeToString = "Next Day"
        Case ProcessingTimeEnum.TimeLimited: ProcessingTimeToString = "Time Limited"
        Case ProcessingTimeEnum.SameDayRush: ProcessingTimeToString = "Same Day Rush"
        Case ProcessingTimeEnum.CallInRush: ProcessingTimeToString = "Call In Rush"
        Case ProcessingTimeEnum.TwoDays: ProcessingTimeToString = "Two Days"
        Case ProcessingTimeEnum.ThreeDays: ProcessingTimeToString = "Three Days"
        Case ProcessingTimeEnum.FiveDays: ProcessingTimeToString = "Five Days"
        Case Else
            Err.Raise vbObjectError + 501, "modUtilities.ProcessingTimeToString", _
                      "Unrecognised ProcessingTimeEnum value: " & V
    End Select
End Function

Public Function ProcessingTimeToDays(PT As ProcessingTimeEnum) As Long
    Select Case PT
        Case ProcessingTimeEnum.SameDayRush, ProcessingTimeEnum.CallInRush
            ProcessingTimeToDays = 0
        Case ProcessingTimeEnum.NextDay, ProcessingTimeEnum.TimeLimited
            ProcessingTimeToDays = 1
        Case ProcessingTimeEnum.TwoDays
            ProcessingTimeToDays = 2
        Case ProcessingTimeEnum.ExtendedTime, ProcessingTimeEnum.ThreeDays
            ProcessingTimeToDays = 3
        Case ProcessingTimeEnum.FiveDays
            ProcessingTimeToDays = 5
        Case Else
            ProcessingTimeToDays = 1
    End Select
End Function

Public Function CloneCollection(Source As Collection) As Collection
    ' for scalar collections only - NOT collections of objects
    Dim Result As New Collection
    Dim Item As Variant
    For Each Item In Source
        Debug.Assert Not IsObject(Item)
        Result.Add Item
    Next Item
    Set CloneCollection = Result
End Function
