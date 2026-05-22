Attribute VB_Name = "modConfig"
Option Explicit
' ================================================================
' Name: modConfig.template.bas
' Author: JT
' Created: 2026-05-08
' Description: Template for environment configuration constants.
'              Copy this file to modConfig.bas, populate with
'              local values, and ensure modConfig.bas remains
'              excluded from source control via .gitignore.
' DependsOn:
' ChangeLog:
'   - 1.0.0 - 2026-05-08 - Initial release, split from
'                           modInvoiceSystem
'
' ================================================================
' SETUP INSTRUCTIONS:
'   1. Copy this file to modConfig.bas in the same directory
'   2. Replace all placeholder values below with local paths
'   3. Confirm modConfig.bas is listed in .gitignore
'   4. Do not commit modConfig.bas to source control
' ================================================================

' Path to the Access database (.accdb) on your network or local drive
Public Const DB_PATH As String = "\\YOUR-SERVER\YourShare\Path\To\Database.accdb"

' Path to the output folder where generated CSV invoices are saved
Public Const OUTPUT_PATH As String = "\\YOUR-SERVER\YourShare\Path\To\Output\Folder\"

' Set to True during development to disable error handling and expose
' full stack traces. Set to False for production use.
Public Const DEBUG_MODE As Boolean = False