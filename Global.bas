Attribute VB_Name = "Global"
Option Explicit

Public gclsSQLServer As MES.SQLConnection
Public gclsMESApplication As MES.MESApplication

Public gstrApplicationRole As String
Public gstrRolePassword As String
Public gstrINIFile As String
Public gblnUpdate As Boolean
Public gstrDivisionPrefix As String

Public gblnFromAppMenu As Boolean

Public gconDatabase As ADODB.Connection

Public garrstrPart() As String
Public garrstrPartDescription() As String

Public gblnBuiltPart As Boolean

Public gblnFindEcnInfo As Boolean
Public gblnFindEcnInfoCancel As Boolean

Public gstrCurrentDate As String

Public gblnMaintPassedUpdates As Boolean

Public gblnLine As String

