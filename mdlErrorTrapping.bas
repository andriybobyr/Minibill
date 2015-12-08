Attribute VB_Name = "mdlErrorTrapping"
Option Explicit

'* Not that this couldn't all be wrapped in a Class module,
'* but for simplicity, it's been created as a standard module.



Public Sub ShowError(strModule As String, _
        strProcedure As String, _
        lngErrorNumber As Long, _
        strErrorDescription As String)
        
   '* Purpose  :  Provide a central error handling mechanism.
   On Error GoTo PROC_ERR
   Dim strMessage As String
   Dim strCaption As String
   Dim intLogFile As Integer
   Dim strErrorLogFileName As String
   
   strErrorLogFileName = App.Path & "\Errors.txt"

   '* Obtain a free file handle.
   intLogFile = FreeFile
   
   '* Open the error log text file in Append mode.
   '* If the file doesn't exist, the Open statement
   '* creates it.
   Open strErrorLogFileName For Append As #intLogFile
   
   '* Write the header.
   Print #intLogFile, "*** Error Encountered " & VBA.Now & " ***"

   '* Write the pertinent error information to the log file.
   Print #intLogFile, "Error: " & lngErrorNumber
   Print #intLogFile, "Description: " & strErrorDescription
   Print #intLogFile, "Procedure: " & strProcedure
   Print #intLogFile, "Module: " & strModule
   
   '* Write a blank line to the log file.
   Print #intLogFile, ""
   
   '* Close the error log text file.
   Close #intLogFile

   '* Build the error message for display to the user.
   strMessage = "Error: " & strErrorDescription & vbCrLf & vbCrLf & _
                "Module: " & strModule & vbCrLf & _
                "Procedure: " & strProcedure & vbCrLf & vbCrLf & _
                "Please notify the HelpDesk at " & _
                gclsMESApplication.DivisionHelpDesk & _
                " about this issue..HelpDeskPhone" & vbCrLf & _
                "Please provide the support technician with " & _
                "information shown in " & vbCrLf & "this dialog " & _
                "box, as well as an explanation of what you " & _
                "were" & vbCrLf & "doing when this " & _
                "error occurred."

   '* Build the caption for the message box. The caption shows
   '* the version number of the program.
   strCaption = "Unexpected Error! Version: " & _
                Trim(Str$(App.Major)) & "." & Trim(Str$(App.Minor)) & "." & _
                Format(App.Revision, "0000")

   MsgBox strMessage, vbCritical, strCaption

PROC_EXIT:
   Exit Sub
   
PROC_ERR:
   Resume Next
   
End Sub
