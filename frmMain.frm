VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mini-Bill Of Material Main Menu"
   ClientHeight    =   5520
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6840
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   6840
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cdlgFile 
      Left            =   420
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   1680
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   2055
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   2760
      Width           =   3675
   End
   Begin VB.Label lblDivision 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   4230
   End
   Begin VB.Label lblMain 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mini-Bill Of Material"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1275
      Left            =   1020
      TabIndex        =   0
      Top             =   900
      Width           =   4890
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileChangePassword 
         Caption         =   "&Change Password"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "&Mini-Bill Setup"
      Begin VB.Menu mnuSetupNewModel 
         Caption         =   "&New Model Setup"
      End
      Begin VB.Menu mnuSetupModelLineToLineCopy 
         Caption         =   "Model &Line To Line Copy"
      End
      Begin VB.Menu mnuSetupSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetupNewPartsReview 
         Caption         =   "New Parts &Review"
      End
      Begin VB.Menu mnuSetupModelsByPart 
         Caption         =   "&Assign Default Parts To Models"
      End
      Begin VB.Menu mnuSetupECN 
         Caption         =   "&ECN Review"
      End
      Begin VB.Menu mnuTemporaryNewPart 
         Caption         =   "&Temporary Part Setup"
      End
      Begin VB.Menu mnuSetupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetupPartLocation 
         Caption         =   "&Part / Location Setup"
      End
      Begin VB.Menu mnuModelLineStockingLocationNotes 
         Caption         =   "Model &Notes Maintenance"
      End
      Begin VB.Menu mnuSetupMiniBill 
         Caption         =   "&Mini-Bill Maintenance"
      End
      Begin VB.Menu mnuSetupSplitPart 
         Caption         =   "&Split Part Maintenance"
      End
      Begin VB.Menu mnuSetupSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetupPartCategory 
         Caption         =   "Part / &Category Setup"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportsDailySchedule 
         Caption         =   "&Daily Schedule/Process Sheet"
      End
      Begin VB.Menu mnuReportRackPickList 
         Caption         =   "&Rack Pick List"
      End
      Begin VB.Menu mnuTempPartsList 
         Caption         =   "&Temporary Parts List"
      End
      Begin VB.Menu mnuReportsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportsDailyScheduleAll 
         Caption         =   "&Minibill Schedule"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsArea 
         Caption         =   "&Area Maintenance"
      End
      Begin VB.Menu mnuToolsLocation 
         Caption         =   "&Location Maintenance"
      End
      Begin VB.Menu mnuToolsCategory 
         Caption         =   "&Category Maintenance"
      End
      Begin VB.Menu menuMNBOverrideECN 
         Caption         =   "&Override ECN Info "
      End
      Begin VB.Menu mnuMNBModelLineInactivity 
         Caption         =   "&Model Line Inactivity Flag Maintenance"
      End
      Begin VB.Menu mnuMNBModelLocationSubAssembly 
         Caption         =   "Model Sub-Assembly By Location Maintenance"
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsAreaLocation 
         Caption         =   "A&rea / Stocking Location Specification"
      End
      Begin VB.Menu mnuToolsLineLocation 
         Caption         =   "L&ine / Location Specification"
      End
      Begin VB.Menu mnuToolsLocationCategory 
         Caption         =   "L&ocation / Category Specification"
      End
      Begin VB.Menu mnuToolsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsUserSetup 
         Caption         =   "&User Setup"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "&Index..."
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrServer As String
Private mstrDatabase As String

Private mstrUser As String
Private mstrPassword As String
Private strCurrentMonth As String
Private strCurrentDay As String
Private strDisplay As String

Private Sub Form_Load()
    ' Purpose:  Show the form and login to the server
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim strCommandLine As String
    Dim intStartPos As Integer
    Dim intEndPos As Integer
    
    ' Show the form
    Me.Show
    DoEvents
    
    strCommandLine = Command()
    mstrUser = vbNullString
    mstrPassword = vbNullString
    If Len(strCommandLine) > 0 Then
        intStartPos = InStr(1, strCommandLine, "/u:")
        If intStartPos > 0 Then
            intStartPos = intStartPos + 3
            intEndPos = InStr(intStartPos, strCommandLine, " ")
            If intEndPos = 0 Then
                intEndPos = Len(strCommandLine)
            Else
                intEndPos = intEndPos - 1
            End If
            mstrUser = Mid(strCommandLine, intStartPos, intEndPos - intStartPos + 1)
        End If
        intStartPos = InStr(1, strCommandLine, "/p:")
        If intStartPos > 0 Then
            intStartPos = intStartPos + 3
            intEndPos = InStr(intStartPos, strCommandLine, " ")
            If intEndPos = 0 Then
                intEndPos = Len(strCommandLine)
            Else
                intEndPos = intEndPos - 1
            End If
            mstrPassword = Mid(strCommandLine, intStartPos, intEndPos - intStartPos + 1)
        End If
    End If
    If Len(mstrUser) = 0 Then
        gblnFromAppMenu = False
    Else
        gblnFromAppMenu = True
    End If
    
    Call ProcessINIFile
    
    ' If the server or database were not found, rase an error
    If Len(mstrServer) = 0 Or Len(mstrDatabase) = 0 Then
        Err.Raise vbObjectError + 1000, "Form_Load", _
            "The INI file is not set up properly."
    End If
    
    ' Call the LoginToSQLServer function to attempt to log
    ' into SQL Server
    If Not LoginToSQLServer Then
        MsgBox "Login Failed", vbCritical + vbOKOnly, _
            "Login To SQL Server"
        Unload Me
        GoTo PROC_EXIT
    End If
    
    If gclsMESApplication.ApplicationRole = "MNB_Administration" Then
        mnuTools.Visible = True
        Me.mnuSetup.Visible = True
    ElseIf gclsMESApplication.ApplicationRole = "MNB_Update" Then
        mnuTools.Visible = False
        mnuSetup.Visible = True
    Else
        mnuTools.Visible = False
        mnuSetup.Visible = False
    End If
    
    lblDivision.Caption = gclsMESApplication.Division
    gstrDivisionPrefix = gclsMESApplication.DivisionPrefix
    
    If gstrDivisionPrefix = "FIN" Then
        mnuReportsDailySchedule.Visible = False
    End If
    
    If gstrDivisionPrefix = "CLY" Then
        mnuReportsDailySchedule.Visible = False
    End If
    
 'Create current date in string format to compare with the start/quit date
    If Month(Date) < 10 Then
        strCurrentMonth = "0" & Month(Date)
    Else
        strCurrentMonth = Month(Date)
    End If
    
    If Day(Date) < 10 Then
        strCurrentDay = "0" & Day(Date)
    Else
        strCurrentDay = Day(Date)
    End If
    
    gstrCurrentDate = Year(Date) & strCurrentMonth & strCurrentDay
    
    Set gconDatabase = gclsSQLServer.Connect( _
        gclsMESApplication.ApplicationRole, _
        gclsMESApplication.ApplicationPassword)
        
        strDisplay = gconDatabase.ConnectionTimeout
        
    If Not gconDatabase Is Nothing Then
        If gconDatabase.State = adStateOpen Then
            Dim rsServer As ADODB.Recordset
            Set rsServer = New ADODB.Recordset
            With rsServer
                Set .ActiveConnection = gconDatabase
                .CursorLocation = adUseClient
                .CursorType = adOpenForwardOnly
                .LockType = adLockReadOnly
                .Source = "select server_type_description " & _
                    "from Administration.dbo.v_adm_server " & _
                    "join Administration.dbo.v_adm_server_type  on " & _
                    "Administration.dbo.v_adm_server.server_type_id = Administration.dbo.v_adm_server_type.server_type_id " & _
                    "where server_name = '" & UCase(mstrServer) & "'"
                .Open
                If .RecordCount > 0 Then
                    lblDivision.Caption = Trim(lblDivision.Caption) & " - " & _
                        Trim(!server_type_description)
                End If
                .Close
            End With
            Set rsServer = Nothing
        End If
    End If


PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Load", Err.Number, _
        Err.Description)
    Unload Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Purpose:  Attempt to unload all open forms before
    '           unloading the current form.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim frm As Form
    
    For Each frm In Forms
        If Not frm Is Me Then
            Unload frm
        End If
    Next frm
    
    CloseHelp
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Cancel = True
    GoTo PROC_EXIT
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Purpose:  De-refence objects before unloading the form
    
    If Not gconDatabase Is Nothing Then
        If gconDatabase.State = adStateOpen Then
            gclsSQLServer.Disconnect gconDatabase
        End If
        Set gconDatabase = Nothing
    End If
    
    Set gclsSQLServer = Nothing
    Set gclsMESApplication = Nothing
    
End Sub

Function LoginToSQLServer() As Boolean
    ' Purpose:  Display the login form and attempt to log the
    '           user into the server.
    
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ' Instantiate site class and SQLConnect class
    Set gclsSQLServer = New MES.SQLConnection
    Set gclsMESApplication = New MES.MESApplication
    
    ' Initialize the application information
    With gclsMESApplication
        Set .CallingApplication = App
        .Initialize mstrServer
    End With
    
    ' Set properties of the class and attempt to log in by
    ' calling the .Connect method.
    With gclsSQLServer
        If gblnFromAppMenu Then
            .AutoLogin mstrUser, mstrPassword, mstrDatabase
        Else
            .Login mstrDatabase
        End If
    
        .Server = mstrServer
        .Database = mstrDatabase
        
        ' Check the state of the connection object.
        If Len(.ConnectString) = 0 Then
            LoginToSQLServer = False
            GoTo PROC_EXIT
        Else
            LoginToSQLServer = True
       End If
    End With
    
    gclsMESApplication.SetCurrentUser
        
    If UCase(gclsMESApplication.ApplicationRole) = "MNB_UPDATE" Or _
            UCase(gclsMESApplication.ApplicationRole) = "MNB_ADMINISTRATION" Or _
            UCase(gclsMESApplication.ApplicationRole) = "MNB_INQUIRY" Then
        gblnUpdate = True
    Else
        MsgBox "User " & gclsSQLServer.UserID & _
            " is not authorized to use this application.  " & _
            "Contact your HelpDesk at " & _
            gclsMESApplication.DivisionHelpDesk, _
                vbExclamation + vbOKOnly
        LoginToSQLServer = False
    End If

PROC_EXIT:
    Exit Function
    
PROC_ERR:
    Call ShowError(Me.Name, "LoginToSQLServer", Err.Number, _
        Err.Description)
    LoginToSQLServer = False
    GoTo PROC_EXIT

End Function


Private Sub ProcessINIFile()
    ' Purpose:  This function reads the INI File and retrieves
    '           the Server name and Database name.
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ' Declare a variable to hold a line from the INI file
    Dim strINILine As String
    
    gstrINIFile = App.Path & "\MiniBill.INI"
    Open gstrINIFile For Input As #1
        
    ' Loop through the each record in the INI file and find
    ' the server and database name.
    Line Input #1, strINILine
    Do While Not EOF(1)
        
        ' If the line just read contains Server= set the
        ' server name variable.
        If Left$(strINILine, 7) = "Server=" Then
            mstrServer = UCase(Trim(Mid(strINILine, 8)))
            
        ' If the line contains the Database name, place it in
        ' the variable.
        ElseIf Left$(strINILine, 9) = "Database=" Then
            mstrDatabase = Trim(Mid(strINILine, 10))
            
        ' If the HelpFile is specified, set the application
        ' HelpFile property.
        ElseIf Left$(strINILine, 9) = "HelpFile=" Then
            App.HelpFile = App.Path & "\" & Mid(strINILine, 10)
            
        End If
        Line Input #1, strINILine

    Loop
    Close #1
    
PROC_EXIT:
    
    Exit Sub
    
PROC_ERR:
    If Err.Number = 53 Then
        MsgBox "INI file not found."
        cdlgFile.DefaultExt = "ini"
        cdlgFile.Filter = "INI Files (*.ini)|*.INI"
        cdlgFile.ShowOpen
        gstrINIFile = cdlgFile.FileName
        Resume
    Else
        Call ShowError(Me.Name, "ProcessINIFile", Err.Number, _
            Err.Description)
        GoTo PROC_EXIT
    End If
End Sub



Private Sub mnuModelLineStockingLocationNotes_Click()
    On Error Resume Next
    frmMNBModelLineStockLocNotes.Show vbModal
End Sub

Private Sub mnuReportsDailyScheduleAll_Click()
    On Error Resume Next
    frmDailyScheduleSheetAllModels.Show vbModal
End Sub

Private Sub mnuTemporaryNewPart_Click()
    On Error Resume Next
    frmTemporaryNewPart.Show vbModal
End Sub

Private Sub menuMNBOverrideECN_Click()
        
    On Error Resume Next
    gblnFindEcnInfo = True
    frmMNBOverrideEcnInfoFind.Show vbModal
    gblnFindEcnInfo = False
    
    If gblnFindEcnInfoCancel = False Then
        frmMNBOverrideECNInfo.Show
    End If
    
    gblnFindEcnInfoCancel = False
End Sub

Private Sub mnuFileChangePassword_Click()
    Dim clsUser As MES.User
    Dim conDatabase As ADODB.Connection
    
    Set conDatabase = gclsSQLServer.Connect
    Set clsUser = New MES.User
    
    clsUser.UserPasswordChange conDatabase
    
    Set clsUser = Nothing
    conDatabase.Close
    Set conDatabase = Nothing
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuHelpAbout_Click()
    Dim clsSplash As MES.Splash
    Set clsSplash = New MES.Splash
    clsSplash.Show
End Sub

Private Sub mnuHelpContents_Click()
    ShowHelpContents
End Sub

Private Sub mnuHelpIndex_Click()
    ShowHelpIndex
End Sub


Private Sub mnuMNBModelLineInactivity_Click()
    On Error Resume Next
    frmMNBModelLineInactive.Show vbModal
End Sub

Private Sub mnuMNBModelLocationSubAssembly_Click()
    On Error Resume Next
    frmMNBModelLocationSubAssembly.Show vbModal
End Sub
Private Sub mnuMNBModelLineStockingLocationNotes_Click()
    On Error Resume Next
    frmMNBModelLineStockLocNotes.Show vbModal
End Sub
Private Sub mnuReportRackPickList_Click()
    On Error Resume Next
    frmRackPickList.Show vbModal
End Sub

Private Sub mnuReportsDailySchedule_Click()
    On Error Resume Next
    frmDailyScheduleSheet.Show vbModal
End Sub

Private Sub mnuSetupECN_Click()
    On Error Resume Next
    frmECNReview.Show vbModal
End Sub

Private Sub mnuTempPartsList_Click()
    On Error Resume Next
    frmTempPartsList.Show vbModal
End Sub

Private Sub mnuSetupMiniBill_Click()
    On Error Resume Next
    frmMiniBillMaintenance.Show vbModal
End Sub

Private Sub mnuSetupModelLineToLineCopy_Click()
    On Error Resume Next
    frmModelLineToLineCopy.Show vbModal
End Sub

Private Sub mnuSetupModelsByPart_Click()
    On Error Resume Next
    frmModelsByPart.Show vbModal
End Sub

Private Sub mnuSetupNewModel_Click()
    On Error Resume Next
    frmNewModelProcessing.Show vbModal
End Sub

Private Sub mnuSetupNewPartsReview_Click()
    On Error Resume Next
    frmNewPartsReview.Show vbModal
End Sub

Private Sub mnuSetupPartCategory_Click()
    On Error Resume Next
    frmCategoryPart.Show vbModal
End Sub

Private Sub mnuSetupPartLocation_Click()
    On Error Resume Next
    frmPartSetup.Show vbModal
End Sub

Private Sub mnuSetupSplitPart_Click()
    On Error Resume Next
    frmSplitPartMaintenance.Show vbModal
End Sub



Private Sub mnuToolsArea_Click()
    On Error Resume Next
    frmArea.Show vbModal
End Sub
Private Sub mnuToolsAreaLocation_Click()
    On Error Resume Next
    frmAreaLocation.Show vbModal
End Sub

Private Sub mnuToolsCategory_Click()
    On Error Resume Next
    frmCategory.Show vbModal
End Sub

Private Sub mnuToolsLocation_Click()
    On Error Resume Next
    frmLocation.Show vbModal
End Sub

Private Sub mnuToolsLineLocation_Click()
    On Error Resume Next
'    frmLineLocation.Show vbModal
End Sub

Private Sub mnuToolsLocationCategory_Click()
    On Error Resume Next
'    frmLocationCategory.Show vbModal
End Sub

Private Sub mnuToolsUserSetup_Click()
    Dim clsUser As MES.User
    Set clsUser = New MES.User
    clsUser.Maintenance gclsMESApplication.ID, True
End Sub
