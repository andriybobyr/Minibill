VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMNBModelLineInactivity 
   Caption         =   "MiniBill - Model/Line Inactivity Flag Maintenance"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   581
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   1  'CenterOwner
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   8715
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   15240
      _cx             =   26882
      _cy             =   15372
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483630
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   6
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   2500
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   -1  'True
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin VB.CheckBox chkMNBModelLineActiveFlag 
         Caption         =   "Model/Line Active Flag"
         Height          =   945
         Left            =   5760
         TabIndex        =   17
         Top             =   6840
         Width           =   2640
      End
      Begin VB.TextBox txtMNBLine 
         Height          =   960
         Left            =   5760
         MaxLength       =   30
         TabIndex        =   16
         Tag             =   "Line:"
         Top             =   5175
         Width           =   5835
      End
      Begin VB.TextBox txtMNBModel 
         Height          =   945
         Left            =   5760
         MaxLength       =   5
         TabIndex        =   15
         Tag             =   "Model:"
         Top             =   3480
         Width           =   4815
      End
      Begin VB.CommandButton cmdCancel 
         CausesValidation=   0   'False
         Height          =   1620
         Left            =   3285
         Picture         =   "frmMNBModelLineInactivity.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   660
         Width           =   1185
      End
      Begin VB.CommandButton Command10 
         CausesValidation=   0   'False
         Height          =   1620
         Left            =   14775
         Picture         =   "frmMNBModelLineInactivity.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Close This Form"
         Top             =   660
         Width           =   1170
      End
      Begin VB.CommandButton Command9 
         CausesValidation=   0   'False
         Height          =   1620
         Left            =   13365
         Picture         =   "frmMNBModelLineInactivity.frx":08EC
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Help"
         Top             =   660
         Width           =   1185
      End
      Begin VB.CommandButton Command8 
         CausesValidation=   0   'False
         Height          =   1620
         Left            =   11955
         Picture         =   "frmMNBModelLineInactivity.frx":09EE
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Refresh Data Without Saving"
         Top             =   660
         Width           =   1170
      End
      Begin VB.CommandButton Command7 
         CausesValidation=   0   'False
         Height          =   1620
         Left            =   10320
         Picture         =   "frmMNBModelLineInactivity.frx":0B78
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Display Last Entry"
         Top             =   660
         Width           =   1185
      End
      Begin VB.CommandButton Command6 
         CausesValidation=   0   'False
         Height          =   1620
         Left            =   8895
         Picture         =   "frmMNBModelLineInactivity.frx":1152
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Display Next Entry"
         Top             =   660
         Width           =   1185
      End
      Begin VB.CommandButton Command5 
         CausesValidation=   0   'False
         Height          =   1620
         Left            =   7500
         Picture         =   "frmMNBModelLineInactivity.frx":172C
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Display Previous Entry"
         Top             =   660
         Width           =   1185
      End
      Begin VB.CommandButton Command4 
         CausesValidation=   0   'False
         Height          =   1620
         Left            =   6120
         Picture         =   "frmMNBModelLineInactivity.frx":1D06
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Display First Entry"
         Top             =   660
         Width           =   1185
      End
      Begin VB.CommandButton Command3 
         CausesValidation=   0   'False
         Height          =   1620
         Left            =   4710
         Picture         =   "frmMNBModelLineInactivity.frx":22E0
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Find An Entry"
         Top             =   660
         Width           =   1155
      End
      Begin VB.CommandButton cmdSave 
         CausesValidation=   0   'False
         Height          =   1620
         Left            =   1875
         Picture         =   "frmMNBModelLineInactivity.frx":2876
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Save Changes"
         Top             =   660
         Width           =   1170
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   1620
         Left            =   465
         Picture         =   "frmMNBModelLineInactivity.frx":2E20
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Add new entry"
         Top             =   660
         Width           =   1185
      End
      Begin VB.CheckBox chkModelLineInactivityFlag 
         Caption         =   "Model/Line Inactivity Flag"
         Height          =   2025
         Left            =   12825
         TabIndex        =   0
         Top             =   12075
         Width           =   10350
      End
      Begin MSComctlLib.StatusBar staDBMaint 
         Height          =   2115
         Left            =   0
         TabIndex        =   2
         Top             =   21210
         Width           =   30120
         _ExtentX        =   53128
         _ExtentY        =   3731
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               Alignment       =   1
               TextSave        =   "1/26/2004"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               Alignment       =   1
               TextSave        =   "2:40 PM"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Bevel           =   2
               Object.Width           =   47943
               Key             =   "Msg"
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   0
         X1              =   465
         X2              =   29055
         Y1              =   2670
         Y2              =   2670
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   1
         X1              =   465
         X2              =   29055
         Y1              =   2925
         Y2              =   2925
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Whirlpool MES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1755
         Left            =   19695
         TabIndex        =   3
         Top             =   660
         Width           =   9660
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileCancel 
         Caption         =   "C&ancel Current Update/Add"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewFind 
         Caption         =   "F&ind"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFirst 
         Caption         =   "&First Entry"
      End
      Begin VB.Menu mnuViewPrevious 
         Caption         =   "&Previous Entry"
      End
      Begin VB.Menu mnuViewNext 
         Caption         =   "&Next Entry"
      End
      Begin VB.Menu mnuViewLast 
         Caption         =   "&Last"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "&Selection List..."
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
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMNBModelLineInactivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents mrsDatabase As ADODB.Recordset
Attribute mrsDatabase.VB_VarHelpID = -1

Private mblnRecChanged As Boolean
Public mstrMNBModelID As String

Private Sub chkModelLineInactivityFlag_Click()
    ' Purpose:  Change the value in the recordset based on
    '           the value new value.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare holding field
    Dim blnObsolete As Boolean
    
    ' Set the holding field and check to see if it has
    ' changed.  If it has changed, change the date field
    ' in the file.
    
    blnObsolete = chkModelLineInactivityFlag.Value
    If blnObsolete = mrsDatabase!Model_Line_inactivity_flag Then
        Exit Sub
    End If
    
    If mrsDatabase!Model_Line_inactivity_flag <> blnObsolete Then
        mblnRecChanged = True
        mrsDatabase!Model_Line_inactivity_flag = chkModelLineInactivityFlag.Value
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "chkModelLineInactivityFlag", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cmdAdd_Click()
    ' Purpose:  Tie the click of this button to the FileNew
    '           menu item.
    
    Call mnuFileNew_Click
End Sub

Private Sub cmdClose_Click()
    ' Purpose:  Close the form by the user's request.
    
    Call mnuFileClose_Click
End Sub

Private Sub cmdFind_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the ViewFind menu item.
    
    Call mnuViewFind_Click
End Sub

Private Sub cmdFirst_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the ViewFirst menu item.
    
    Call mnuViewFirst_Click
End Sub

Private Sub cmdHelp_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the HelpContents menu item.
    
    Call mnuHelpContents_Click
End Sub

Private Sub cmdLast_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the ViewLast menu item.
    
    Call mnuViewLast_Click
End Sub

Private Sub cmdList_Click()
    Call mnuViewList_Click
    
End Sub

Private Sub cmdNext_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the ViewNext menu item.
    
    Call mnuViewNext_Click
End Sub

Private Sub cmdPrevious_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the ViewPrevious menu item.
    
    Call mnuViewPrevious_Click
End Sub


Private Sub cmdSave_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the FileSave menu item.
    
    Call mnuFileSave_Click
End Sub



Private Sub Form_Load()
    ' Purpose:  Show the form and login to the server
    
    ' Set up error handling
    On Error GoTo PROC_ERR
       
    
    ' Connect to the database
    If gconDatabase Is Nothing Then
        Set gconDatabase = gclsSQLServer.Connect( _
            gclsMESApplication.ApplicationRole, _
            gclsMESApplication.ApplicationPassword)
    
        If gconDatabase.State <> adStateOpen Then
            Err.Raise gconDatabase.Errors(0).NativeError, _
                "Form_Load", gconDatabase.Errors(0).Description
        End If
    End If
    
    
    ' Hide buttonws and menu choices for non-update
    If Not gblnUpdate Then
        cmdAdd.Enabled = False
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        
        mnuFileNew.Enabled = False
        mnuFileSave.Enabled = False
        mnuFileCancel.Enabled = False
        
        txtMNBModel.Enabled = False
        txtMNBLine.Enabled = False
        chkModelLineInactivityFlag.Enabled = False
    End If
    
    ' Retrieve the data
    Call RetrieveData
    If mrsDatabase.RecordCount > 1 Then
        frmMNBModelLineInactivityDisplay.Show vbModal
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Load", Err.Number, _
        Err.Description)
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Purpose:  Ask the user if he would like to save changes
    '           before closing the form.  If "Yes", call the
    '           mnuFileSave_Click procedure.  If "Cancel",
    '           set the Cancel flag to true and exit.  If no
    '           exit without setting the Cancel flag to
    '           true.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare Return Code variable
    Dim intRetCode As Integer
    
    ' Check to see if any changes have been made to the
    ' recordset.
    If mblnRecChanged Then
    
        ' Ask the user if he would like to save the changes.
        intRetCode = MsgBox("Save Changes?", _
            vbQuestion + vbYesNoCancel, "Closing")
        If intRetCode = vbYes Then
        ' Validate controls
            If Not ValidEntries Then
                Cancel = True
                GoTo PROC_EXIT
            End If
            
            Call mnuFileSave_Click
            Cancel = False
        ElseIf intRetCode = vbCancel Then
            Cancel = True
        Else
            Cancel = False
        End If
        
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_QueryUnload", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Purpose:  Close and dereference the connection and
    '           recoredset.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Check to see that the connection has been instantiated.
    If Not gconDatabase Is Nothing Then
    
        ' If the connection is open, close it.
        If gconDatabase.State = adStateOpen Then
        
            ' Check to see if the recordset is instantiated.
            If Not mrsDatabase Is Nothing Then
            
                ' If the recordset is open, close it.
                If mrsDatabase.State = adStateOpen Then
                    mrsDatabase.Close
                End If
                Set mrsDatabase = Nothing
            End If
        End If
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Unload", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuFileCancel_Click()
    ' Purpose:  This procedure will delete a newly added
    '           record or return an updated record to its
    '           original state.
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ' Set up a field object
    Dim fld As ADODB.Field
    
    ' If the record is newly added, delete it.
    With mrsDatabase
        If .EditMode = adEditAdd Then
            .Delete
            .MoveFirst
            GoTo PROC_EXIT
        End If
        
        ' Loop through the existing record and reset
        ' each field to it original value.
        For Each fld In .Fields
            fld.Value = fld.OriginalValue
        Next fld
        Set fld = Nothing
        mrsDatabase_MoveComplete adRsnMove, Nothing, adStatusOK, _
            Nothing
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuFileCancel_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cmdCancel_Click()
    mnuFileCancel_Click
End Sub

Private Sub mnuFileClose_Click()
    '  Purpose: Close this form
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Unload Me
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuClose_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub mnuFileExit_Click()
    ' Purpose:  Close the application
    
    Unload frmMain
End Sub

Private Sub mnuFileNew_Click()
    ' Purpose:  Create a new record and set fields on the
    '           form to their inital values.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Validate the controls
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If
    
    ' Call a Recordset Add
    mrsDatabase.AddNew
    
    ' Check for an error being returned from the operation.
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "cmdAdd_Click", _
            gconDatabase.Errors(0).Description
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdAdd_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub mnuFileSave_Click()
    ' Purpose:  Save the current changes to the database
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Validate the controls
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If
    
    ' Attempt to update the data
    mrsDatabase.UpdateBatch
    
    ' Check for errors
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "mnuSave_Click", _
            gconDatabase.Errors(0).Description
    End If
    
    ' Reset the Record Changed flag to false.
    mblnRecChanged = False
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuFileSave_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuHelpAbout_Click()
    ' Purpose:  Display the About form for the application
    
    ' Declare and instantiate Splash screen object.
    Dim clsSplash As MES.Splash
    Set clsSplash = New MES.Splash
    
    clsSplash.Show
    
End Sub

Private Sub mnuHelpContents_Click()
    ' Purpose:  Display Help Contents
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ShowHelpContents
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuHelpContents_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuHelpIndex_Click()
    ' Purpose:  Display Help Index
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ShowHelpIndex
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuHelpContents_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuViewList_Click()
    ' Purpose:  Find a specific record in the database and move
    '            to that record.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
        
    ' Validate controls before moving
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If

    ' Call the procedure to load the form to be used to find
    ' a specific Model/Line.
    frmMNBModelLineInactivityDisplay.Show vbModal
    

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewFind_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuViewFind_Click()
    ' Purpose:  Find a specific record in the database and move
    '            to that record.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
        
    ' Validate controls before moving
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If

    ' Call the procedure to load the form to be used to find
    ' a specific Model/Line.
    frmMNBModelLineInactivityFind.Show vbModal
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewFind_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuViewFirst_Click()
    ' Purpose:  Move to the first record in the recordset
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Validate controls before moving
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If
    
    ' Move to the first record in the recordset
    mrsDatabase.MoveFirst
    
    ' Check return code from Recordset operation.  Raise an
    ' error if the operation failed.  If the errors collection
    ' of the connection object contains errors, this means the
    ' operation failed.
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "mnuViewFirst_Click", _
            gconDatabase.Errors(0).Description
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewFirst_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub mnuViewPrevious_Click()
    ' Purpose:  Move to the previous record in the recordset
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Validate controls before moving
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If
    
    ' Move to the previous record in the recordset
    With mrsDatabase
        .MovePrevious
        
        ' If the cursor is at the beginning of file, move to
        ' the last record.
        If .BOF Then
            .MoveLast
        End If
    End With
    
    ' Check return code from Recordset operation.  Raise an
    ' error if the operation failed.  If the errors collection
    ' of the connection object contains errors, this means the
    ' operation failed.
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "mnuViewPrevious_Click", _
            gconDatabase.Errors(0).Description
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewPrevious_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuViewNext_Click()
    ' Purpose:  Move to the next record in the recordset
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Validate controls before moving
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If
    
    ' Move to the next record in the recordset
    With mrsDatabase
        .MoveNext
        
        ' If the cursor is at the end of file, move to
        ' the first record.
        If .EOF Then
            .MoveFirst
        End If
    End With
    
    ' Check return code from Recordset operation.  Raise an
    ' error if the operation failed.  If the errors collection
    ' of the connection object contains errors, this means the
    ' operation failed.
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "mnuViewNext_Click", _
            gconDatabase.Errors(0).Description
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewNext_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuViewLast_Click()
    ' Purpose:  Move to the last record in the recordset
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Validate controls before moving
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If
    
    ' Move to the last record in the recordset
    mrsDatabase.MoveLast
    
    ' Check return code from Recordset operation.  Raise an
    ' error if the operation failed.  If the errors collection
    ' of the connection object contains errors, this means the
    ' operation failed.
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "mnuViewLast_Click", _
            gconDatabase.Errors(0).Description
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewLast_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mrsDatabase_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    ' Purpose:  Fill the fields after a move in the
    '           recordset is complete.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' If the recordset is at a valid record, fill the
    ' controls.
    With mrsDatabase
        If Not .EOF And Not .BOF Then
            ' Add code here to fill the controls on the form
            ' with the data from the Recordset.
            txtMNBModel.Text = !model_number
            txtMNBLine.Text = !line_id
            If !Model_Line_inactive_flag Then
                chkMNBModelLineActiveFlag.Value = 1
            Else
                chkMNBModelLineActiveFlag.Value = 0
            End If
            
            If .EditMode = adEditAdd Then
                cmdCancel.ToolTipText = "Cancel Add"
                txtMNBModel.Enabled = True
            Else
                cmdCancel.ToolTipText = "Cancel Update of Current Entry"
                txtMNBModel.Enabled = False
            End If
        
            ' Set focus to the Model ID field
            If Screen.ActiveForm Is Me And gblnUpdate Then
                If .EditMode = adEditAdd Then
                    txtMNBModel.SetFocus
                Else
                    txtMNBLine.SetFocus
                End If
            End If
        End If
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mrsDatabase_MoveComplete", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
    
End Sub

Sub RetrieveData()
    ' Purpose:  Instantiate and open the recordset.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Instantiate the recordset
    Set mrsDatabase = New ADODB.Recordset
    
    ' Set values of fields
    With mrsDatabase
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        ' Change the literal below to the name of your view
        .Source = "v_mnb_model_line_all"
        .LockType = adLockBatchOptimistic
        .Open
    End With
    
    ' Check for errors returned from the recordset
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise vbObjectError + 1000, "RetrieveData", _
        gconDatabase.Errors(0).Description
    End If
    
    ' if no records were retrieved, add a new record to the
    ' recordset and reset fields to their original value.
    If mrsDatabase.EOF Then
        mrsDatabase.AddNew
        txtMNBModel.Text = vbNullString
        txtMNBLine.Text = vbNullString
        chkModelLineInactivityFlag.Value = 0
        chkModelLineInactivityFlag.Value = 0
    Else
        mblnRecChanged = False
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveData", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mskobsoleteDate_KeyPress(KeyAscii As Integer)
    mblnRecChanged = True
End Sub

Private Sub txtMNBLine_GotFocus()
    ' Purpose:  Select the field for easy update
    
    txtMNBLine.SelStart = 0
    txtMNBLine.SelLength = Len(txtMNBLine.Text)
End Sub

Private Sub txtMNBLine_KeyPress(KeyAscii As Integer)
    mblnRecChanged = True
End Sub

Private Sub txtMNBLine_Validate(Cancel As Boolean)
    ' Purpose:  Make sure that the line field is
    '           not empty.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    If Len(Trim(txtMNBLine.Text)) = 0 Then
        Cancel = True
        MsgBox "Line Is Required!", _
            vbExclamation + vbOKOnly, _
            "Line Validation"
        GoTo PROC_EXIT
Else
        mrsDatabase!Line = _
            txtMNBLine.Text
        Cancel = False
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "txtLine_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub txtMNBModel_GotFocus()
    ' Purpose:  Select the field for easy update
    
    If Len(txtMNBModel.Text) > 0 Then
        txtMNBModel.SelStart = 0
        txtMNBModel.SelLength = Len(txtMNBModel.Text)
    End If
End Sub

Private Sub txtMNBModel_KeyPress(KeyAscii As Integer)
    ' Purpose:  Change any character to uppercase
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    mblnRecChanged = True
End Sub

Private Sub txtMNBModel_Validate(Cancel As Boolean)
    ' Purpose:  Validate that the Model is valid
    '           in length.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
        
    ' Set a variable to hold a copy of the recordset to be used
    ' to check for a duplicate.
    Dim rsDuplicate As ADODB.Recordset
    
    ' If the length of the field is zero, give an error
    If Len(txtMNBModel.Text) = 0 Then
        Cancel = True
        MsgBox "Model is required", _
            vbExclamation + vbOKOnly, _
            "Model Validation"
        GoTo PROC_EXIT
    Else
        ' Check to see if the Model has changed.  If it has,
        ' see if the Model entered already exists.  If it
        ' exists, give an error.  If not, change the Area ID
        ' in the recordset.
        If Trim(txtMNBModel.Text) <> Trim(mrsDatabase!model_number) Then
            Set rsDuplicate = mrsDatabase.Clone
            rsDuplicate.Find ( _
                "Model_Number = '" & txtMNBModel.Text & "'")
            If rsDuplicate.EOF Then
                mrsDatabase!model_number = txtMNBModel.Text
                Cancel = False
            Else
                MsgBox "Model " & txtMNBModel.Text & _
                    " already exists."
                Cancel = True
            End If
            Set rsDuplicate = Nothing
        End If
    End If
       
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "txtMNBModel_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub
Function ValidEntries() As Boolean
    ' Purpose:  Validate fields before changing reocrds or
    '           updating.
    ' Returns:  Boolean determining whether the updates were
    '           successful.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Set up variable to hold cancel parameter
    Dim blnCancel As Boolean
    
    
    ' Validate the Model Number field.
    Call txtMNBModel_Validate(blnCancel)
    
    ' If the Model Number was not valid or was missing, set focus
    ' back to the Model Number and exit the function.
    If blnCancel Then
        txtMNBModel.SetFocus
        GoTo PROC_EXIT
    End If
        
    ' Validate the Line
    Call txtMNBLine_Validate(blnCancel)
    
    ' If the Line was not valid, set focus to the
    ' field and leave the sub.
    If blnCancel Then
        txtMNBLine.SetFocus
        GoTo PROC_EXIT
    End If
    
 
    
    ' Set up variable to hold field in field collection.
    Dim fld As ADODB.Field
    
    ' Check through the fields collection to see if any field
    ' has changed.  If it has, call the sub to set the update
    ' fields in the record and leave the loop.
'    For Each fld In mrsDatabase.Fields
'        If fld.OriginalValue <> fld.Value Then
'            mblnRecChanged = True
'            mrsDatabase!color_code_last_updated = Now()
'            mrsDatabase!color_code_updated_by = gclsSQLServer.UserID
'            Exit For
'        End If
'    Next fld
    
PROC_EXIT:

    ' Set the return value of the function to the opposite of
    ' the cancel boolean.  This is done because the validate
    ' procedures called set the cancel boolean to true if the
    ' validation failed and false if it succeeded.
    ValidEntries = Not blnCancel
    Exit Function
    
PROC_ERR:
    Call ShowError(Me.Name, "ValidEntries", Err.Number, _
        Err.Description)
    blnCancel = True
    GoTo PROC_EXIT
End Function

