VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmArea 
   Caption         =   "MiniBill - Area Maintenance"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   StartUpPosition =   1  'CenterOwner
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   4110
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   7710
      _cx             =   13600
      _cy             =   7250
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
      Begin VB.CommandButton cmdCancel 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   840
         Picture         =   "frmArea.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdClose 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   3780
         Picture         =   "frmArea.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Close This Form"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdHelp 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   3420
         Picture         =   "frmArea.frx":08EC
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Help"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdList 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   3060
         Picture         =   "frmArea.frx":09EE
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Refresh Data Without Saving"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdLast 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   2640
         Picture         =   "frmArea.frx":0B78
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Display Last Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdNext 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   2280
         Picture         =   "frmArea.frx":1152
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Display Next Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdPrevious 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   1920
         Picture         =   "frmArea.frx":172C
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Display Previous Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdFirst 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   1560
         Picture         =   "frmArea.frx":1D06
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Display First Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdFind 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   1200
         Picture         =   "frmArea.frx":22E0
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Find An Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdSave 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   480
         Picture         =   "frmArea.frx":2876
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Save Changes"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   120
         Picture         =   "frmArea.frx":2E20
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Add new entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.TextBox txtAreaID 
         Height          =   360
         Left            =   3300
         MaxLength       =   5
         TabIndex        =   0
         Tag             =   "Area:"
         Top             =   900
         Width           =   795
      End
      Begin VB.TextBox txtAreaDescription 
         Height          =   360
         Left            =   3300
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Description:"
         Top             =   1545
         Width           =   2955
      End
      Begin VB.CheckBox chkObsolete 
         Caption         =   "Obsolete?"
         Height          =   360
         Left            =   3300
         TabIndex        =   2
         Top             =   2175
         Width           =   1335
      End
      Begin VB.CommandButton cmdObsoleteDateCalendar 
         CausesValidation=   0   'False
         Height          =   450
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2820
         Width           =   450
      End
      Begin MSComctlLib.StatusBar staDBMaint 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   3795
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               Alignment       =   1
               TextSave        =   "10/5/2007"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               Alignment       =   1
               TextSave        =   "2:35 PM"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Bevel           =   2
               Object.Width           =   8414
               Key             =   "Msg"
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin MSMask.MaskEdBox mskObsoleteDate 
         Height          =   360
         Left            =   3300
         TabIndex        =   3
         Tag             =   "Obsolete Date:"
         Top             =   2820
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   10
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   0
         X1              =   120
         X2              =   7440
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   1
         X1              =   120
         X2              =   7440
         Y1              =   525
         Y2              =   525
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
         Height          =   315
         Left            =   5040
         TabIndex        =   7
         Top             =   120
         Width           =   2475
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
Attribute VB_Name = "frmArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents mrsDatabase As ADODB.Recordset
Attribute mrsDatabase.VB_VarHelpID = -1

Private mblnRecChanged As Boolean
Public mstrAreaID As String

Private Sub chkObsolete_Click()
    ' Purpose:  Change the value in the recordset based on
    '           the value new value.  Set the enabled property
    '           of the obsolete date based on the value of
    '           this control.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare holding field
    Dim blnObsolete As Boolean
    
    ' Set the holding field and check to see if it has
    ' changed.  If it has changed, change the date field
    ' in the file.
    
    blnObsolete = chkObsolete.Value
    If blnObsolete = mrsDatabase!stocking_area_obsolete_flag Then
        Exit Sub
    End If
    
    If mrsDatabase!stocking_area_obsolete_flag <> blnObsolete Then
        mblnRecChanged = True
        mrsDatabase!stocking_area_obsolete_flag = chkObsolete.Value
        mskObsoleteDate.Enabled = blnObsolete
        cmdObsoleteDateCalendar.Enabled = mskObsoleteDate.Enabled
        
        ' If the Obsolete Date field is enabled, set the date
        ' to today.  Otherwise set it to an empty date.
        If blnObsolete Then
            mskObsoleteDate.Text = _
                Format$(Date, "mm/dd/yyyy")
            mrsDatabase!stocking_area_obsolete_date = _
                CDate(mskObsoleteDate.Text)
        Else
            mskObsoleteDate.Text = "__/__/____"
            mrsDatabase!stocking_area_obsolete_date = Null
        End If
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "chkObsolete_Click", _
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

Private Sub cmdObsoleteDateCalendar_Click()
    dlgCalendar.mdteSelectedDate = CDate(mskObsoleteDate.Text)
    dlgCalendar.Show vbModal
    If Not IsNull(dlgCalendar.mdteSelectedDate) Then
        mskObsoleteDate.Text = Format( _
            dlgCalendar.mdteSelectedDate, "mm/dd/yyyy")
        If mrsDatabase!stocking_area_obsolete_date <> _
                dlgCalendar.mdteSelectedDate Then
            mblnRecChanged = True
            mrsDatabase!stocking_area_obsolete_date = _
                dlgCalendar.mdteSelectedDate
        End If
    End If
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
        
        txtAreaID.Enabled = False
        txtAreaDescription.Enabled = False
        chkObsolete.Enabled = False
        mskObsoleteDate.Enabled = False
        cmdObsoleteDateCalendar.Enabled = mskObsoleteDate.Enabled
    End If
    
    ' Retrieve the data
    Call RetrieveData
    If mrsDatabase.RecordCount > 1 Then
        frmAreaDisplay.Show vbModal
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
    ' a specific Area.
    frmAreaDisplay.Show vbModal
    

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
    ' a specific Area.
    frmAreaFind.Show vbModal
    
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
            txtAreaID.Text = !stocking_area_id
            txtAreaDescription.Text = Trim(!stocking_area_description)
            If !stocking_area_obsolete_flag Then
                chkObsolete.Value = 1
                mskObsoleteDate.Text = _
                    Format$(!stocking_area_obsolete_date, "mm/dd/yyyy")
            Else
                mskObsoleteDate.Text = "__/__/____"
                chkObsolete.Value = 0
            End If
            If gblnUpdate Then
                mskObsoleteDate.Enabled = !stocking_area_obsolete_flag
                cmdObsoleteDateCalendar.Enabled = mskObsoleteDate.Enabled
            End If
            If .EditMode = adEditAdd Then
                cmdCancel.ToolTipText = "Cancel Add"
                txtAreaID.Enabled = True
            Else
                cmdCancel.ToolTipText = "Cancel Update of Current Entry"
                txtAreaID.Enabled = False
            End If
        
            ' Set focus to the color ID field
            If Screen.ActiveForm Is Me And gblnUpdate Then
                If .EditMode = adEditAdd Then
                    txtAreaID.SetFocus
                Else
                    txtAreaDescription.SetFocus
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
        .Source = "v_prod_stocking_area_all"
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
        txtAreaID.Text = vbNullString
        txtAreaDescription.Text = vbNullString
        chkObsolete.Value = 0
        mskObsoleteDate.Text = "__/__/____"
        mskObsoleteDate.Enabled = False
        cmdObsoleteDateCalendar.Enabled = False
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

Private Sub txtAreaDescription_GotFocus()
    ' Purpose:  Select the field for easy update
    
    txtAreaDescription.SelStart = 0
    txtAreaDescription.SelLength = Len(txtAreaDescription.Text)
End Sub

Private Sub txtAreaDescription_KeyPress(KeyAscii As Integer)
    mblnRecChanged = True
End Sub

Private Sub txtAreaDescription_Validate(Cancel As Boolean)
    ' Purpose:  Make sure that the description field is
    '           not empty.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    If Len(Trim(txtAreaDescription.Text)) = 0 Then
        Cancel = True
        MsgBox "Area Description Is Required!", _
            vbExclamation + vbOKOnly, _
            "Area Description Validation"
        GoTo PROC_EXIT
Else
        mrsDatabase!stocking_area_description = _
            txtAreaDescription.Text
        Cancel = False
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "txtAreaDescription_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub txtAreaID_GotFocus()
    ' Purpose:  Select the field for easy update
    
    If Len(txtAreaID.Text) > 0 Then
        txtAreaID.SelStart = 0
        txtAreaID.SelLength = Len(txtAreaID.Text)
    End If
End Sub

Private Sub txtAreaID_KeyPress(KeyAscii As Integer)
    ' Purpose:  Change any character to uppercase and ignore spaces
    
    If KeyAscii <> vbKeySpace Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        mblnRecChanged = True
    Else
        KeyAscii = vbCancel
    End If

End Sub

Private Sub txtAreaID_Validate(Cancel As Boolean)
    ' Purpose:  Validate that the color ID field is 3 digits
    '           in length.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
        
    ' Set a variable to hold a copy of the recordset to be used
    ' to check for a duplicate.
    Dim rsDuplicate As ADODB.Recordset
    
    ' If the length of the field is zero, give an error
    If Len(txtAreaID.Text) = 0 Then
        Cancel = True
        MsgBox "Area ID is required", _
            vbExclamation + vbOKOnly, _
            "Area ID Validation"
        GoTo PROC_EXIT
    Else
        ' Check to see if the Area has changed.  If it has,
        ' see if the Area entered already exists.  If it
        ' exists, give an error.  If not, change the Area ID
        ' in the recordset.
        If Trim(txtAreaID.Text) <> Trim(mrsDatabase!stocking_area_id) Then
            Set rsDuplicate = mrsDatabase.Clone
            rsDuplicate.Find ( _
                "stocking_area_ID = '" & txtAreaID.Text & "'")
            If rsDuplicate.EOF Then
                mrsDatabase!stocking_area_id = txtAreaID.Text
                Cancel = False
            Else
                MsgBox "Area ID " & txtAreaID.Text & _
                    " already exists."
                Cancel = True
            End If
            Set rsDuplicate = Nothing
        End If
    End If
       
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "txtAreaID_Validate", _
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
    
    
    ' Validate the color ID field.
    Call txtAreaID_Validate(blnCancel)
    
    ' If the color ID was not valid or was missing, set focus
    ' back to the color id and exit the function.
    If blnCancel Then
        txtAreaID.SetFocus
        GoTo PROC_EXIT
    End If
        
    ' Validate the Unit description
    Call txtAreaDescription_Validate(blnCancel)
    
    ' If the Area description was not valid, set focus to the
    ' field and leave the sub.
    If blnCancel Then
        txtAreaDescription.SetFocus
        GoTo PROC_EXIT
    End If
    
    ' If the obsolete date is enabled, validate it.
    If mskObsoleteDate.Enabled Then
        Call mskObsoleteDate_Validate(blnCancel)
        
        ' If the obsolete date was not valid, set focus to it
        ' and exit the sub.
        If blnCancel Then
            mskObsoleteDate.SetFocus
            GoTo PROC_EXIT
        End If
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



Private Sub mskobsoleteDate_GotFocus()
    ' Purpose:  Select the field for easy update
    
    mskObsoleteDate.SelStart = 0
    mskObsoleteDate.SelLength = 10
End Sub

Private Sub mskObsoleteDate_Validate(Cancel As Boolean)
    ' Purpose:  Validate the field to make sure that it is
    '           either a valid date or empty.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    If IsDate(mskObsoleteDate.Text) Then
        mrsDatabase!stocking_area_obsolete_date = _
            CDate(mskObsoleteDate.Text)
    ElseIf chkObsolete.Value = 0 And Len(mskObsoleteDate.ClipText) = 0 Then
        mrsDatabase!stocking_area_obsolete_date = Null
    Else
        Cancel = True
        MsgBox "Invalid Obsolete Date Entered!", _
            vbExclamation + vbOKOnly, _
            "Obsolete Date Validateion Error"
        mskObsoleteDate.SelStart = 0
        mskObsoleteDate.SelLength = 10
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mskObsoleteDate_Validate", _
        Err.Number, Err.Description)
End Sub


