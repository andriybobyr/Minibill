VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMNBModelLineStockLocNotes 
   Caption         =   "MiniBill - Notes by Model, Line, Stocking Location"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7950
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   530
   StartUpPosition =   1  'CenterOwner
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   5205
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7950
      _cx             =   14023
      _cy             =   9181
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Begin VB.ComboBox cboModel 
         Height          =   360
         Left            =   2280
         TabIndex        =   14
         Tag             =   "Model:  "
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtNotes 
         Height          =   975
         Left            =   2280
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   18
         Tag             =   "Notes:  "
         Top             =   3360
         Width           =   4695
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   360
         Left            =   465
         Picture         =   "frmMNBModelLineStockLocNotes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Delete Current Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.ComboBox cboLocation 
         Enabled         =   0   'False
         Height          =   360
         Left            =   2280
         TabIndex        =   17
         Tag             =   "Location:"
         Top             =   2600
         Width           =   3060
      End
      Begin VB.ComboBox cboLine 
         Enabled         =   0   'False
         Height          =   360
         Left            =   2280
         TabIndex        =   15
         Tag             =   "Line:"
         Top             =   1840
         Width           =   990
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   120
         Picture         =   "frmMNBModelLineStockLocNotes.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Add New Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdSave 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   825
         Picture         =   "frmMNBModelLineStockLocNotes.frx":0724
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Save Changes"
         Top             =   120
         Width           =   315
      End
      Begin VB.CommandButton cmdFind 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   1575
         Picture         =   "frmMNBModelLineStockLocNotes.frx":0CCE
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Find An Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdFirst 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   2100
         Picture         =   "frmMNBModelLineStockLocNotes.frx":1264
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Display First Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdPrevious 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   2460
         Picture         =   "frmMNBModelLineStockLocNotes.frx":183E
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Display Previous Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdNext 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   2820
         Picture         =   "frmMNBModelLineStockLocNotes.frx":1E18
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Display Next Entry"
         Top             =   120
         Width           =   315
      End
      Begin VB.CommandButton cmdLast 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   3195
         Picture         =   "frmMNBModelLineStockLocNotes.frx":23F2
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Display Last Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdList 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   3630
         Picture         =   "frmMNBModelLineStockLocNotes.frx":29CC
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Selection List"
         Top             =   120
         Width           =   480
      End
      Begin VB.CommandButton cmdHelp 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   4155
         Picture         =   "frmMNBModelLineStockLocNotes.frx":2B56
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Help"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdClose 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   4515
         Picture         =   "frmMNBModelLineStockLocNotes.frx":2C58
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Close This Form"
         Top             =   120
         Width           =   420
      End
      Begin VB.CommandButton cmdCancel 
         CausesValidation=   0   'False
         Height          =   390
         Left            =   1200
         Picture         =   "frmMNBModelLineStockLocNotes.frx":33BA
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Cancel Changes"
         Top             =   120
         Width           =   315
      End
      Begin MSComctlLib.StatusBar staDBMaint 
         Height          =   435
         Left            =   15
         TabIndex        =   1
         Top             =   4755
         Width           =   7950
         _ExtentX        =   14023
         _ExtentY        =   767
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
               TextSave        =   "4:15 PM"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Bevel           =   2
               Object.Width           =   8837
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
         X1              =   120
         X2              =   7680
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   1
         X1              =   120
         X2              =   7680
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Whirlpool MES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   5265
         TabIndex        =   2
         Top             =   120
         Width           =   2550
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete"
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
Attribute VB_Name = "frmMNBModelLineStockLocNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents mrsDatabase As ADODB.Recordset
Attribute mrsDatabase.VB_VarHelpID = -1
Private rsDuplicate As ADODB.Recordset

Private mblnRecChanged As Boolean
Private marrstrLocation() As String
Private intRecordCounter As Integer
Private strLocationId As String

Private strDisplay As String

Private Sub cboLine_Change()
    cboFindFirst cboLine
    mblnRecChanged = True
End Sub

Private Sub cboLine_Click()
    Call RetrieveLocationData(cboModel.Text, cboLine.Text)
End Sub

Private Sub cboLine_KeyPress(KeyAscii As Integer)
    cboKeyPress cboLine, KeyAscii
    mblnRecChanged = True
End Sub

Private Sub cboLine_Validate(Cancel As Boolean)
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    ' If there is no data in the field, display a message,
    ' set cancel to true and exit.
    If Len(Trim(cboLine.Text)) = 0 Then
        Cancel = True
        MsgBox "Line Is Required!", _
            vbExclamation + vbOKOnly, _
            "Line Validation"
        GoTo PROC_EXIT
     End If
        
    ' If the line has changed, look up the new
    ' line in the listbox.
    If mrsDatabase!line_id <> cboLine.Text Then
        cboFindFirst cboLine
            
        ' If the line was not found in the list,
        ' display a message, set cancel to true and
        ' exit.
        If cboLine.ListIndex = -1 Then
            MsgBox "Line " & cboLine.Text & " is not valid"
            Cancel = True
            GoTo PROC_EXIT
        End If
            
        Cancel = False
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboLine_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cboLocation_Change()
    cboFindFirst cboLocation
    mblnRecChanged = True
End Sub

Private Sub cboLocation_KeyPress(KeyAscii As Integer)
    cboKeyPress cboLocation, KeyAscii
    mblnRecChanged = True
End Sub

Private Sub cboLocation_Validate(Cancel As Boolean)
    
    On Error GoTo PROC_ERR
    
    ' If there is no data in the field, display a message,
    ' set cancel to true and exit.
    If Len(Trim(cboLocation.Text)) = 0 Then
        Cancel = True
        MsgBox "Location Is Required!", _
            vbExclamation + vbOKOnly, _
            "Location Validation"
        GoTo PROC_EXIT
     End If
        
    cboFindFirst cboLocation
    strLocationId = marrstrLocation(cboLocation.ItemData(cboLocation.ListIndex))

    ' If the Stocking Location id has changed, look up the new
    ' stocking location id in the listbox.
    If Trim(mrsDatabase!stocking_location_id) <> strLocationId Then
            
        Cancel = False
        ' If the stocking location id was not found in the list,
        ' display a message, set cancel to true and
        ' exit.
        If cboLocation.ListIndex = -1 Then
            MsgBox "Location " & cboLocation.Text & " is not valid"
            Cancel = True
            GoTo PROC_EXIT
        End If
            
        Set rsDuplicate = mrsDatabase.Clone
    
        With rsDuplicate
            .MoveFirst
            .Find ("model_number = '" & cboModel.Text & "'")

            Do While Not .EOF
                If Trim(!model_number) <> cboModel.Text Then
                    Exit Do
                End If

                If Trim(!line_id) = cboLine.Text And _
                   Trim(!stocking_location_id) = strLocationId Then
                    Cancel = True
                    MsgBox "Model/Line/Location already exists "
                    Exit Do
                End If
                .Find "model_number = '" & cboModel.Text & "'", 1, adSearchForward, .Bookmark
            Loop
            .Close
        End With

        Set rsDuplicate = Nothing
        
        If Cancel Then
            GoTo PROC_EXIT
        End If
            
   End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboLocation_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub
Private Sub cboModel_Change()
    cboFindFirst cboModel
    mblnRecChanged = True
End Sub

Private Sub cboModel_Click()
    Call RetrieveLineData(cboModel.Text)
End Sub

Private Sub cboModel_KeyPress(KeyAscii As Integer)
    cboKeyPress cboModel, KeyAscii
    mblnRecChanged = True
End Sub

Private Sub cboModel_Validate(Cancel As Boolean)
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    ' If there is no data in the field, display a message,
    ' set cancel to true and exit.
    If Len(Trim(cboModel.Text)) = 0 Then
        Cancel = True
        MsgBox "Model Is Required!", _
            vbExclamation + vbOKOnly, _
            "Model Validation"
        GoTo PROC_EXIT
     End If
        
    ' If the model number has changed, look up the new
    ' model number in the listbox.
    If mrsDatabase!model_number <> cboModel.Text Then
        cboFindFirst cboModel
            
        ' If the model number was not found in the list,
        ' display a message, set cancel to true and
        ' exit.
        If cboModel.ListIndex = -1 Then
            MsgBox "Model " & cboModel.Text & " is not valid"
            Cancel = True
            GoTo PROC_EXIT
        End If
            
        Cancel = False
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboModel_Validate", _
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

Private Sub cmdDelete_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the ViewFind menu item.
    
    Call mnuFileDelete_Click
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
' Purpose:  Tie the click of this button to the selection
'           of the ViewList menu item.
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
    
    ' Hide buttons and menu choices for non-update
    If Not gblnUpdate Then
        cmdAdd.Enabled = False
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        
        mnuFileNew.Enabled = False
        mnuFileSave.Enabled = False
        mnuFileCancel.Enabled = False
        
        cboModel.Enabled = False
        cboLocation.Enabled = False
        cboLine.Enabled = False
        txtNotes.Enabled = False
    End If
    
    ' Retrieve the data existing in the V_MNB_Model_Line_StockLoc_Notes record
    ' and find the first one to display on the selection screen.
    Call RetrieveData
    
    'Create the dropdown for the Model List
    Call RetrieveModelData

    
    If intRecordCounter > 0 Then
        frmMNBModelLineStockLocNotesDisplay.Show vbModal
        mblnRecChanged = False
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
    '           recordset.
    
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


Private Sub mnuFileDelete_Click()
'Purpose:  Delete a record.  The record will actually be deleted once the
'          user exits the form.
    Dim intReturnValue As Integer
    
    With mrsDatabase
        .Delete
        intReturnValue = MsgBox("Are you sure????", vbYesNo, "Delete Minibill Information")
        If intReturnValue = vbYes Then
            mrsDatabase.UpdateBatch
        Else
            mrsDatabase.CancelUpdate
        End If
        If .RecordCount = 0 Then
            RetrieveData
'            mrsDatabase.AddNew
        Else
            .MoveNext
            If .EOF Then
                .MoveFirst
            End If
        End If
    End With
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
      
    ' Clear the dropdowns, and enable the combo boxes to be changed.
    cboModel.Enabled = True
    cboModel.Text = ""
    cboLine.Enabled = True
    cboLine.Text = ""
    cboLocation.Enabled = True
    cboLocation.Text = ""
    cboModel.SetFocus
    txtNotes = ""
    
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
    
    cboModel.Enabled = False
    cboLine.Enabled = False
    cboLocation.Enabled = False
    txtNotes.SetFocus
    
    MsgBox ("Save is Successful!"), vbOKOnly
    
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
    frmMNBModelLineStockLocNotesDisplay.Show vbModal
    

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
    frmMNBModelLineStockLocNotesFind.Show vbModal
    
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
       
    Dim intI As Integer
    ' If the recordset is at a valid record, fill the
    ' controls.
    With mrsDatabase
        If Not .EOF And Not .BOF Then
            ' Add code here to fill the controls on the form
            ' with the data from the Recordset.
            cboModel.Text = !model_number
            txtNotes.Text = !notes
            
            If !model_number <> "" And !line_id <> "" Then
                Call RetrieveLineData(!model_number)
                cboLine.Text = !line_id
            
            'Stocking location combo contains the description, not Stocking Loc Id.
            'Comparing the Stocking Location ID to Stocking location ID array to
            'determine what sequence it is in the list.
               For intI = 0 To UBound(marrstrLocation)
                    If marrstrLocation(intI) = Trim(!stocking_location_id) Then
                        cboLocation.ListIndex = intI
                        Exit For
                    End If
                Next
    
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
        .Source = "select model_number, line_id, stocking_location_id, notes, last_updated_date " & _
            "from V_MNB_Model_Line_StockLoc_Notes order by model_number"
        .LockType = adLockBatchOptimistic
        .Open
         intRecordCounter = .RecordCount

    End With
      
       
    ' Check for errors returned from the recordset
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise vbObjectError + 1000, "RetrieveData", _
        gconDatabase.Errors(0).Description
    End If
    
    ' if no records were retrieved, add a new record to the
    ' recordset and reset fields to their original value.
        If mrsDatabase.EOF Then
            txtNotes.Text = vbNullString
            cboModel.Enabled = True
            cboLine.Enabled = True
            cboLocation.Enabled = True
            cboModel.Text = " "
            cboLine.Text = " "
            cboLocation.Text = " "
            mrsDatabase.AddNew
        Else
            mblnRecChanged = False
            cboModel.Enabled = False
            cboLine.Enabled = False
            cboLocation.Enabled = False
        End If
       
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveData", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Function ValidEntries() As Boolean
    ' Purpose:  Validate fields before changing reocrds or
    '           updating.
    ' Returns:  Boolean determining whether the updates were
    '           successful.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
'    Dim rsDuplicate As ADODB.Recordset
    
    ' Set up variable to hold cancel parameter
    Dim blnCancel As Boolean
    
    ' Validate the Model Number field.
    Call cboModel_Validate(blnCancel)
    
    ' If the Model Number was not valid or was missing, set focus
    ' back to the Model Number and exit the function.
    If blnCancel Then
        cboModel.SetFocus
        GoTo PROC_EXIT
    End If
       
    ' Validate the line id field.
    Call cboLine_Validate(blnCancel)
    
    ' If the line id was not valid or was missing, set focus
    ' back to the line id and exit the function.
    If blnCancel Then
        cboLine.SetFocus
        GoTo PROC_EXIT
    End If
       
    ' Validate the location field.
    Call cboLocation_Validate(blnCancel)
    
    ' If the location was not valid or was missing, set focus
    ' back to the location and exit the function.
    If blnCancel Then
        cboLocation.SetFocus
        GoTo PROC_EXIT
    End If
    
    'Validate the Notes
    Call txtNotes_Validate(blnCancel)
    
    ' If the Notes were not valid or was missing, set focus
    ' back to the Notes and exit the function.
    If blnCancel Then
        txtNotes.SetFocus
        GoTo PROC_EXIT
    End If
    
    'Compare the Model, Line, Stocking Location ID on form to a data base record
    'to prevent duplicates.
    If mrsDatabase.EditMode = adEditAdd Then
        Set rsDuplicate = mrsDatabase.Clone

        With rsDuplicate
            .Find ("model_number = '" & cboModel.Text & "'")

            Do While Not .EOF
                If Trim(!model_number) <> cboModel.Text Then
                    Exit Do
                End If

                If !line_id = cboLine.Text And _
                   !stocking_location_id = strLocationId Then
                    blnCancel = True
                    MsgBox "Model/Line/Location already exists "
                    Exit Do
                End If
                .MoveNext
            Loop
            .Close
        End With

        Set rsDuplicate = Nothing
        
        If Not blnCancel Then
            mrsDatabase!line_id = cboLine.Text
            mrsDatabase!model_number = cboModel.Text
            mrsDatabase!stocking_location_id = marrstrLocation(cboLocation.ItemData(cboLocation.ListIndex))
            mrsDatabase!notes = txtNotes.Text
            mrsDatabase!last_updated_date = Now()
        End If
        
    End If

    If mblnRecChanged Then
        mrsDatabase!notes = txtNotes.Text
        mrsDatabase!last_updated_date = Now()
    End If

    If blnCancel Then
        GoTo PROC_EXIT
    End If
        
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

Sub RetrieveLineData(strModelNumber As String)
    ' Purpose:  Fill the model combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim strLine As String
    
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
    
    cboLine.Clear
    
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select distinct line_id from v_mnb_model_line " & _
            "where model_number = '" & strModelNumber & "' " & _
            "union select distinct line_id From V_MNB_Model_Line_StockLoc_Notes " & _
            "where model_number = '" & strModelNumber & "' " & _
            "order by line_id asc"
        .Open
        
        Do While Not .EOF
            cboLine.AddItem Trim(!line_id)
            .MoveNext
        Loop
        .Close
    End With
      
    Set rsList = Nothing
    
    If mrsDatabase.EditMode = adEditAdd Then
        cboLine.Enabled = True
    End If
        
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveLineData", _
        Err.Number, Err.Description)
End Sub

Private Sub RetrieveLocationData(strModelNumber As String, strLineId As String)
    ' Purpose:  Fill the location combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Dim intIndex As Integer
    Set rsList = New ADODB.Recordset
    
    cboLocation.Clear
    
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "Select distinct stloc.stocking_location_id, " & _
             "a.stocking_location_description from v_prod_stocking_location a " & _
             "join v_prod_line_stocking_location stloc on a.stocking_location_id = stloc.stocking_location_id " & _
             "join V_mnb_model_line b on stloc.line_id = b.line_id " & _
             "where model_number = '" & strModelNumber & "' and stloc.line_id = '" & _
             strLineId & "' Union Select distinct stloc.stocking_location_id, a.stocking_location_description " & _
             "from v_prod_stocking_location a join v_prod_line_stocking_location stloc on " & _
             "a.stocking_location_id = stloc.stocking_location_id join V_MNB_Model_Line_StockLoc_Notes b on " & _
             "stloc.line_id = b.line_id where model_number = '" & strModelNumber & "' and stloc.line_id = '" & _
             strLineId & "' order by a.stocking_location_description asc"
        .Open
        
        If .RecordCount > 0 Then
            ReDim marrstrLocation(.RecordCount - 1)
        End If

        intIndex = 0
        Do While Not .EOF
            cboLocation.AddItem Trim(!stocking_location_description)
            marrstrLocation(intIndex) = Trim(!stocking_location_id)
            cboLocation.ItemData(cboLocation.NewIndex) = intIndex
            intIndex = intIndex + 1
            .MoveNext
        Loop
        .Close
    End With
    
    Set rsList = Nothing
    
    If mrsDatabase.EditMode = adEditAdd Then
        cboLocation.Enabled = True
    End If
        
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveLocationData", _
        Err.Number, Err.Description)
End Sub

Sub RetrieveModelData()
    ' Purpose:  Fill the model combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
       
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select distinct model_number from v_mnb_model_line " & _
            "order by model_number asc"
        .Open
        
        cboModel.Clear
        
        Do While Not .EOF
            cboModel.AddItem Trim(!model_number)
            .MoveNext
        Loop
        .Close
    End With
    
    Set rsList = Nothing
    
    
        
PROC_EXIT:
    
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveModelData", _
        Err.Number, Err.Description)
End Sub
Private Sub txtNotes_Change()
    mblnRecChanged = True
End Sub

Private Sub txtNotes_Validate(Cancel As Boolean)
    ' Purpose:  Validate that the Notes are valid length.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
                
    ' Set a variable to hold a copy of the recordset to be used
    ' to check for a duplicate.
    Dim rsModel As ADODB.Recordset
    
    ' If the length of the field is zero, give an error
    If Len(txtNotes.Text) = 0 Then
        Cancel = True
        MsgBox "Notes are required", _
            vbExclamation + vbOKOnly, _
            "Notes Validation"
        GoTo PROC_EXIT
    End If
    
PROC_EXIT:

    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "txtNotes_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

